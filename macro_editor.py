from __future__ import annotations
from typing import TYPE_CHECKING, Optional

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QLineEdit,
    QListWidget, QComboBox, QMessageBox, QAbstractItemView
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QShortcut, QKeySequence

from actions import list_actions, BaseAction
from macro import Macro, MacroStep, MacroManager

if TYPE_CHECKING:
    from Skener import ScanApp


_SHORTCUTS = [""] + [f"F{i}" for i in range(2, 13)]


class MacroEditorDialog(QDialog):
    def __init__(
        self,
        app: ScanApp,
        manager: MacroManager,
        macro: Optional[Macro] = None,
    ) -> None:
        super().__init__(app)
        self.setWindowTitle("Editor maker")
        self.setMinimumSize(650, 500)

        self.app = app
        self.manager = manager
        self.macro = macro or Macro(name="Nové makro")
        self._modified = False

        self._build_ui()
        self._populate_from_macro()
        self._setup_shortcuts()

        if macro:
            self.app.speak(f"Editor maker: upravujete makro {macro.name}")
        else:
            self.app.speak("Editor maker: vytváříte nové makro")

    # ---------- UI ----------
    def _build_ui(self) -> None:
        layout = QVBoxLayout()

        # Název
        name_layout = QHBoxLayout()
        name_layout.addWidget(QLabel("Název makra:"))
        self.name_edit = QLineEdit()
        self.name_edit.setAccessibleName("Název makra")
        name_layout.addWidget(self.name_edit)
        layout.addLayout(name_layout)

        # Zkratka
        sc_layout = QHBoxLayout()
        sc_layout.addWidget(QLabel("Klávesová zkratka:"))
        self.shortcut_combo = QComboBox()
        self.shortcut_combo.setAccessibleName("Klávesová zkratka")
        self.shortcut_combo.addItems(_SHORTCUTS)
        sc_layout.addWidget(self.shortcut_combo)
        layout.addLayout(sc_layout)

        # Seznamy
        lists_layout = QHBoxLayout()

        # Dostupné akce
        left_layout = QVBoxLayout()
        left_layout.addWidget(QLabel("Dostupné akce:"))
        self.available_list = QListWidget()
        self.available_list.setAccessibleName("Seznam dostupných akcí")
        for action in list_actions():
            self.available_list.addItem(f"{action.name}  ({action.description})")
            item = self.available_list.item(self.available_list.count() - 1)
            item.setData(Qt.ItemDataRole.UserRole, action.id)
        left_layout.addWidget(self.available_list)
        lists_layout.addLayout(left_layout)

        # Tlačítka mezi seznamy
        mid_layout = QVBoxLayout()
        mid_layout.addStretch()
        self.btn_add = QPushButton("Přidat krok")
        self.btn_add.setAccessibleName("Přidat vybranou akci jako krok makra")
        self.btn_add.clicked.connect(self._add_step)
        mid_layout.addWidget(self.btn_add)

        self.btn_remove = QPushButton("Odebrat krok")
        self.btn_remove.setAccessibleName("Odebrat vybraný krok z makra")
        self.btn_remove.clicked.connect(self._remove_step)
        mid_layout.addWidget(self.btn_remove)

        self.btn_up = QPushButton("Posunout nahoru")
        self.btn_up.setAccessibleName("Posunout vybraný krok nahoru")
        self.btn_up.clicked.connect(self._move_up)
        mid_layout.addWidget(self.btn_up)

        self.btn_down = QPushButton("Posunout dolů")
        self.btn_down.setAccessibleName("Posunout vybraný krok dolů")
        self.btn_down.clicked.connect(self._move_down)
        mid_layout.addWidget(self.btn_down)
        mid_layout.addStretch()
        lists_layout.addLayout(mid_layout)

        # Kroky makra
        right_layout = QVBoxLayout()
        right_layout.addWidget(QLabel("Kroky makra:"))
        self.steps_list = QListWidget()
        self.steps_list.setAccessibleName("Seznam kroků makra")
        right_layout.addWidget(self.steps_list)
        lists_layout.addLayout(right_layout)

        layout.addLayout(lists_layout)

        # Akční tlačítka
        btn_layout = QHBoxLayout()
        self.btn_save = QPushButton("Uložit makro")
        self.btn_save.setAccessibleName("Uložit makro")
        self.btn_save.setDefault(True)
        self.btn_save.clicked.connect(self._save)
        btn_layout.addWidget(self.btn_save)

        self.btn_save_run = QPushButton("Uložit a spustit")
        self.btn_save_run.setAccessibleName("Uložit makro a rovnou spustit")
        self.btn_save_run.clicked.connect(self._save_and_run)
        btn_layout.addWidget(self.btn_save_run)

        self.btn_cancel = QPushButton("Zrušit")
        self.btn_cancel.setAccessibleName("Zrušit úpravy")
        self.btn_cancel.clicked.connect(self.reject)
        btn_layout.addWidget(self.btn_cancel)

        layout.addLayout(btn_layout)
        self.setLayout(layout)

    def _setup_shortcuts(self) -> None:
        QShortcut(QKeySequence("Insert"), self).activated.connect(self._add_step)
        QShortcut(QKeySequence("Delete"), self).activated.connect(self._remove_step)
        QShortcut(QKeySequence("Ctrl+Up"), self).activated.connect(self._move_up)
        QShortcut(QKeySequence("Ctrl+Down"), self).activated.connect(self._move_down)
        QShortcut(QKeySequence("Ctrl+S"), self).activated.connect(self._save)
        QShortcut(QKeySequence("Escape"), self).activated.connect(self.reject)

    # ---------- Populate ----------
    def _populate_from_macro(self) -> None:
        self.name_edit.setText(self.macro.name)
        idx = _SHORTCUTS.index(self.macro.shortcut) if self.macro.shortcut in _SHORTCUTS else 0
        self.shortcut_combo.setCurrentIndex(idx)
        self._refresh_steps()

    def _refresh_steps(self) -> None:
        self.steps_list.clear()
        for i, step in enumerate(self.macro.steps):
            action = self._find_action(step.action_id)
            if action:
                self.steps_list.addItem(f"{i + 1}. {action.name}")
            else:
                self.steps_list.addItem(f"{i + 1}. {step.action_id} (???)")

    @staticmethod
    def _find_action(action_id: str) -> Optional[BaseAction]:
        for a in list_actions():
            if a.id == action_id:
                return a
        return None

    # ---------- Modify steps ----------
    def _add_step(self) -> None:
        row = self.available_list.currentRow()
        if row < 0:
            self.app.speak("Nejdříve vyberte akci z levého seznamu.")
            return
        action_id = self.available_list.item(row).data(Qt.ItemDataRole.UserRole)
        self.macro.steps.append(MacroStep(action_id=action_id))
        self._refresh_steps()
        self.steps_list.setCurrentRow(self.steps_list.count() - 1)
        action = self._find_action(action_id)
        self.app.speak(f"Přidán krok: {action.name if action else action_id}")
        self._modified = True

    def _remove_step(self) -> None:
        row = self.steps_list.currentRow()
        if row < 0:
            self.app.speak("Nejdříve vyberte krok k odebrání.")
            return
        self.macro.steps.pop(row)
        self._refresh_steps()
        remaining = len(self.macro.steps)
        if remaining > 0:
            self.steps_list.setCurrentRow(min(row, remaining - 1))
            self.steps_list.setFocus()
        self.app.speak(f"Krok odebrán. Zbývá {remaining} kroků.")
        self._modified = True

    def _move_up(self) -> None:
        row = self.steps_list.currentRow()
        if row <= 0:
            return
        self.macro.steps[row], self.macro.steps[row - 1] = (
            self.macro.steps[row - 1], self.macro.steps[row]
        )
        self._refresh_steps()
        self.steps_list.setCurrentRow(row - 1)
        self._modified = True

    def _move_down(self) -> None:
        row = self.steps_list.currentRow()
        if row < 0 or row >= len(self.macro.steps) - 1:
            return
        self.macro.steps[row], self.macro.steps[row + 1] = (
            self.macro.steps[row + 1], self.macro.steps[row]
        )
        self._refresh_steps()
        self.steps_list.setCurrentRow(row + 1)
        self._modified = True

    # ---------- Save ----------
    def _collect(self) -> Optional[Macro]:
        name = self.name_edit.text().strip()
        if not name:
            QMessageBox.warning(self, "Chyba", "Zadejte název makra.")
            self.name_edit.setFocus()
            return None
        if not self.macro.steps:
            QMessageBox.warning(self, "Chyba", "Přidejte alespoň jeden krok.")
            self.available_list.setFocus()
            return None

        self.macro.name = name
        self.macro.shortcut = self.shortcut_combo.currentText()
        return self.macro

    def _save(self) -> None:
        macro = self._collect()
        if macro is None:
            return
        self.manager.save(macro)
        self.manager.load_all()
        self.app._rebuild_macro_buttons()
        self.app.speak(f"Makro {macro.name} uloženo.")
        self.accept()

    def _save_and_run(self) -> None:
        macro = self._collect()
        if macro is None:
            return
        self.manager.save(macro)
        self.manager.load_all()
        self.app._rebuild_macro_buttons()
        self.app.speak(f"Makro {macro.name} uloženo a spouštím.")
        self.accept()
        self.app._run_macro(macro)
