from __future__ import annotations
import json
import os
from dataclasses import dataclass, field, asdict
from typing import TYPE_CHECKING, Optional

from actions import PipelineContext, get_action

if TYPE_CHECKING:
    from Skener import ScanApp


MACROS_DIR = os.path.join(
    os.environ.get("APPDATA", os.path.expanduser("~")),
    "HlasovySkener", "macros"
)


@dataclass
class MacroStep:
    action_id: str
    params: dict = field(default_factory=dict)


@dataclass
class Macro:
    name: str
    steps: list[MacroStep] = field(default_factory=list)
    shortcut: str = ""
    file_path: str = ""


def _macros_dir() -> str:
    os.makedirs(MACROS_DIR, exist_ok=True)
    return MACROS_DIR


class MacroManager:
    def __init__(self) -> None:
        self.macros: list[Macro] = []

    def load_all(self) -> list[Macro]:
        self.macros.clear()
        mdir = _macros_dir()
        if not os.path.isdir(mdir):
            return self.macros
        for fname in sorted(os.listdir(mdir)):
            if fname.lower().endswith(".json"):
                fpath = os.path.join(mdir, fname)
                macro = self._load_file(fpath)
                if macro:
                    self.macros.append(macro)
        return self.macros

    def _load_file(self, fpath: str) -> Optional[Macro]:
        try:
            with open(fpath, "r", encoding="utf-8") as f:
                data = json.load(f)
            steps = [MacroStep(**s) for s in data.get("steps", [])]
            return Macro(
                name=data.get("name", "Nepojmenované"),
                steps=steps,
                shortcut=data.get("shortcut", ""),
                file_path=fpath,
            )
        except Exception as e:
            print(f"Chyba načítání makra {fpath}: {e}")
            return None

    def save(self, macro: Macro, filename: str = "") -> None:
        mdir = _macros_dir()
        if not filename:
            filename = self._safe_filename(macro.name) + ".json"
        if not filename.endswith(".json"):
            filename += ".json"
        fpath = os.path.join(mdir, filename)
        data = {
            "name": macro.name,
            "shortcut": macro.shortcut,
            "steps": [asdict(s) for s in macro.steps],
        }
        with open(fpath, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        macro.file_path = fpath

    def delete(self, macro: Macro) -> None:
        if macro.file_path and os.path.exists(macro.file_path):
            os.remove(macro.file_path)
        if macro in self.macros:
            self.macros.remove(macro)

    @staticmethod
    def _safe_filename(name: str) -> str:
        safe = "".join(c if c.isalnum() or c in " _-" else "_" for c in name)
        return safe.strip().replace(" ", "_") or "makro"


class PipelineRunner:
    def __init__(self, app: ScanApp) -> None:
        self.app = app

    def run(self, macro: Macro) -> None:
        ctx = PipelineContext()
        for i, step in enumerate(macro.steps):
            if ctx.cancelled:
                break
            action = get_action(step.action_id)
            if action is None:
                self.app.speak(f"Krok {i + 1}: neznámá akce {step.action_id}")
                continue
            self.app.speak(f"Krok {i + 1}: {action.name}")
            action.run(ctx, self.app)
