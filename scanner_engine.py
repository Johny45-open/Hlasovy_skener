import subprocess
import tempfile
import os
import xml.etree.ElementTree as ET
from typing import Optional

from PIL import Image
from PyQt6.QtCore import QThread, pyqtSignal


def get_naps2_path() -> Optional[str]:
    paths = [
        r"C:\Program Files\NAPS2\NAPS2.Console.exe",
        r"C:\Program Files (x86)\NAPS2\NAPS2.Console.exe",
        os.path.join(os.environ.get("LOCALAPPDATA", ""), r"NAPS2\NAPS2.Console.exe"),
        os.path.join(os.environ.get("APPDATA", ""), r"NAPS2\NAPS2.Console.exe"),
    ]
    for p in paths:
        if os.path.exists(p):
            return p
    return None


def get_naps2_profiles() -> list[str]:
    profiles: list[str] = []
    config_dirs = [
        os.path.join(os.environ.get("APPDATA", ""), r"NAPS2"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), r"NAPS2"),
        os.path.join(os.environ.get("APPDATA", ""), r"NAPS2\v8"),
    ]

    for base_dir in config_dirs:
        xml_files = [
            os.path.join(base_dir, "profiles.xml"),
            os.path.join(base_dir, "profiles", "profiles.xml"),
            os.path.join(base_dir, "v8", "profiles.xml"),
        ]
        for config_path in xml_files:
            if os.path.exists(config_path):
                try:
                    tree = ET.parse(config_path)
                    root = tree.getroot()
                    found = root.findall(".//Profile") or root.findall(".//ScanningProfile")
                    for profile in found:
                        name_elem = profile.find("Name")
                        if name_elem is not None and name_elem.text:
                            profiles.append(name_elem.text.strip())
                except Exception as e:
                    print(f"Chyba při čtení XML ({config_path}): {e}")

    return sorted(set(profiles))


class NAPS2Scanner:
    def __init__(self) -> None:
        self.naps2_exe = get_naps2_path()
        self.selected_profile: Optional[str] = None

    def list_devices(self) -> list[str]:
        profiles = get_naps2_profiles()
        if not profiles:
            print("Diagnostika: Žádné profily nenalezeny v XML, zkouším --listdevices.")
            try:
                if self.naps2_exe:
                    result = subprocess.run(
                        [self.naps2_exe, "--driver", "wia", "--listdevices"],
                        capture_output=True, text=True, timeout=10,
                    )
                    if result.returncode == 0:
                        for line in result.stdout.splitlines():
                            line = line.strip()
                            if line and not line.startswith("-") and \
                               not line.startswith("NAPS2") and \
                               "available" not in line.lower():
                                profiles.append(line)
            except Exception as e:
                print(f"Diagnostika: --listdevices selhalo: {e}")

        if not profiles:
            profiles = ["Výchozí skener (NAPS2 Default)"]

        print(f"Diagnostika: Konečný seznam pro GUI: {profiles}")
        return profiles

    def connect_device(self, profile_or_device: str) -> None:
        self.selected_profile = profile_or_device

    def scan(
        self,
        dpi: int = 300,
        color_mode: str = "Color",
        source: str = "Glass",
    ) -> Image.Image:
        if not self.naps2_exe:
            raise RuntimeError(
                "NAPS2.Console.exe nebyl nalezen. Nainstalujte prosím NAPS2."
            )

        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            tmp_path = tmp.name

        try:
            cmd = [self.naps2_exe, "--driver", "wia", "-o", tmp_path, "--force", "--progress"]
            cmd.extend(["--dpi", str(dpi)])
            cmd.extend(["--source", "glass" if source == "Sklo" else "feeder"])

            bitdepth = "color"
            if "Šedý" in color_mode or "Grayscale" in color_mode:
                bitdepth = "gray"
            elif "ČB" in color_mode or "BlackWhite" in color_mode:
                bitdepth = "bw"
            cmd.extend(["--bitdepth", bitdepth])

            if self.selected_profile and "Default" not in self.selected_profile:
                all_profiles = get_naps2_profiles()
                if self.selected_profile in all_profiles:
                    cmd.extend(["--profile", self.selected_profile])
                else:
                    cmd.extend(["--device", self.selected_profile])
            else:
                cmd.append("--noprofile")

            print(f"Spouštím sken (NAPS2 v8): {' '.join(cmd)}")
            result = subprocess.run(cmd, capture_output=True, text=True)

            if result.returncode != 0:
                error_msg = result.stderr.strip() or result.stdout.strip()
                raise RuntimeError(f"NAPS2 chyba ({result.returncode}): {error_msg}")

            if os.path.exists(tmp_path) and os.path.getsize(tmp_path) > 0:
                with Image.open(tmp_path) as img:
                    return img.copy()
            else:
                raise RuntimeError(
                    "Skenování neprodukovalo žádný soubor. Zkontrolujte připojení skeneru."
                )
        finally:
            if os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except OSError:
                    pass


class ScanThread(QThread):
    finished = pyqtSignal(object)
    error = pyqtSignal(str)

    def __init__(
        self,
        scanner: NAPS2Scanner,
        dpi: int,
        color_mode: str,
        source: str,
    ) -> None:
        super().__init__()
        self.scanner = scanner
        self.dpi = dpi
        self.color_mode = color_mode
        self.source = source

    def run(self) -> None:
        try:
            img = self.scanner.scan(self.dpi, self.color_mode, self.source)
            self.finished.emit(img)
        except Exception as e:
            self.error.emit(str(e))
