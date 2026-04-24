import comtypes.client
import tempfile
from PIL import Image
from PyQt6.QtCore import QThread, pyqtSignal

class WIAScanner:
    def __init__(self):
        self.dm = comtypes.client.CreateObject("WIA.DeviceManager")
        self.device = None

    def list_devices(self):
        return [self.dm.DeviceInfos.Item(i + 1) for i in range(self.dm.DeviceInfos.Count)]

    def connect_device(self, index=0):
        devices = self.list_devices()
        if not devices:
            raise RuntimeError("Žádný skener nenalezen")
        if index < 0 or index >= len(devices):
            raise RuntimeError("Neplatný index skeneru")
        self.device = devices[index].Connect()

    def scan(self, dpi=300, color_mode=1, source="Sklo"):
        if not self.device:
            raise RuntimeError("Skener není připojen")

        def set_prop(name, val):
            for p in self.device.Properties:
                if p.Name == name:
                    p.Value = val
                    return True
            return False

        set_prop("6147", dpi)
        set_prop("6148", dpi)
        source_val = 1 if source == "Sklo" else 2
        set_prop("6146", source_val)
        set_prop("6151", color_mode)

        item = self.device.Items[1]
        image = item.Transfer()

        with tempfile.NamedTemporaryFile(suffix=".bmp") as tmp:
            image.SaveFile(tmp.name)
            with Image.open(tmp.name) as img:
                pil_img = img.copy()
        return pil_img

class ScanThread(QThread):
    finished = pyqtSignal(object)
    error = pyqtSignal(str)

    def __init__(self, scanner, dpi, color_mode, source):
        super().__init__()
        self.scanner = scanner
        self.dpi = dpi
        self.color_mode = color_mode
        self.source = source

    def run(self):
        try:
            img = self.scanner.scan(dpi=self.dpi, color_mode=self.color_mode, source=self.source)
            self.finished.emit(img)
        except Exception as e:
            self.error.emit(str(e))
