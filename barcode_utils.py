import os
import tempfile
import barcode
from barcode.writer import ImageWriter


def create_barcode_temp(value: str) -> str:
    writer = ImageWriter()

    # 🔧 TAMAÑO NORMAL (equilibrado)
    writer.set_options({
        "module_width": 0.25,     # grosor normal
        "module_height": 12.0,    # altura normal
        "quiet_zone": 1.5,        # margen lateral correcto
        "font_size": 0,           # ❌ sin texto debajo
        "write_text": False,      # ❌ no escribir texto
        "dpi": 300
    })

    code128 = barcode.get("code128", value, writer=writer)

    fd, path = tempfile.mkstemp(suffix=".png")
    os.close(fd)

    code128.save(path.replace(".png", ""))
    return path
