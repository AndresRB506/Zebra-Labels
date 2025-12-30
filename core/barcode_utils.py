from __future__ import annotations

import os
import tempfile

import barcode
from barcode.writer import ImageWriter

def create_barcode_temp(value: str) -> str:
    """
    Creates a Code128 barcode PNG in a temp location and returns the full path.
    """
    if not value:
        raise ValueError("Barcode value is empty.")

    writer = ImageWriter()
    writer.set_options({
        "module_width": 0.2,
        "module_height": 8.0,
        "quiet_zone": 1.0,
        "font_size": 10,
        "text_distance": 1.0,
        "write_text": False,  # human-readable text off; enable if you want
    })

    code128 = barcode.get("code128", value, writer=writer)

    fd, path = tempfile.mkstemp(suffix=".png")
    os.close(fd)

    # python-barcode saves adding extension by itself if you pass path without extension
    base = path[:-4]  # remove ".png"
    code128.save(base)
    return path
