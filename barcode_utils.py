import os
from barcode import Code128
from barcode.writer import ImageWriter


def create_barcode(value, filename):
    png_file = filename + ".png"

    # Borrar barcode viejo si existe
    if os.path.exists(png_file):
        os.remove(png_file)

    writer = ImageWriter()
    writer.set_options({
        "write_text": False,     # SIN texto debajo
        "module_height": 8.0,    # Altura del barcode
        "module_width": 0.3,     # Grosor de barras
        "quiet_zone": 2.0
    })

    barcode = Code128(value, writer=writer)
    barcode.save(filename)
