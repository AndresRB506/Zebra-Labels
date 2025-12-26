import os
from barcode import Code128
from barcode.writer import ImageWriter


def create_barcode(data: str, filename: str):
    folder = os.path.dirname(filename)
    if folder:
        os.makedirs(folder, exist_ok=True)

    barcode = Code128(data, writer=ImageWriter())
    barcode.save(filename)
