from barcode import Code128
from barcode.writer import ImageWriter
import tempfile

def create_barcode_temp(value):
    barcode = Code128(value, writer=ImageWriter())

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    tmp.close()

    barcode.save(tmp.name.replace(".png", ""))
    return tmp.name
