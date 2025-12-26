import tempfile
import barcode
from barcode.writer import ImageWriter

def create_barcode_temp(value):
    code = barcode.get("code128", value, writer=ImageWriter())
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    code.write(tmp)
    tmp.close()
    return tmp.name
