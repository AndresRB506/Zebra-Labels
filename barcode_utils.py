import os
import tempfile
import barcode
from barcode.writer import ImageWriter


def create_barcode_temp(value: str) -> str:
    """
    Generate a temporary Code128 barcode image.

    - The barcode is generated as a PNG image
    - The image is stored in a temporary file
    - The file is NOT permanent and should be deleted after use
    - No human-readable text is printed under the barcode

    Parameters:
        value (str): The string to encode into the barcode

    Returns:
        str: Path to the temporary PNG file
    """

    # Create an ImageWriter so we can control barcode appearance
    writer = ImageWriter()

    # Barcode appearance configuration
    writer.set_options({
        "module_width": 0.25,     # Width of a single barcode bar
        "module_height": 12.0,    # Height of barcode bars
        "quiet_zone": 1.5,        # Left/right margin
        "font_size": 0,           # Disable text under barcode
        "write_text": False,      # Do not print barcode value as text
        "dpi": 300                # High resolution for printing
    })

    # Create Code128 barcode object
    code128 = barcode.get("code128", value, writer=writer)

    # Create a temporary file (system-managed location)
    fd, path = tempfile.mkstemp(suffix=".png")
    os.close(fd)  # Close file descriptor (barcode library will write)

    # Save barcode image (python-barcode auto-adds extension)
    code128.save(path.replace(".png", ""))

    # Return full path to the temporary PNG file
    return path
