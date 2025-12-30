from __future__ import annotations

import os
import main as console_main  # <-- this imports your ORIGINAL console main.py at project root

def generate_doc_with_gui_color(lot_number: str, work_orders: list[dict], sheets: list[dict], color: str) -> str:
    """
    Calls the ORIGINAL generate_doc() without changing any DOCX layout logic.
    We override input_choice() at runtime so it returns the GUI-selected color.
    """

    # Keep a reference to the original function
    original_input_choice = console_main.input_choice

    try:
        # Override input_choice so generate_doc() doesn't ask the terminal
        def _input_choice_override(prompt: str, choices: list[str]) -> str:
            return color.strip().upper()

        console_main.input_choice = _input_choice_override

        # Ensure output folder exists (original code saves to output/...)
        os.makedirs("output", exist_ok=True)

        # Call your ORIGINAL generator (layout untouched)
        return console_main.generate_doc(lot_number, work_orders, sheets)

    finally:
        # Restore it back
        console_main.input_choice = original_input_choice
