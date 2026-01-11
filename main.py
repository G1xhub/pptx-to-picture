"""
Converter Suite - Main Entry Point
Universal file converter with modern CustomTkinter UI.
"""

import sys
from pathlib import Path

# Add src to path for imports
src_path = Path(__file__).parent / "src"
sys.path.insert(0, str(src_path))

from src.ui.app import ConverterApp


def main():
    """Main entry point for the application."""
    app = ConverterApp()
    app.mainloop()


if __name__ == "__main__":
    main()
