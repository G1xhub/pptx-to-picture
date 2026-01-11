# Converter Suite Pro

**Version:** 0.1.3  
**Developer:** graeLabs

Universal file converter supporting documents, images, video, audio, and presentations with a modern, tech-focused interface.

![Converter Suite UI](https://placehold.co/800x500?text=Converter+Suite+Pro+Preview)

## Features

- **Cross-Platform**: Runs on Windows, macOS, and Linux.
- **Techy UI**: Modern "Core Converter" interface with dark mode and monospace typography.
- **50+ Format Conversions**:
    - **Documents**: DOCX, PDF, TXT, MD, HTML, RTF
    - **Images**: PNG, JPG, WebP, BMP, GIF, TIFF
    - **Video**: MP4, MKV, AVI, MOV, WebM, FLV
    - **Audio**: MP3, WAV, FLAC, OGG, M4A
    - **Presentations**: PPTX, PPT, ODP -> PDF, Images
- **Modular Architecture**: Backend support for industry-standard tools (FFmpeg, LibreOffice, Pandoc).
- **Drag & Drop**: Easy batch processing with drag-and-drop support.

## Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/graeLabs/converter-suite.git
   cd converter-suite
   ```

2. **Install Python dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **External Dependencies:**
   For full functionality, ensure the following are installed (or placed in a `./deps` folder):
   - **FFmpeg**: Required for Video/Audio conversion.
   - **LibreOffice**: Required for Presentation/Document conversion.
   - **Pandoc**: Required for advanced Document formats.

## Usage

Run the application:
```bash
python main.py
```

## Supported Conversions

| Category | Input Formats | Export Formats | Backend |
|----------|---------------|----------------|---------|
| **Documents** | DOCX, ODT, TXT, MD, HTML, RTF | PDF, DOCX, MD, HTML, TXT | Pandoc / LibreOffice |
| **Images** | PNG, JPG, BMP, WebP, TIFF | PNG, JPG, WebP, GIF, TIFF, ICO | Pillow |
| **Video** | MP4, MKV, AVI, MOV, WebM | MP4, WebM, MKV, AVI, GIF | FFmpeg |
| **Audio** | MP3, WAV, FLAC, OGG, M4A | MP3, WAV, FLAC, OGG, M4A | FFmpeg |
| **Presentations** | PPTX, PPT, ODP | PDF, PNG, JPG | LibreOffice |

## License

MIT License - Copyright © 2026 graeLabs
