# PPTX to Picture Converter

🎉 **Comprehensive file converter for Windows with modern UI** - converts presentations and documents to various image formats.

## Features

### 📁 Input Formats
- **PowerPoint**: .pptx, .ppt, .ppsx, .pps
- **PDF**: .pdf (requires Poppler)
- **LibreOffice**: .odp (requires LibreOffice)

### 🖼️ Output Formats
- **Images**: PNG, JPG, BMP, WebP, TIFF, GIF
- **Documents**: PDF
- **Vector**: SVG

### 🎨 Modern UI
- **Tab System**: Output, Format, Advanced, Dependencies, Recent
- **Sidebar**: Quick Actions + Theme Toggle
- **Card Design**: Modern, clean layout
- **Icon System**: Emoji icons for intuitive navigation
- **Dark/Light Theme**: Toggle between themes
- **Smooth Transitions**: Animated UI elements

### 💾 Usability Features
- **Presets System**: 4 built-in presets + custom presets
  - Web Optimized (WebP, 1500px, 80%)
  - Print Quality (PNG, 300 DPI, 100%)
  - Fast Export (JPG, 720p, 70%)
  - High Quality (PNG, 4K, 95%)
- **Recent Files**: Last 10 converted files with context menu
- **Drag & Drop**: Enhanced drop zone with animations
- **Multi-Progress**: Per-file + overall progress bars
- **Toast Notifications**: Success, Error, Warning, Info popups
- **Success Animation**: Visual feedback on completion
- **Keyboard Shortcuts**: Full shortcut support
- **Context Menus**: Right-click on recent files

### ⚙️ Advanced Options
- **Resolution Control**: HD (1920x1080), 4K (3840x2160), Custom
- **Format Options**:
  - **PDF DPI**: 72-300 (for PDF input)
  - **WebP**: Lossless, Quality (0-100%)
  - **TIFF**: Compression (None, LZW, JPEG, ZIP), Multi-page
  - **GIF**: Frame rate (1-30 FPS), Animated
  - **PDF**: Multi-page (all pages in one PDF)
- **Watermark**: Custom text, opacity (0-100%), 5 positions
- **Page Selection**: Export specific pages (e.g. "1,3,5-10")
- **Reverse Order**: Export pages in reverse
- **Auto Delete**: Remove original files after export
- **ZIP Creation**: Package exports automatically
- **Output Pattern**: Custom structure using `{date}` and `{filename}`
- **Settings Persistence**: All settings saved automatically

### ⌨️ Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| Ctrl+O | Select Files |
| Ctrl+F | Convert Folder |
| Ctrl+S | Save Settings |
| Ctrl+H | Toggle Theme (Dark/Light) |
| Ctrl+Z | Undo Settings |
| Ctrl+Y | Redo Settings |
| Esc | Cancel Conversion |
| F1 | Show Help |

### 🔧 Dependencies

**Optional but recommended for full functionality:**

#### Poppler (for PDF conversion)
- **Required**: PDF input support
- **Download**: [GitHub Releases](https://github.com/oschwartz10612/poppler-windows/releases)
- **Install to**: `C:\Program Files\poppler`
- **Auto-detect**: The app will find it automatically

#### LibreOffice (for ODP conversion)
- **Required**: ODP input support
- **Download**: [Portable Versions](https://www.libreoffice.org/download/portable-versions/)
- **Install**: Any location or portable version
- **Auto-detect**: The app will find it automatically

## Installation

### From Source

1. Clone or download this repository
2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Run the application:
```bash
python pptx_to_picture.py
```

### EXE (Standalone)

**Pre-built EXE available in `dist/` folder** - no Python installation required!

Or build your own:
```bash
python build_exe.py
```

The EXE includes all dependencies and is ready to run on any Windows machine.

## Usage

### Basic Workflow

1. **Select Output Directory**: Choose where to save converted files
2. **Choose Output Format**: PNG, JPG, BMP, WebP, TIFF, GIF, SVG, or PDF
3. **Configure Options**: Resolution, quality, format-specific settings
4. **Add Files**: 
   - Drag & Drop files onto the drop zone
   - Click "Select Files" button
   - Or click "Convert Folder" for batch processing
5. **Convert**: Files are processed automatically with progress feedback

### Presets

Use presets for quick configuration:

1. Select a preset from the "Format" tab dropdown
2. All settings are automatically applied
3. Customize further if needed
4. Or create your own custom preset!

### Recent Files

Access your recently converted files:

1. Go to "Recent" tab
2. Double-click a file to convert it again
3. Right-click for context menu:
   - Convert with current settings
   - Open in File Explorer
   - Remove from list
   - Copy filename

### Batch Conversion

Convert entire folders:

1. Click "Convert Folder" in sidebar
2. Select folder containing files
3. All supported files in the folder will be converted
4. Progress shows overall and per-file status

### Advanced Workflows

#### Custom Output Pattern
Organize exports using variables:
- `{filename}` - Original filename
- `{date}` - Current date (YYYY-MM-DD)
- Examples:
  - `{filename}` → `presentation`
  - `{date}/{filename}` → `2024-01-15/presentation`
  - `{filename}_{date}` → `presentation_2024-01-15`

#### Watermark
Add watermarks to all exports:
- Enter text in Advanced tab
- Set opacity (0-100%)
- Choose position: top-left, top-right, bottom-left, bottom-right, center

#### Page Selection
Export specific pages:
- Enter range in Advanced tab
- Examples: `1,3,5-10,15`
- Single pages: `1, 3, 5`
- Ranges: `1-10`
- Combined: `1,3,5-10,15`

## Settings

All settings are automatically saved to:
```
%USERPROFILE%\pptx_converter_settings.json
```

Settings include:
- Output directory
- Output pattern
- Format preferences
- Resolution & quality settings
- Format-specific options (WebP, TIFF, GIF, PDF)
- Watermark settings
- Dependency paths (Poppler, LibreOffice)
- Recent files (last 10)
- Theme preference
- Undo/Redo stack
- And more...

## Troubleshooting

### Application won't start
- Ensure Windows 10/11
- Try running as administrator
- Check Windows Event Viewer for errors
- Disable antivirus temporarily

### PowerPoint not working
- Ensure Microsoft PowerPoint is installed
- Close other PowerPoint instances
- Try running as administrator

### PDF conversion fails
- Install Poppler from Dependencies tab
- Or download manually from GitHub releases
- Restart the application after installation
- Verify Poppler path in settings

### ODP conversion fails
- Install LibreOffice from Dependencies tab
- Ensure LibreOffice is accessible
- Restart the application after installation
- Verify LibreOffice path in settings

### Drag & Drop not working
- Try using "Select Files" button instead
- Ensure file is a supported format
- Check file isn't corrupted

### EXE doesn't start
- Disable antivirus temporarily
- Run as administrator
- Check Windows Event Viewer for errors
- Verify all dependencies are included

### Performance issues
- Use lower resolution for faster conversion
- Disable unnecessary options (watermark, animations)
- Process files in smaller batches
- Use SSD for output directory
- Close other PowerPoint instances

## File Size Guide

| Format | Quality | Size (approx.) |
|--------|---------|----------------|
| PNG | High | Large |
| JPG | 90% | Medium |
| WebP | 80% | Small |
| TIFF | LZW | Medium-Large |
| GIF | 10fps | Small-Medium |
| PDF | 100% | Medium |

## Tips

### Best Performance
- Use HD resolution for web, 4K for print
- Lower quality for faster exports
- Batch convert folders instead of individual files
- Use SSD for output directory

### Best Quality
- PNG for lossless compression
- WebP for web optimization
- 100% quality for print
- 300 DPI for PDF input

### Workflow Optimization
- Create presets for common use cases
- Use Recent Files for frequent conversions
- Organize exports with output patterns
- Use keyboard shortcuts for speed

## File Size

### Source Code
- **Python file**: ~1068 lines
- **Dependencies**: 6 packages
- **Total**: ~2 MB (uncompressed)

### Standalone EXE
- **Size**: ~37 MB
- **Includes**: All dependencies + Python runtime
- **No dependencies**: Runs on any Windows machine

## Version History

### v2.0 (Current)
- ✅ Complete UI redesign with modern interface
- ✅ Tab system with 5 organized tabs
- ✅ Sidebar with quick actions
- ✅ Icon system with emojis
- ✅ Dark/Light theme toggle
- ✅ Preset system with 4 built-in + custom
- ✅ Recent files with context menu
- ✅ Toast notification system
- ✅ Multi-progress bars (per-file + overall)
- ✅ Success animations
- ✅ Keyboard shortcuts
- ✅ Smooth transitions
- ✅ Modern input controls
- ✅ Format-specific options panel
- ✅ PDF input support (with Poppler)
- ✅ ODP input support (with LibreOffice)
- ✅ WebP, TIFF, GIF, SVG, PDF output
- ✅ Multi-page support for TIFF, GIF, PDF
- ✅ Undo/Redo for settings
- ✅ Context menus for recent files
- ✅ Auto-detection and download of dependencies
- ✅ Batch folder conversion
- ✅ Standalone EXE (37 MB)

### v1.0
- Initial release
- Basic PowerPoint to image conversion
- PNG, JPG, BMP support
- Simple dark UI

## License

This tool is provided as-is for personal and commercial use.

## Credits

Built with:
- **Python** - Core language
- **tkinter / tkinterdnd2** - GUI framework
- **Pillow (PIL)** - Image processing
- **pdf2image** - PDF conversion
- **svgwrite** - SVG generation
- **PyInstaller** - EXE creation
- **win32com** - PowerPoint automation

## Support

For issues or feature requests:
1. Check Troubleshooting section
2. Ensure all dependencies are installed
3. Verify file formats are supported
4. Check Windows Event Viewer for errors
5. Test with fresh installation

---

**🎉 Enjoy converting your files with the new modern interface!**
