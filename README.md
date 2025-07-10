# PDF Toolkit

A user-friendly desktop application for PDF manipulation with an intuitive graphical interface. Built with Python and packaged using PyInstaller for cross-platform compatibility.

<img width="1606" height="1262" alt="image" src="https://github.com/user-attachments/assets/47dbfcc9-ceb8-49e3-a734-7476040df71c" />

## Features

- **Merge PDFs**: Combine multiple PDF files into a single document
- **Split PDF**: Split large PDF files into smaller sections or individual pages
- **Compress PDF**: Reduce PDF file size while maintaining quality
- **PDF to Word**: Convert PDF documents to Microsoft Word format (.docx)

## Installation

### Download Pre-built Binaries

Download the latest release for your operating system:
- **macOS**: [Download PDFToolkit.app](https://github.com/marudevansh25111/PDFToolkit/releases)
- **Windows**: [Download PDFToolkit.exe](https://github.com/marudevansh25111/PDFToolkit/releases)

### Run from Source

```bash
# Clone the repository
git clone https://github.com/marudevansh25111/PDFToolkit.git
cd PDFToolkit

# Install dependencies
pip install -r requirements.txt

# Run the application
python main.py
```

## Usage

1. **Launch the application**
   - On macOS: Double-click `PDFToolkit.app`
   - On Windows: Double-click `PDFToolkit.exe`
   - From source: Run `python main.py`

2. **Select your operation**
   - Click on the desired tab: Merge PDFs, Split PDF, Compress PDF, or PDF to Word

3. **Add files**
   - Click "Add Files" to select your PDF files
   - Files will appear in the list area

4. **Configure output**
   - Click "Select Output" to choose where to save the result
   - Set output filename and location

5. **Process**
   - Click the main action button (e.g., "Merge PDFs") to start processing

## Building from Source

### Prerequisites

```bash
pip install -r requirements.txt
pip install pyinstaller
```

### Build for macOS

```bash
# Create standalone application
pyinstaller --onefile --windowed --name "PDFToolkit" main.py

# Create .app bundle (recommended for macOS)
pyinstaller --onefile --windowed --name "PDFToolkit" --add-data "assets:assets" main.py

# The built app will be in dist/PDFToolkit.app
```

### Build for Windows

```bash
# Create standalone executable
pyinstaller --onefile --windowed --name "PDFToolkit" --icon=icon.ico main.py

# Add resources if needed
pyinstaller --onefile --windowed --name "PDFToolkit" --add-data "assets;assets" --icon=icon.ico main.py

# The built executable will be in dist/PDFToolkit.exe
```

### Build Configuration

Create a `build.spec` file for more control:

```python
# PDFToolkit.spec
import sys
from PyInstaller.building.build_main import Analysis, PYZ, EXE, BUNDLE

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('assets', 'assets')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='PDFToolkit',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico'  # Windows only
)

# macOS App Bundle
if sys.platform == 'darwin':
    app = BUNDLE(
        exe,
        name='PDFToolkit.app',
        icon='icon.icns',  # macOS icon
        bundle_identifier='com.marudevansh.pdftoolkit',
        info_plist={
            'CFBundleShortVersionString': '1.0.0',
            'CFBundleVersion': '1.0.0',
        },
    )
```

Then build with:
```bash
pyinstaller PDFToolkit.spec
```

## Requirements

- Python 3.7+
- PyPDF2 or PyPDF4
- tkinter (usually included with Python)
- Pillow (PIL)
- python-docx (for PDF to Word conversion)
- Additional dependencies listed in `requirements.txt`


### Running in Development Mode

```bash
python main.py
```

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/new-feature`)
3. Commit your changes (`git commit -am 'Add new feature'`)
4. Push to the branch (`git push origin feature/new-feature`)
5. Create a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

If you encounter any issues or have questions:
1. Check the [Issues](https://github.com/marudevansh25111/PDFToolkit/issues) page
2. Create a new issue with detailed information about your problem
3. Include your operating system and Python version

## Changelog

### v1.0.0
- Initial release
- PDF merging functionality
- PDF splitting functionality
- PDF compression
- PDF to Word conversion
- Cross-platform GUI application
- Standalone executables for macOS and Windows

---

**Author**: Devansh Maru  
**GitHub**: [@marudevansh25111](https://github.com/marudevansh25111)  
**Built with**: Python, tkinter, PyInstaller
