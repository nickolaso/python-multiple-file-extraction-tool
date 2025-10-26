# python-multi-extraction-bot

Python Multiple File Extraction Tool
Extracts all archives (zip, rar, 7z, tar*) in a selected folder into a single "unarchived" directory.

----

## Setup

```
git@github.com:nickolaso/python-multiple-file-extraction-tool.git
cd python-multiple-file-extraction-tool
pip install py7zr rarfile
python ./extraction-tool.py
```
---
### It runs .zip files without any extra setup.  But if you run into trouble check the section that applies to your operating system.
---

Minimal installs to make RAR/7Z work everywhere

### Windows:
```
winget install 7zip.7zip (or choco install 7zip)
```
If 7z isn’t on PATH, set:
```
SEVENZ_PATH_OVERRIDE = r"C:\Program Files\7-Zip\7z.exe"
```

### macOS:
```
brew install p7zip (adds 7z)
```
Optional RAR helpers: brew install unar or brew install unrar

### Ubuntu/Debian:
```
sudo apt install p7zip-full
```
Optional RAR helpers: sudo apt install unrar or sudo apt install unar

Pure-Python fallbacks (optional):
```
pip install py7zr rarfile
```