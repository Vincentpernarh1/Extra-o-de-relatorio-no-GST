# 🔨 Build Instructions for GST Report Extraction

## Prerequisites

1. **Install PyInstaller** (if not already installed):
   ```bash
   pip install pyinstaller
   ```

2. **Install Playwright browsers**:
   ```bash
   playwright install chromium
   ```

3. **Verify Playwright path exists**:
   - Check that `C:\Users\perna\AppData\Local\ms-playwright\chromium-1187\chrome-win\chrome.exe` exists
   - If the version number is different (e.g., chromium-1188), update the path in both:
     - `main.py` → `get_playwright_browser_path()` function
     - `build_exe.bat` → `--add-data` parameter

## Build Methods

### Method 1: Using the Batch Script (Recommended)

Simply double-click or run:
```bash
build_exe.bat
```

### Method 2: Manual PyInstaller Command

Run this command from the project directory:

```bash
 pyinstaller --noconfirm --onefile --windowed --noconsole --name "GST Report Extraction" --icon "C:/Users/perna/Desktop/STALLANTIS/Extra-o-de-relatorio-no-GST/Credencial/icon.ico" --add-data "C:\Users\perna\AppData\Local\ms-playwright\chromium-1187\chrome-win;ms-playwright\chromium-1187\chrome-win" main.py
```

## Build Options Explained

| Option | Purpose |
|--------|---------|
| `--noconfirm` | Overwrite existing build without asking |
| `--onefile` | Bundle everything into a single .exe file |
| `--windowed` | No console window (GUI only) |
| `--noconsole` | Same as windowed (redundant but explicit) |
| `--name "GST Report Extraction"` | Name of the output executable |
| `--icon "Credencial\icon.ico"` | Custom icon for the .exe |
| `--add-data "source;dest"` | Bundle Playwright browser with the .exe |

## After Building

### Output Location
The executable will be created in:
```
dist\GST Report Extraction.exe
```

### Required Files for Distribution
When distributing the .exe to other machines, include these folders in the same directory:
```
GST Report Extraction.exe
Credencial/
  └── usuario.json     (User credentials)
  └── icon.ico
Downloads_Auxiliar/    (Created automatically if missing)
Arquivos_Consolidados/ (Created automatically if missing)
```

### First Run on Target Machine
1. Copy the .exe and `Credencial` folder to the target machine
2. The application will automatically create `Downloads_Auxiliar` and `Arquivos_Consolidados` folders
3. Ensure `usuario.json` has valid credentials

## Troubleshooting

### Issue: "Chromium executable not found"
**Solution**: 
- Verify the Playwright browser version matches the path in code
- Check: `%LOCALAPPDATA%\ms-playwright\`
- Update version number in `get_playwright_browser_path()` if different

### Issue: Build fails with module errors
**Solution**:
```bash
pip install -r requirements.txt
```

Create `requirements.txt` with:
```
playwright>=1.40.0
pandas>=1.3.0
xlsxwriter>=3.0.0
openpyxl>=3.0.0
```

### Issue: .exe runs but browser doesn't open
**Solution**:
- The bundled Chromium may not have loaded correctly
- Check if the `--add-data` path in build command is correct
- Verify the source path exists before building

### Issue: .exe is too large (>200MB)
**Expected**: This is normal when bundling Chromium browser (~150-200MB)

## Testing the Build

1. **Test locally first**:
   ```bash
   dist\"GST Report Extraction.exe"
   ```

2. **Test on a clean machine** (VM recommended):
   - Copy .exe + Credencial folder
   - Run without Python installed
   - Verify all features work

## Updating the Build

When code changes:
1. Update `main.py` with your changes
2. Re-run `build_exe.bat`
3. Old build in `dist/` will be replaced
4. Test the new .exe before distribution

## Clean Build

To start fresh:
```bash
rmdir /s /q build dist
del *.spec
build_exe.bat
```

---

**Last Updated**: $(Get-Date -Format "yyyy-MM-dd")  
**Compatible with**: Windows 10/11  
**Python Version**: 3.8+  
**Playwright Version**: 1.40+
