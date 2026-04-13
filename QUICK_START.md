# 🎯 GST Automation - Quick Start Guide

## ✅ What Changed

### 1. **Automatic Process Chaining**
- After scraping completes successfully, you'll be asked if you want to process the files
- If you click "Yes", it automatically runs the file processing
- If you click "No", you can run `main_organizar.py` manually later

### 2. **Bundled Executable Support**
- Added `get_playwright_browser_path()` function to handle .exe bundling
- When running as .exe, it uses the bundled Chromium browser
- When running as Python script, it uses installed Playwright browser

### 3. **Relative Path Handling**
- All paths now use `BASE_DIR` for portability
- Works correctly whether running from different locations
- Compatible with PyInstaller bundling

## 🚀 How to Use

### Running the Script (Development)

```bash
# Install dependencies first
pip install -r requirements.txt

# Run the main script
python main.py
```

**Workflow:**
1. Click "Yes" to start extraction
2. Browser opens, logs in, downloads 4 files
3. Success message appears
4. You're asked: "Deseja processar os arquivos baixados agora?"
   - **Yes** → Automatically processes files (creates formatted Excel files)
   - **No** → Exit (you can run `python main_organizar.py` later)

### Building the Executable

```bash
# Easy way - just run the batch script
build_exe.bat

# Or manually
 pyinstaller --noconfirm --onefile --windowed --noconsole --name "GST Report Extraction" --icon "C:/Users/perna/Desktop/STALLANTIS/Extra-o-de-relatorio-no-GST/Credencial/icon.ico" --add-data "C:\Users\perna\AppData\Local\ms-playwright\chromium-1187\chrome-win;ms-playwright\chromium-1187\chrome-win" main.py
```

**Output:** `dist\GST Report Extraction.exe`

### Distributing to Other Machines

Copy these files:
```
GST Report Extraction.exe     ← The executable
Credencial/
  └── usuario.json            ← Login credentials
  └── icon.ico
```

The exe will automatically create these folders when it runs:
- `Downloads_Auxiliar/` (for raw downloads)
- `Arquivos_Consolidados/` (for processed files)

## 📂 Project Structure

```
Extra-o-de-relatorio-no-GST/
│
├── main.py                    ← Main script (extraction + processing)
├── main_organizar.py          ← Standalone processing script (if needed)
├── requirements.txt           ← Python dependencies
│
├── build_exe.bat              ← Build script (creates .exe)
├── BUILD_INSTRUCTIONS.md      ← Detailed build guide
├── QUICK_START.md             ← This file
│
├── Credencial/
│   ├── usuario.json          ← Login credentials
│   └── icon.ico              ← Application icon
│
├── Downloads_Auxiliar/        ← Raw downloaded files
│   ├── 01_source_package_management.xlsx
│   ├── 02_sourcing_management_global.xlsx
│   ├── 03_sourcing_management_latam.xlsx
│   └── 04_sourcing_management_neutral.xlsx
│
└── Arquivos_Consolidados/     ← Processed output files
    ├── Source_Package_Management.xlsx
    └── Sourcing_Management.xlsx
```

## ⚡ Key Features

### Extraction Module
- ✅ Automated login to GSt portal
- ✅ Downloads 4 Excel files automatically
- ✅ Handles iframe navigation
- ✅ Works with readonly fields
- ✅ 5-minute download timeout per file

### Processing Module
- ✅ Merges multiple files
- ✅ Filters by status and type
- ✅ Professional Excel formatting
- ✅ Automatic column width optimization
- ✅ Black header with white text styling

### Bundled .exe
- ✅ Single executable file
- ✅ No Python installation needed
- ✅ Chromium browser bundled
- ✅ Works on any Windows 10/11 machine

## 🔧 Troubleshooting

### Issue: "Chromium executable not found"
Check the Playwright version in your AppData:
```
C:\Users\perna\AppData\Local\ms-playwright\
```

If you see `chromium-1188` instead of `chromium-1187`, update in:
1. [main.py](main.py#L156) → `get_playwright_browser_path()` function
2. [build_exe.bat](build_exe.bat#L15) → `--add-data` parameter

### Issue: Processing fails
Make sure you have all dependencies:
```bash
pip install -r requirements.txt
```

### Issue: .exe doesn't work on other machines
Make sure you copied the `Credencial` folder with `usuario.json`

## 📝 Configuration

### Update Login Credentials

Edit `Credencial/usuario.json`:
```json
{
  "Usuario": "YOUR_USERNAME",
  "Senha": "YOUR_PASSWORD"
}
```

### Change Filters

Edit [main.py](main.py#L227) in the `process_downloaded_files()` function:
```python
# Current filter
filtro = (df['Status'] == 'Technical Data Completed') & (df['Tipo'] != 'TAG')

# Example: multiple statuses
filtro = df['Status'].isin(['Status1', 'Status2'])
```

## 📚 Additional Resources

- [BUILD_INSTRUCTIONS.md](BUILD_INSTRUCTIONS.md) - Detailed .exe creation guide
- [README.md](README.md) - Complete project documentation
- [main_organizar.py](main_organizar.py) - Standalone processing script

## 💡 Tips

1. **First run?** Test the Python script first before building .exe
2. **Building .exe?** Close all Excel files in the project folder
3. **Distributing?** Test on a clean VM before deploying
4. **Network issues?** Increase timeout values in code

---

**Need Help?**  
Check [BUILD_INSTRUCTIONS.md](BUILD_INSTRUCTIONS.md) for detailed troubleshooting

**Last Updated**: April 13, 2026  
**Version**: 2.1 (with chained processing)
