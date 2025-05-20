
# EOL CPU Reporting

This Excel VBA tool highlights End-of-Life (EOL) CPUs in a system inventory report and applies conditional formatting based on system attributes like OS version, RAM, storage, and virtualization status.

## 🚀 Features

- Automatically formats raw data into an Excel table
- Highlights EOL CPUs based on an external list
- Flags systems with:
  - Low RAM (< 16 GB)
  - Low SSD space (< 25% free)
  - Specific OS versions (Windows 10/11 variants)
  - Virtual machines (VMware)
- Color-coded for quick visual analysis

## 📁 Requirements

- Microsoft Excel (with macros enabled)
- A reference file: `EOL_CPU_List.xlsx` (placed in your Downloads folder or selected manually)

## 🛠️ Setup

1. Download or clone this repository.
2. Open the Excel workbook.
3. Press `Alt + F11` to open the VBA editor.
4. Paste the contents of `HighlightEOLCPU.bas` into a new module.
5. Save and close the editor.

## ▶️ Usage

1. Ensure your data is on a sheet named `Table`.
2. Run the macro `HighlightEOLCPUs`.
3. If the EOL list is not found in your Downloads folder, you will be prompted to select it manually.
4. The macro will:
   - Format the data as a table
   - Normalize numeric columns
   - Highlight rows based on CPU status and system attributes

## ✅ TODO List

- [x] Highlight EOL CPUs in red.
- [x] Highlight Servers in blue.
- [x] Convert memory and storage columns to numeric values.
- [x] Highlight **RAM upgrade** if `Agent Memory Total` < 16GB (purple cell).
- [x] Highlight **SSD upgrade** if `C Drive Free Percent` < 25% (light blue cells).
- [x] Highlight **OS versions**:
  - Win11 Pro → green
  - Win10 Pro → yellow
  - Win10/11 Home → amber
- [x] Highlight **VMware systems** in brown.
- [ ] Label **Needs Pro** if Windows version is Home.
- [ ] Label **Upgradeable to Win11 Pro** if on Win10 Pro and no issues.
- [ ] Label **Already on Win11 Pro** if no issues and already on Win11 Pro.

## 📊 Future Enhancements

### Code Enhancements
- Modularize the rest of the logic
- Add unit tests for each subroutine
- Add logging (e.g., write actions to a hidden sheet or log file)

### User Experience
- Add a dashboard sheet with summary stats (e.g., % of EOL systems, upgrade candidates)


## 🛡️ License

This work is licensed under the Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0).  
You are free to share and adapt the material, but not for commercial purposes. You must give appropriate credit.

## 🚫 Non-Distribution Notice

This project is proprietary and should not be redistributed without explicit permission from the author.  
Unauthorized distribution or monetization is strictly prohibited.
