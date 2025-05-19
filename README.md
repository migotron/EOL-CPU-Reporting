# EOL-CPU-Reporting

## 🧠 Project Overview
This project automates the process of generating End-of-Life (EOL) reports for customer devices. It streamlines the manual review of hardware and flags systems for replacement or upgrade based on CPU age, agent type, and hardware specs.

---

## 🛠️ Features Implemented

### ✅ 1. Table Formatting
- Automatically converts raw data into an Excel **table** (`ListObject`) if not already formatted.
- Applies the **"Normal"** cell style to remove legacy formatting.
- AutoFits all **columns and rows** for clean presentation.

### ✅ 2. Data Normalization
- Converts key columns to **numeric values**:
  - **Column I**: `Agent Memory Total`
  - **Column N**: `C Drive Free Percent` (handles text like `"85%"`)
  - **Column O**: `Total Internal Drive`
- Suppresses Excel's **"number stored as text"** warning.

### ✅ 3. EOL CPU Highlighting
- Compares values in **Column K** (`CPU`) against an external list of EOL CPUs.
- Highlights matching rows in **red** (`RGB(255, 0, 0)`).
- EOL CPU list is loaded from:
  - A default path in the user's **Downloads** folder.
  - Or a **file picker** if the file is not found.

### ✅ 4. Server Agent Highlighting
- Checks **Column D** (`Agent Type`) for `"Server"`.
- Highlights the row in **blue** (`RGB(0, 112, 192)`) if not already marked red for EOL.

---

## ⚙️ Setup Instructions

### What You Need to Do
1. Ensure your exported report is saved in a worksheet named `"Table"`.
2. Place your EOL CPU list in a file named `EOL_CPU_List.xlsx` in your **Downloads** folder.
   - The list should be in **Column A** of the first sheet.
3. Open the Excel file and run the macro:
   - Press `Alt + F11` to open the VBA editor.
   - Insert a new module (`Insert > Module`) and paste the macro code.
   - Press `F5` or run it from Excel via `Alt + F8`.

---

## ✅ TODO List

- [x] Highlight EOL CPUs in red.
- [x] Highlight Servers in blue.
- [x] Convert memory and storage columns to numeric values.
- [ ] Highlight **RAM upgrade** if `Agent Memory Total` < 16GB (purple cell).
- [ ] Highlight **SSD upgrade** if `C Drive Free Percent` < 25% (light blue cell).
- [ ] Highlight **Needs Pro** if Windows version is Home (orange cell).
- [ ] Highlight **Upgradeable to Win11 Pro** if on Win10 Pro and no issues (yellow row).
- [ ] Highlight **Already on Win11 Pro** if no issues and already on Win11 Pro (green row).

---

## 📄 VBA Code
The macro is stored in `HighlightEOLCPUs` and can be found in the VBA editor under `Module1`.

---
