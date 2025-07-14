# 📊 Report Automation Tool (PowerShell + Batch)

This tool automates link updating and file handling for **Stability Reports** and **Transactional Reports** in PowerPoint and Excel using a combination of `.bat` and `.ps1` scripts.

---

## ⚙️ Setup Instructions

Follow these steps carefully to prepare and run the tool.

---

### ✅ Step 1: Prepare the Files

1. **Copy the code** for:
   - The `.bat` script → name it: `reports.bat`
   - The `.ps1` script → name it: `change_links.ps1`

2. Ensure **both files are placed in the same folder**.

---

### 📂 Step 2: Folder and Naming Conventions

The scripts rely on existing folder names to extract and apply dates correctly.

- For **Stability Reports**, use:
  ```
  Folder name format: July 12
  ```
- For **Transactional Reports**, use:
  ```
  Folder name format: 12th July
  ```

> 🟡 The script reads the folder name to determine how to build new file paths. Make sure the folders are named exactly in those formats depending on the report type.

---

### ▶️ Step 3: Run the Script

1. Double-click `reports.bat`
   - This will execute `change_links.ps1` with the correct parameters.
   - Make sure you **do not close, rename, or move** any files while the script is running.

2. Wait until you see:
   ```
   Complete
   Press any key to continue . . .
   ```

---

### ⚠️ Step 4: Open PowerPoint and Update Links

1. Open the resulting `.pptx` file.
2. A dialog box will appear asking if you'd like to **update links**.
3. **Click "Update Links"** to reflect the latest Excel data.

---

## 📝 Notes

- You **do not** need to install Git or any tools — just copy/paste the code and use as instructed.
- Do **not interrupt or modify** the folder or files while the script is running.

