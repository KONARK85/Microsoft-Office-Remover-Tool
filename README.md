<div align="center">

# 🗑️ Microsoft Office Remover Tool

**A PowerShell script to forcefully remove corrupted or stubborn Microsoft Office installations from Windows**

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-00e5ff?style=flat-square&logo=powershell&logoColor=white)](https://github.com/KONARK85/Microsoft-Office-Remover-Tool)
[![Windows](https://img.shields.io/badge/Windows-10%2F11-0078D6?style=flat-square&logo=windows&logoColor=white)](https://github.com/KONARK85/Microsoft-Office-Remover-Tool)
[![License](https://img.shields.io/badge/License-MIT-39ff6a?style=flat-square)](LICENSE)
[![Author](https://img.shields.io/badge/Author-KONARK85-00e5ff?style=flat-square)](https://github.com/KONARK85)

</div>

---

## 🔍 What Is This?

This PowerShell script is designed to **forcefully remove any version of Microsoft Office** from Windows — especially when Office has become corrupted and cannot be uninstalled through normal methods like Control Panel or Settings.

---

## ⚠️ When Should You Use This?

Use this tool when:

- Microsoft Office **won't uninstall** from Control Panel or Settings
- Office installation is **corrupted or broken**
- You're getting errors when trying to **reinstall Office**
- The standard **Microsoft Support and Recovery Assistant** has failed
- You need a **clean slate** before reinstalling Office

---

## 🖥️ Requirements

| Requirement | Details |
|---|---|
| **OS** | Windows 10 or Windows 11 |
| **PowerShell** | Version 5.1 or higher |
| **Permissions** | Must run as **Administrator** |
| **Internet** | Not required |

---

## 🚀 How to Use

### Step 1 — Download the Script

Click the green **Code** button → **Download ZIP**, or clone it:

```bash
git clone https://github.com/KONARK85/Microsoft-Office-Remover-Tool.git
```

### Step 2 — Open PowerShell as Administrator

- Press `Windows + S` → search **PowerShell**
- Right-click → **Run as Administrator**

### Step 3 — Navigate to the Script Folder

```powershell
cd C:\path\to\Microsoft-Office-Remover-Tool
```

### Step 4 — Run the Script

```powershell
powershell -ExecutionPolicy Bypass -File .\Remove-MicrosoftOffice.ps1
```

---

## 🛡️ Safety Notice

> ⚠️ **Always create a System Restore Point before running this script.**
>
> This script makes deep changes to your system. While it is designed to be safe, it is recommended to back up important data first.

---

## 📋 What the Script Does

- Detects all installed versions of Microsoft Office
- Terminates any running Office processes
- Removes Office via
