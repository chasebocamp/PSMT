# 🛠️ Precision System Maintenance Tool (PSMT)
> **The All-In-One (AIO) solution for stable, hardware-aware Windows maintenance.**

PSMT is designed to automate the deployment of BIOS/Firmware, Cumulative, and Feature updates while bypassing common hardware "hangs" associated with docking stations and high-resolution displays.

---

## 🚀 Quick Start
1.  **Place Files:** Ensure `run.bat` and `tool.ps1` are located in the same directory.
2.  **Launch:** Right-click `run.bat` and select **Run as Administrator**.
3.  **Select:** Follow the on-screen menu to choose your maintenance path.

---

## 🧬 How it Works
The tool performs a **Hardware Fingerprint** check on your machine using the `Win32_ComputerSystem` Model string. It then queries the **Windows Update Catalog** to find precision-matched updates specifically for your hardware ID.

### Anti-Hang Features:
* **Cold Boot Enforcement:** Disables Fast Startup to ensure a clean hardware handshake.
* **BitLocker Management:** Automatically suspends BitLocker for 2 reboots to prevent recovery lockouts during firmware flashes.
* **Docking Station Optimization:** Clears driver states before rebooting to prevent splash-screen freezes.

---

## ⚠️ Important Disclaimer
**Use at your own risk.** * **Proprietary First:** It is always recommended to use official manufacturer tools (Dell Command Update, HP Support Assistant, etc.) before using this scripted solution.
* **Power:** Ensure AC Power is connected. The script will attempt to block execution on battery, but user vigilance is required.
* **Liability:** No responsibility is taken for bricked motherboards or data loss. **You have been warned.**

---

## 📝 Requirements
* **OS:** Windows 10 or 11
* **Permissions:** Local Administrator Rights
* **Module:** `PSWindowsUpdate` (Script will auto-install if missing)
