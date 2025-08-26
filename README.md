---

# SM9 Ticket Generator

SSCTicketgen is a PyQt6 vibe coded app designed to assist SSC Helpdesk agents for SM9 ticket creation.

> This tool is built to speed up the ticketing process and reduce repetitive typing for common issues.

---

## 🛠️ How to Run

### 1. Install Python (without Admin Rights)

Download Python from python.org

#### Installation Steps:
1. Run the python installer.
2. ✅ Check **“Add Python to PATH”**.
3. 🔽 Click **“Customize installation”**.
4. Leave defaults checked, click **Next**.
5. On “Advanced Options”:
   - ✅ Check **“Install for me only”** *(if you don’t have admin rights)*.
   - ❌ Skip **“Install for all users”** unless you have admin rights.
6. Click **Install**.

---

### 2. Install Dependencies

Open PowerShell and run:

```powershell
pip install PyQt6
pip install Jinja2
```

Helpful links:
- [PyQt6 Packaging Guide](https://www.pythonguis.com/tutorials/packaging-pyqt6-applications-windows-pyinstaller/)
- [Jinja2 Documentation](https://jina.ai/serve/get-started/install/)

---

### Alternatively, run the installer bat file

   - Need git installed to clone the repo
   - Open the `installer` folder in your SSCTicketGen project directory.
   - Run the batch file `sscticketgen-autoinstall.bat`.

   **Follow the Prompts**
   - The installer will ask if you want to clone the SSCTicketGen repo.
   - If you choose "Y", you will be prompted for a directory to save the repo.
   - The script will check for Python and prompt you to install it if not found.
   - Dependencies will be installed automatically.

### 3. Run the Application

Navigate to the project folder and run:

```powershell
python main.py
```

Or compile it into an executable using `pyinstaller`:

```powershell
pyinstaller --onefile main.py
```

---
<p align="center">
  
   ![](https://github.com/jamieisonline/SSCTicketGen/blob/main/Screenshot%202025-06-29%20163030.png)
</p>

---

## 📌 Notes

- Common issues are still being documented and added.
- Still buggy and unoptimized

---
