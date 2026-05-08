═══════════════════════════════════════════════════
  FEAAM Contact Matcher — Setup & Run Instructions
═══════════════════════════════════════════════════

REQUIREMENTS
────────────
• Python 3.9 or higher  (https://www.python.org/downloads/)
• pip (comes with Python)

FIRST-TIME SETUP  (do this once)
─────────────────────────────────
1. Extract this ZIP to any folder on your computer.

2. Open a terminal / command prompt in that folder:
     Windows: Shift + right-click the folder → "Open PowerShell window here"
     Mac/Linux: open Terminal, then  cd /path/to/feaam-matcher

3. Install dependencies:
     pip install -r requirements.txt

RUN THE APP
───────────
1. In the same terminal:
     python app.py

2. Open your browser and go to:
     http://localhost:5051

3. The app is ready. Upload your files, generate messages, download Excel.

STOP THE APP
────────────
Press  Ctrl + C  in the terminal.

REFERENCE FILES
───────────────
The  reference_files/  folder already contains:
  • FEAAM_Outreach_FullSet.xlsx   — output format template (Output Reference tab)
  • Christians_annotations.docx  — style notes
  • Ebus_Feedback_Per_Mail.docx  — message correction rules

These persist between sessions. Add more via the app's Reference Files section.

NOTES
─────
• All uploaded files are saved permanently in  reference_files/
• The app runs locally — no internet connection required after setup
• Default port is 5051. Change it by setting the PORT environment variable:
    PORT=8080 python app.py

═══════════════════════════════════════════════════
