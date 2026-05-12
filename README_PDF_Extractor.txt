# README: Outlook Email PDF Extractor
=====================================

## 1. PROJECT OBJECTIVE
This script automates the extraction of PDF attachments from Outlook saved mail files (.msg). 
It is designed to handle batch processing across multiple sub-folders (e.g., 170+ folders), 
saving the extracted PDF directly into the same directory where the source mail is located.

## 2. SYSTEM REQUIREMENTS
- Python 3.10 or higher.
- Windows OS (recommended for file path compatibility).
- A Python Virtual Environment (e.g., .venv312) is recommended for dependency management.

## 3. DEPENDENCIES
The script requires the 'extract-msg' library to read .msg files without needing 
the Outlook application open.

Installation:
Open your terminal/command prompt and run:
    pip install extract-msg

## 4. DIRECTORY STRUCTURE ASSUMPTIONS
The script assumes the following structure:
[Root Folder]
    |-- [Sub-folder 560468]
    |       |-- email_name.msg
    |-- [Sub-folder 560657]
    |       |-- another_email.msg
    ...

## 5. SETUP & EXECUTION
1. Open 'extract_pdfs.py' in your editor (e.g., VS Code).
2. Update the 'target_dir' variable with your specific path:
   target_dir = r"C:\Users\Rajesh.plaiwal\Downloads\4 may 2026"
3. Run the script:
   python extract_pdfs.py

## 6. SCRIPT LOGIC (ENGINEER'S SUMMARY)
- OS Walk: Iterates through the root directory to find sub-folders.
- Filter: Specifically looks for files with the .msg extension.
- Parser: Uses 'extract_msg.Message' to load the binary email data.
- Attachment Handler: 
    - Identifies 'longFilename' or 'name' of attachments.
    - Filters for '.pdf' extension (case-insensitive).
    - Writes the attachment 'data' (binary) to a new file in the sub-folder.
- Cleanup: Calls 'msg.close()' to ensure file handles are released after processing each folder.

## 7. TROUBLESHOOTING
- PermissionError: Ensure the sub-folders are not 'Read-Only' and the destination PDFs are not open in another app.
- ModuleNotFoundError: Run 'pip install extract-msg' inside the active virtual environment.
- No PDFs Found: Verify that the emails actually contain .pdf attachments (not just links or images).
