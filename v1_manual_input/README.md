# Job Application Automation Tool

## Overview

This is a Python-based automation tool that streamlines the job application process by handling repetitive tasks such as folder creation, cover letter generation, and application tracking.

It eliminates manual effort and keeps everything structured and organized in one place.

---

## Why This Exists (Problem)

Applying to multiple jobs involves repetitive and time-consuming steps:

* Creating folders for each company
* Copying and editing cover letters
* Replacing company/job details manually
* Tracking applications in Excel
* Managing job links separately

Doing this repeatedly leads to:

* Wasted time
* Inconsistent organization
* Higher chances of mistakes

---

## Solution

This tool automates the entire workflow:

* Creates a structured folder for each application
* Generates a customized cover letter from a template
* Tracks applications automatically in Excel
* Stores job links in a clickable format

Additionally, a `.bat` file is included to make running the tool faster and effortless.

---

## Features

* Auto-numbered job folders (e.g., `01 Company_Role`)
* Dynamic cover letter generation using a template
* Automatic placeholder replacement:

  * `[COMPANY_NAME]`
  * `[JOB_TITLE]`
  * `[LOCATION]`
* Styled text formatting in generated document
* Excel-based application tracker
* Clickable job links stored in Excel
* One-click execution using batch file

---

## Project Structure

```
project-root/
│── data/
│   ├── 01 Company_Role/
│   │   └── Anschreiben_yashodhar_Hajare.docx
│   └── application_tracker.xlsx
│
│── templates/
│   └── cover_letter_template.docx
│
│── main.py
│── run_main.bat
```

---

## How It Works

### 1. User Input

The script prompts for:

* Company name
* Job title
* Location
* Job link

### 2. Folder Creation

* Creates a numbered folder inside `data/`
* Format:

  ```
  01 Company_Role
  ```

### 3. Cover Letter Generation

* Loads template
* Replaces placeholders dynamically
* Applies formatting (Calibri, blue, bold)
* Saves into the created folder

### 4. Excel Tracking

* Creates tracker if not present
* Appends new entry with job details
* Stores clickable job link

---

## Usage

### Option 1: Run Using Python

```bash
pip install python-docx openpyxl
python main.py
```

### Option 2: Run Using Batch File (Recommended)

Simply double-click:

```
run_main.bat
```

This will automatically execute the script without manually opening terminal or typing commands.

---

## Output

After execution:

* Job folder is created
* Cover letter is generated
* Excel tracker is updated

---

## Example

Input:

```
Company: BMW
Role: Data Analyst
Location: Munich
```

Output:

```
data/
└── 01 BMW_Data_Analyst/
    └── Anschreiben_yashodhar_Hajare.docx
```

---

## Requirements

* Python 3.x installed
* Required libraries:

  * python-docx
  * openpyxl

---

## Notes

* Folder numbering auto-increments
* Existing Excel file is reused
* Template file must exist before running
* Batch file assumes Python is installed and added to PATH

---

## Code Reference

Main logic implemented in:

* `main.py`

---

## Future Improvements

* GUI interface
* PDF export integration
* Email automation
* Multi-language template support

---

## License

Free to use and modify.
