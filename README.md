# Job Application Automation System

## Overview

This project is a complete automation system designed to reduce time, effort, and mental fatigue during job applications.

It combines:

* Manual input tools (V1)
* Browser-based automation (V2)
* Utility tools (Word → PDF)

Together, these eliminate repetitive steps and streamline the entire workflow.

---

## Why This Exists (Problem)

Applying to jobs repeatedly involves:

* Copying job details manually
* Creating folders in defined name for each application
* Editing repetative things in cover letters (company name, Location, Job title, date etc)
* Saving documents in proper formats
* Tracking everything in Excel

This process is:

* Time-consuming
* Repetitive
* Mentally exhausting

---

## Solution

This project introduces two workflows:

### V1 — Manual Input (First and Fallback System)

Created for all the job and is first version 

Used when automation is not possible (e.g., non-StepStone websites).

### V2 — Automated Input (Primary System)

Uses browser + backend integration to eliminate manual work completely.

Additionally, utility tools further reduce effort (e.g., Word → PDF macro).

---

## V1 — Manual Input Workflow

### Use Case

* Job sites StepStone and other 
* When automation is not supported

### How It Works

* Run Python script manually
* Enter:

  * Company
  * Job title
  * Location
  * Job link

### What It Does

* Creates structured folder with predefined name structure
* Generates cover letter from template
* Updates Excel tracker

---

## V2 — Automated Input Workflow (StepStone)

### Use Case

* StepStone job listings
* Fully automated process

### How It Works

1. Open job listing
2. Click custom browser bookmark
3. Bookmark extracts job data using DOM (Inspect-based)
4. Sends data to local Python server
5. Backend processes everything automatically

---

## Architecture

```id="archflow"
Browser (Bookmark JS)
        ↓
Flask Server (Python)
        ↓
File System + Excel + Word Automation
```

---

## Components

### 1. Manual Tool (CLI)

* `main.py`
* Input-based processing

---

### 2. Automation Server

* `server.py`
* Flask-based API
* Handles automated requests from browser

---

### 3. Browser Bookmark

* JavaScript-based
* Built using Inspect (DOM extraction)
* Sends job data to backend

---

### 4. Utility Tool — Word to PDF Macro

* VBA macro
* Converts Word documents to PDF instantly
* Saves in same location
* Eliminates repetitive “Save As” steps

---

## Project Structure

```id="projstruct"
project-root/
│
├── data/
│   └── (generated folders + generated cover letters inside)
│   └── application_tracker.xlsx
│
├── project_demo/
│   └── project_demo.mp4
│
├── templates/
│   └── cover_letter_template.docx
│
├── Tools/
│   └── export-to-pdf_macro.bas
│   └── README.md
│
├── V1_manual/
│   ├── main.py
│   └── run_main.bat
│   └── README.md
│
├── V2_automation/
│   ├── server.py
│   └── start_server.bat
│   └── README.md
│
└── README.md
```

---

## Setup

### Install Dependencies

```bash id="deps"
pip install flask flask-cors python-docx openpyxl
```

---

### Run Manual Tool (V1)

```id="runv1"
run_main.bat
```

---

### Run Automation Server (V2)

```id="runv2"
start_server.bat
```

Server runs on:

```id="m0vcvm"
http://localhost:5000
```

---

### Setup Bookmark (V2)

* Create a browser bookmark
* Paste JavaScript (DOM extraction script)
* Use it on StepStone job pages

---

## Usage Summary

### Scenario 1 — StepStone (Fastest)

* Start server
* Open job page
* Click bookmark
* Done ✅

---

### Scenario 2 — Other Websites

* Run manual script
* Enter details
* Done ✅

---

### Final Step — PDF Conversion

* Use Word macro
* Convert cover letter to PDF instantly

---

## Important Note

* The **main body of the cover letter should always be tailored** for each company and job role
* This system is designed to **support and accelerate** the process, not replace thoughtful customization

---

## Key Benefits

* Massive reduction in repetitive work
* Saves **15+ minutes per application**
* Reduces mental fatigue significantly
* Keeps everything structured and consistent
* Works for both automated and manual scenarios

---

## Limitations

* Automation currently works only for StepStone
* Bookmark depends on website structure (DOM)
* Server must be running for automation

---

## Future Improvements

* Support more job platforms (LinkedIn, Indeed)
* Convert bookmark → browser extension
* Add GUI dashboard
* Auto PDF generation integration
* Cloud deployment

---

## License

Free to use and modify.
