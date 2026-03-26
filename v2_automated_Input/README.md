# Job Application Automation (Browser + Python Integration)

## Overview

This tool automates the job application workflow by connecting a **browser bookmark (JavaScript)** with a **Python backend server**.

With a single click on a job listing page (only StepStone), job details are extracted and sent to a local Python server, which then:

* Creates a structured folder
* Generates a customized cover letter
* Updates an Excel tracker

---

## Why This Exists (Problem)

While applying for jobs, the process was repetitive:

* Copy company name, role, location manually
* Create folders
* Edit cover letters
* Track everything in Excel

Even with a Python script, it still required manual input.

---

## Key Idea

Instead of typing details manually, extract them directly from the job page.

Using browser **Inspect (DOM analysis)**, a custom bookmark (JavaScript) was created to:

* Read job details from the webpage
* Send them automatically to a Python backend

---

## Solution Architecture

### Flow

1. Open job listing (e.g., StepStone)
2. Click custom bookmark
3. Bookmark script extracts:

   * Company
   * Job title
   * Location
   * Job link
4. Sends data to local server (`localhost:5000`)
5. Python server processes everything

---

## Components

### 1. Browser Bookmark (Frontend)

* Built using JavaScript
* Created via browser **Inspect tool**
* Extracts data from webpage DOM
* Sends POST request to backend

---

### 2. Python Server (Backend)

* Built with Flask
* Receives job data via API
* Handles:

  * Folder creation
  * Cover letter generation
  * Excel tracking

Main logic:

* `server.py` 

---

### 3. Batch File

* `start_server.bat`
* Starts the backend instantly without manual commands

---

## Features

* One-click job processing from browser
* No manual data entry
* Auto folder organization
* Dynamic cover letter generation
* Excel tracker with clickable links
* Fully automated workflow

---

## Project Structure

```id="m1uxh9"
project-root/
│── data/
│── templates/
│   └── cover_letter_template.docx
│
│── server.py
│── start_server.bat
```

---

## Setup

### 1. Install Dependencies

```bash
pip install flask flask-cors python-docx openpyxl
```

---

### 2. Start Server

```id="2g8w1n"
start_server.bat
```

Server runs on:

```
http://localhost:5000
```

---

### 3. Create Bookmark (Important)

1. Create a new bookmark in your browser
2. Paste your JavaScript code (from Inspect-based extraction)
3. Save it

👉 This bookmark is the trigger for automation

---

## Usage

1. Open job listing page (e.g., StepStone)
2. Click the bookmark
3. Done ✅

Behind the scenes:

* Data is extracted
* Sent to backend
* Files + Excel updated automatically

---

## Example Flow

```id="9zzp0z"
StepStone Page → Click Bookmark → Python Server → Files Created
```

---

## Notes

* Works based on webpage structure (DOM)
* If website layout changes, bookmark script may need updates
* Server must be running before clicking bookmark
* Designed specifically for StepStone (can be adapted)

---

## Future Improvements

* Support multiple job portals (LinkedIn, Indeed)
* Add browser extension (instead of bookmark)
* Add UI dashboard
* PDF auto-generation
* Cloud deployment

---

## License

Free to use and modify.
