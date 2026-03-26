\# Word to PDF Macro



\## Overview



This tool is a Microsoft Word VBA macro that converts the currently open Word document into a PDF file and saves it in the same directory as the source document.



\---



\## Why This Exists (Problem)



I work extensively with Microsoft Word documents, and at the end of each task, I need to convert them into PDF.



The usual process is repetitive and slow:



\* Click \*\*Save As\*\*

\* Search/select the correct folder

\* Change file format to PDF

\* Click Save



This involves multiple steps and unnecessary clicks every single time.



To eliminate this friction and save time, I leveraged VBA macros to automate the entire process into a single action.



\---



\## Solution



This macro:



\* Automatically detects the document location

\* Keeps the same file name

\* Saves a PDF version instantly in the same folder



No manual navigation. No repeated clicks.



\---



\## Features



\* One-click Word → PDF conversion (hot keys combination)

\* Saves in the same directory as source file

\* Maintains original file name

\* Fast and minimal interaction



\---



\## Files



\* `export-to-pdf.bas` — VBA macro module containing the conversion logic



\---



\## How It Works



The macro:



1\. Reads the active document path

2\. Extracts the file name (without extension)

3\. Generates a `.pdf` file name

4\. Exports the document as a PDF



\---



\## Usage



\### Step 1: Import the Macro



1\. Open Microsoft Word

2\. Press `Alt + F11` to open VBA Editor

3\. Go to \*\*File → Import File\*\*

4\. Select `export-to-pdf.bas`



\### Step 2: Run the Macro



1\. Open the Word document you want to convert

2\. Press `Alt + F8`

3\. Select the macro

4\. Click \*\*Run\*\*



\---



\## Output



\* PDF is saved in the same folder as the Word document

\* File name format:



&#x20; ```

&#x20; original\_filename.pdf

&#x20; ```



\---



\## Requirements



\* Microsoft Word (Desktop version)

\* Macros enabled



\---



\## Notes



\* Works only on saved documents (needs a valid file path)

\* Existing PDF with same name will be overwritten



\---



\## Example



Input:



```

C:\\Docs\\Report.docx

```



Output:



```

C:\\Docs\\Report.pdf

```



\---



\## License



Free to use and modify.



