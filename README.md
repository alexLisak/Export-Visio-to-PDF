# Export-Visio-to-PDF
Vbs script and batch file you can add to the path to execute a Visio to PDF conversion on a batch of files in the current directory

**VBScript to Handle the Visio Export**
The VBScript (ExportVisioToPDF.vbs) opens each Visio file in the directory and exports it to a PDF. To ensure all pages are exported correctly, the script explicitly loops through all pages of each document and sets each page to be visible and printable.

**Batch File to Run the VBScript**
The batch file (ExportVisioToPDF.bat) easily calls the VBScript from the command line.

**Using the Scripts**
Save ExportVisioToPDF.vbs and ExportVisioToPDF.bat in a convenient location.
Add the Folder to PATH (Optional): Add the folder containing the batch and VBScript files to your system PATH environment variable.
Open Command Prompt in the Desired Folder
Run the Batch File: In the command prompt, run the batch file by providing the current folder as an argument:
ExportVisioToPDF.bat .
The . represents the current folder, meaning the script will export all Visio files in this folder to PDFs.
