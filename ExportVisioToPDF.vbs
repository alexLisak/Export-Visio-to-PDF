Dim fso, folder, files, file
Dim visioApp, visioDoc, visioPage, visioLayer, logFile

' Define constants for exporting to PDF
Const visFixedFormatPDF = 1
Const visDocExIntentPrint = 0
Const visPrintAll = 0

' Get the folder path from the command line argument
Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder(WScript.Arguments(0))

' Create a log file to track progress
Set logFile = fso.OpenTextFile(fso.BuildPath(folder.Path, "ExportLog.txt"), 8, True)
logFile.WriteLine "Starting Visio to PDF export process at " & Now

' Create the Visio application object
Set visioApp = CreateObject("Visio.Application")
visioApp.Visible = False ' Run Visio in the background

' Loop through all Visio files in the folder
For Each file In folder.Files
    If LCase(fso.GetExtensionName(file.Name)) = "vsdx" Or LCase(fso.GetExtensionName(file.Name)) = "vsd" Then
        On Error Resume Next
        ' Open the Visio document
        Set visioDoc = visioApp.Documents.OpenEx(file.Path, 64)
        If Err.Number <> 0 Then
            logFile.WriteLine "Error opening Visio file: " & file.Path & " - " & Err.Description
            Err.Clear
            Set visioDoc = Nothing
            ' Skip to the next file
            Err.Clear
        End If
        On Error Resume Next
        
        ' Loop through all pages to ensure they are fully visible and printable
        For Each visioPage In visioDoc.Pages
            ' Ensure all layers are visible and printable
            If visioPage.Type <> 2 Then ' Only process foreground pages (2 = visTypeBackground)
                For Each visioLayer In visioPage.Layers
                    visioLayer.CellsC(0).FormulaU = "1" ' Make layer visible
                    visioLayer.CellsC(2).FormulaU = "1" ' Ensure layer is printable
                Next
            End If
        Next
        
        ' Define the output PDF path
        pdfPath = fso.BuildPath(folder.Path, fso.GetBaseName(file.Name) & ".pdf")
        
        ' Check if the PDF already exists and delete it if necessary
        If fso.FileExists(pdfPath) Then
            fso.DeleteFile pdfPath, True
            logFile.WriteLine "Overwriting existing file: " & pdfPath
        End If
        
        ' Export the Visio document to PDF
        On Error Resume Next
        visioDoc.ExportAsFixedFormat visFixedFormatPDF, pdfPath, visDocExIntentPrint, visPrintAll
        If Err.Number <> 0 Then
            logFile.WriteLine "Error exporting Visio file to PDF: " & file.Path & " - " & Err.Description
            Err.Clear
        Else
            logFile.WriteLine "Successfully exported: " & file.Path & " to " & pdfPath
        End If
        On Error Resume Next
        
        ' Close the Visio document
        visioDoc.Close
        Set visioDoc = Nothing
    End If
Next

' Quit Visio
visioApp.Quit

' Write completion message to log file
logFile.WriteLine "Completed Visio to PDF export process at " & Now
logFile.Close

' Clean up
Set visioApp = Nothing
Set fso = Nothing
Set logFile = Nothing