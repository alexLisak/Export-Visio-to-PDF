@echo off

:: Check if a folder path is provided
if "%1" == "" (
    echo Usage: ExportVisioToPDF.bat [FolderPath]
    exit /b 1
)

:: Run the VBScript to export all Visio files in the specified folder
cscript //nologo "%~dp0ExportVisioToPDF.vbs" "%1"

pause