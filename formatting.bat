@echo off
setlocal enabledelayedexpansion

chcp 65001

cls

rem Prompt for ZIP file path using PowerShell
for /f "usebackq tokens=*" %%A in (`powershell -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName System.Windows.Forms; $dialog = New-Object System.Windows.Forms.OpenFileDialog; $dialog.Filter = 'ZIP files (*.zip)|*.zip|All files (*.*)|*.*'; $dialog.Title = 'Select the ZIP file'; if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $dialog.FileName } else { Write-Error 'No file selected' }"`) do (
    set "zip_file=%%A"
)

rem Check if the ZIP file path is provided
if "!zip_file!"=="" (
    echo ZIP file path is required.
    pause
    exit /b 1
)

echo 请输入你的文件名单（直接从excel复制粘贴后按Enter）：

set "index=0"
set "file_list="

:readLines
set "line="
set /p "line="

if not defined line (
    goto endRead
)

rem Store the line in the list variable
set "file_list=!file_list! !line!"
set /a "index+=1"

rem Repeat the loop
goto :readLines

:endRead

set keepfiles=""

for %%A in (!file_list!) do (
    for /f "delims=_" %%B in ("%%A") do (
        set "keepfiles=!keepfiles! %%B"
    )
)

rem Extract the directory path from the selected ZIP file
for %%A in ("!zip_file!") do set "destination=%%~dpA"

cls

rem Use PowerShell to unzip the file (example using Expand-Archive)
powershell -Command "Expand-Archive -Force -Path '!zip_file!' -DestinationPath '!destination!'"

rem remove files

cls

echo 请稍等（约1-2分钟）
for /r %%F in (*.docx) do (
      set "deleteFile=true"
      for %%K in (!keepfiles!) do (
          echo %%F | find "%%K" >nul
          if not errorlevel 1 (
              set "deleteFile=false"
          )
      )

      rem Delete the file if deleteFile is still true
      if "!deleteFile!"=="true" (
          del "%%F"
      )
)

endlocal
pause
