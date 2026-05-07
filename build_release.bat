@echo off
setlocal
set "PYTHON=py -3"
py -3 --version >nul 2>nul
if errorlevel 1 set "PYTHON=python"
%PYTHON% -m pip install --upgrade pip
%PYTHON% -m pip install -r requirements.txt
%PYTHON% -m PyInstaller --noconfirm --clean pre_submission_ai_qc.spec
if not exist release mkdir release
powershell -NoProfile -ExecutionPolicy Bypass -Command "if (Test-Path 'release\PreSubmissionAIQC-portable.zip') { Remove-Item 'release\PreSubmissionAIQC-portable.zip' -Force }; Compress-Archive -Path 'dist\PreSubmissionAIQC\*' -DestinationPath 'release\PreSubmissionAIQC-portable.zip' -Force"
echo Release package: release\PreSubmissionAIQC-portable.zip
endlocal
pause
