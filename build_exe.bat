@echo off
set "PYTHON=py -3"
py -3 --version >nul 2>nul
if errorlevel 1 set "PYTHON=python"
%PYTHON% -m pip install --upgrade pip
%PYTHON% -m pip install -r requirements.txt
%PYTHON% -m PyInstaller --noconfirm --clean pre_submission_ai_qc.spec
echo Build complete. Check dist\PreSubmissionAIQC
pause
