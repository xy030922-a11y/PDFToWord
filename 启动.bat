@echo off
setlocal
cd /d "%~dp0"

set "ENV_DIR=%~dp0.conda"
set "CONDA_EXE="

if exist "%ENV_DIR%\python.exe" goto check_dependencies

for /f "delims=" %%i in ('where conda 2^>nul') do if not defined CONDA_EXE set "CONDA_EXE=%%i"
if not defined CONDA_EXE if exist "%USERPROFILE%\anaconda3\Scripts\conda.exe" set "CONDA_EXE=%USERPROFILE%\anaconda3\Scripts\conda.exe"
if not defined CONDA_EXE if exist "%USERPROFILE%\miniconda3\Scripts\conda.exe" set "CONDA_EXE=%USERPROFILE%\miniconda3\Scripts\conda.exe"
if not defined CONDA_EXE goto no_conda

echo First run: creating the Conda environment. Please wait...
call "%CONDA_EXE%" create --prefix "%ENV_DIR%" python=3.12 pip --yes
if errorlevel 1 goto error

:check_dependencies
"%ENV_DIR%\python.exe" -c "import pymupdf, docx, pdf2docx, win32com" >nul 2>nul
if not errorlevel 1 goto launch

echo Installing required packages. Please wait...
"%ENV_DIR%\python.exe" -m pip install --disable-pip-version-check -r "%~dp0requirements.txt"
if errorlevel 1 goto error

:launch
start "" "%ENV_DIR%\pythonw.exe" "%~dp0pdf_to_word.py"
exit /b 0

:no_conda
echo Conda was not found. Please install Anaconda or Miniconda first.
pause
exit /b 1

:error
echo.
echo Setup failed. Check your network or Conda installation and try again.
pause
exit /b 1
