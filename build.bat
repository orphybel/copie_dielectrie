@echo off
rem Compile copie_dielectrie_gui.py en exe (a placer dans le meme dossier que le .py)
cd /d "%~dp0"

where python >nul 2>&1
if errorlevel 1 (
    echo Python n'est pas installe ou pas dans le PATH.
    echo Installe-le depuis https://www.python.org/downloads/ en cochant "Add python.exe to PATH".
    pause
    exit /b 1
)

python -m pip install --upgrade pyinstaller pywin32
if errorlevel 1 goto :error

python -m PyInstaller --noconfirm --onefile --windowed --name copie_dielectrie_gui copie_dielectrie_gui.py
if errorlevel 1 goto :error

echo.
echo OK : l'exe est dans "%~dp0dist\copie_dielectrie_gui.exe"
pause
exit /b 0

:error
echo.
echo ECHEC de la compilation, voir les messages ci-dessus.
pause
exit /b 1
