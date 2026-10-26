@echo off
chcp 65001 >nul
py -m pip install --upgrade pip
py -m pip install -r requirements.txt
py -m pip install pyinstaller
pyinstaller --noconfirm --clean --windowed --name "Attestations Ouarzazate" --hidden-import win32timezone app.py
echo.
echo Terminé. Vérifiez le dossier dist.
pause
