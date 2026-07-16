@echo off
python -m pip install --upgrade pip
pip install -r requirements.txt
pyinstaller --clean --noconfirm build.spec
pause
