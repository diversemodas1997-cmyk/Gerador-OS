@echo off
cd /d "%~dp0"
python -c "import bleak" 2>/dev/null || python -m pip install --user bleak
start "" pythonw colorimetro.py
