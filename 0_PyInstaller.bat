cls
rmdir /s /q build
rmdir /s /q dist

pyinstaller --windowed --noconfirm ^
    --icon=icon.ico ^
    --add-data "icon.ico;." ^
    --hidden-import win32gui ^
    --hidden-import win32con ^
    --hidden-import win32api ^
    --hidden-import win32process ^
    --hidden-import win32event ^
    --hidden-import psutil ^
    --hidden-import PySide6 ^
    KeyForge.py