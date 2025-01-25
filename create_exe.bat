@echo off
cd C:\Users\user-pc\Desktop\palm_download_project\pm_proxy
pip freeze > requirements.txt
pyinstaller -F main.py --hidden-import pytubefix
move /y dist\main.exe main.exe
rmdir /s /q build
rmdir dist
del main.spec