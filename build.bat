@echo off
REM 로컬 PC에서 RadarExport.exe 를 직접 만들 때 사용합니다 (결과: dist\RadarExport.exe)
py -m pip install pyinstaller selenium webdriver-manager
py -m PyInstaller --onefile --windowed --name RadarExport --collect-all selenium radar_export_download.py
pause
