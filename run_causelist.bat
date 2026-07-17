@echo off
timeout /t 10 /nobreak > nul
cd /d "C:\Users\Hafiz Ahmed\Desktop\Sindh\Sindh"
call .venv\Scripts\activate.bat
echo C:\Users\Hafiz Ahmed\Desktop\Sindh\Sindh\cause_lists | python sindh_causelist_auto.py