@echo off
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
python model_notify.py >> logs\model_notify.log 2>&1
