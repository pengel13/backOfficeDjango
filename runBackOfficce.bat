@echo off
cd \
cd C:\backOfficeDjango
call venv\Scripts\activate.bat
"python" "manage.py" "runserver" "0.0.0.0:8051"
pause

