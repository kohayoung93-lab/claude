@echo off
REM 더블클릭으로 실행하는 Windows용 실행 파일
cd /d "%~dp0"
python append_ledger.py
echo.
pause
