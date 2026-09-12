@echo off
setlocal
cd /d "%~dp0"
echo ClassPay Updater - GitHub source to GAS
echo Target Script ID: 1eNC3EuiClJWHoJTZ2-AuT0WTJc8gaiabhiLfm5VJJSgQT4SKEjCSJELh
call clasp push
if errorlevel 1 (
  echo.
  echo Push failed. Check clasp login and Apps Script API settings.
  exit /b 1
)
echo.
echo Updater source was pushed successfully.
endlocal
