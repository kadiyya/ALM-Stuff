@echo off
REM ============================================================
REM  Windows Gorev Zamanlayici'ya "her Sali 16:00" gorevi ekler.
REM  Script sprint bitis gunu degilse kendiliginden cikar,
REM  boylece mail sadece 2 haftada bir (sprint son gunu) gider.
REM  Saati degistirmek icin /ST degerini duzenleyin.
REM ============================================================
set SCRIPT_DIR=%~dp0
for /f "delims=" %%P in ('where pythonw') do (set PYW=%%P& goto :found)
echo pythonw bulunamadi. Python kurulu ve PATH'te olmali.
exit /b 1
:found
schtasks /Create /F /TN "RTC Sprint Story Mailer" ^
  /TR "\"%PYW%\" \"%SCRIPT_DIR%sprint_story_mailer.py\"" ^
  /SC WEEKLY /D TUE /ST 16:00
echo.
echo Gorev olusturuldu. Test icin:  schtasks /Run /TN "RTC Sprint Story Mailer"
