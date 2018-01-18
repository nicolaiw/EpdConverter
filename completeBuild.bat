@echo off
cls


.paket\paket.exe restore -v -f
if errorlevel 1 (
  exit /b %errorlevel%
)

"packages\FAKE\tools\Fake.exe" completeBuild.fsx
pause