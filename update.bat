@echo off

CALL cd %0\..\

echo Fuehre Update in "%cd%" aus...
echo.

git pull

echo Update fertig.
echo.

pause