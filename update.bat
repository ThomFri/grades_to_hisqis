@echo off

CALL cd %0\..\

echo Fuehre Update in "%cd%" aus...
echo.

git submodule update --recursive --remote

echo.
echo Update fertig.
echo.

pause