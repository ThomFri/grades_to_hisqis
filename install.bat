@echo off

echo Checking Python...

python --version
if errorlevel 0 goto continueEins
echo Python ist nicht installiert!
echo Download unter https://www.python.org/downloads/
exit

:continueEins

echo Checking PIP..

pip --version
if errorlevel 0 goto continueZwei
echo PIP ist nicht installiert!
echo (Bitte auch überprüfen, dass Python 3 und nicht Python 2 installiert ist!)
echo Installationsanleitung unter https://geekflare.com/de/python-pip-installation/
exit

:continueZwei

echo Installiere Python-Bibliotheken

pip install -r requirements.txt

echo Ende
pause