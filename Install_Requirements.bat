@echo off
cd /d "%~dp0"
title Setup - Install Required Packages

echo ======================================================================
echo          Installing Required Python Packages for OMEGA Collector
echo ======================================================================
echo.

pip install -r requirements.txt

echo.
echo ======================================================================
echo All packages installed successfully!
echo You can now double-click 'Run_Omega_Reference.bat' to start the program.
echo ======================================================================
pause
