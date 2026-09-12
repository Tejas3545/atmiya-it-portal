@echo off
title Atmiya University - Dean LMS Portal
echo ============================================
echo  Atmiya University - Dean LMS Portal
echo ============================================
echo.
echo Starting server...
echo Open your browser at: http://localhost:5000
echo.
echo Press Ctrl+C to stop the server.
echo.
cd /d "%~dp0"
python app.py
pause
