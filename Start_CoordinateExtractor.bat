@echo off
echo Starting Coordinate Extractor...
cd /d "%~dp0Application"
set PATH=%PATH%;%~dp0Tesseract-OCR
start CoordinateExtractor.exe
