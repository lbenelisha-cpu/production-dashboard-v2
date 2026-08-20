@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Opening IML Production Dashboard V4.8.1 in local preview mode...
start "" "%CD%\index.html"
