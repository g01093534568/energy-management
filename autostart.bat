@echo off
chcp 65001 > nul
cd /d "E:\[업무]\에너지관리"
start /MIN "" "C:\Program Files\nodejs\node.exe" server.js
