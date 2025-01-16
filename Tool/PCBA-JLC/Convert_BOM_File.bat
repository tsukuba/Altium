@echo off

cmd /c start /min "" PowerShell -NoProfile -ExecutionPolicy RemoteSigned -WindowStyle Hidden -File "%~dp0Convert_BOM_File.ps1" %*

