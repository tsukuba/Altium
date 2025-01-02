@echo off

cmd /c start /min "" PowerShell -NoProfile -ExecutionPolicy RemoteSigned -WindowStyle Hidden -File "%~dp0Convert_Centroid_File.ps1" %*

