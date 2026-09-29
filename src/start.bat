@echo off
title Batch Print Tool
rem =====================================================
rem  Batch Print Tool - launcher
rem  Double-click to open the tool. No admin required.
rem  The .ps1 file name is ASCII-only to avoid codepage
rem  issues with Chinese characters in cmd.exe.
rem =====================================================
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0batch-print-tool.ps1"
pause
