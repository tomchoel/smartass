@echo off
rem TN Claude task pane - local server (https://localhost:3000). Keep this window open while you use the pane.
cd /d "%~dp0"
title TN Claude pane
node server.mjs
