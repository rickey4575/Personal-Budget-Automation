@echo off
cd /d "%~dp0"
python budget_tool.py --input statements --use-ai-categories
pause
