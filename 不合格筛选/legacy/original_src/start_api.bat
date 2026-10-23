@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo   审核不合格数据 Web 服务
echo   浏览器打开: http://127.0.0.1:8000
echo   停止服务: 关闭本窗口 或 Ctrl+C
echo ============================================
python api.py
pause
