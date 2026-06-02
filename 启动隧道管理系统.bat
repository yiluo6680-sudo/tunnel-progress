@echo off
chcp 65001 >nul
title 隧道工程管理系统

cd /d %~dp0

echo.
echo ╔════════════════════════════════════╗
echo ║     隧道工程管理系统               ║
echo ║     启动中...                      ║
echo ║                                   ║
echo ║  浏览器打开后，侧边栏可切换：       ║
echo ║   🚇 隧道可视化                     ║
echo ║   📋 计量生成器                     ║
echo ╚════════════════════════════════════╝
echo.

streamlit run app.py
pause
