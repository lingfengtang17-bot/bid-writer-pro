@echo off
chcp 65001 >nul
title AI智能编标专家 Pro v2.0 - Codex & WorkBuddy 一键部署工具
echo ============================================================
echo   AI 智能编标专家 Pro v2.0 (WorkBuddy / Codex 免外接 API 版)
echo   正在同步技能目录并检查 Python 依赖...
echo ============================================================
echo.

set "SKILL_SRC=%~dp0.codebuddy\skills\bid-writer-pro"
set "CODEX_DST=%USERPROFILE%\.agents\skills\bid-writer-pro"

echo [1/3] 正在同步技能包到 Codex 与 Claude Code 全局目录...
if not exist "%USERPROFILE%\.agents\skills" mkdir "%USERPROFILE%\.agents\skills"
xcopy /E /I /Y "%SKILL_SRC%" "%CODEX_DST%" >nul
if not exist "%USERPROFILE%\.claude\skills" mkdir "%USERPROFILE%\.claude\skills"
xcopy /E /I /Y "%SKILL_SRC%" "%USERPROFILE%\.claude\skills\bid-writer-pro" >nul
echo       [OK] Codex 目录 (.agents\skills\bid-writer-pro) 已同步
echo       [OK] Claude Code 目录 (.claude\skills\bid-writer-pro) 已同步
echo       [OK] WorkBuddy 目录 (.codebuddy\skills\bid-writer-pro) 已就绪
echo.

echo [2/3] 正在检查并安装 Word/PDF 核心依赖库...
python -m pip install -r "%SKILL_SRC%\requirements.txt" -i https://pypi.tuna.tsinghua.edu.cn/simple
echo.

echo [3/3] 环境配置完成！
echo ------------------------------------------------------------
echo 使用方法（无需配置第三方 API Key）：
echo   1. WorkBuddy 用户：直接打开本文件夹即可对话编标。
echo   2. Codex 用户：打开本文件夹或任意目录，输入 "使用 $bid-writer-pro 解析招标文件"。
echo.
echo 算力补给站（作者自营链动小店，长标书防中断）：
echo   主站：https://wzyp.cn/shop/7KKGS49M
echo   Plus直充：https://wzyp.cn/item/qpbnfn (专属码 DOUYINVIP888 立减 10 元)
echo   Codex额度：https://wzyp.cn/item/ctma3h
echo   作者微信：x2494023067
echo ============================================================
pause
