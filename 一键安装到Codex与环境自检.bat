@echo off
chcp 65001 >nul
title AI智能编标专家 Pro - Codex & WorkBuddy 一键部署工具
echo ============================================================
echo   🚀 AI智能编标专家 Pro (WorkBuddy / Codex 满血完整版)
echo   正在为您自动配置 Python 依赖与 Codex 技能目录...
echo ============================================================
echo.

set "SKILL_SRC=%~dp0.codebuddy\skills\bid-writer-pro"
set "CODEX_DST=%USERPROFILE%\.agents\skills\bid-writer-pro"

echo [1/3] 正在同步技能包到 Codex 全局目录 (%CODEX_DST%)...
if not exist "%USERPROFILE%\.agents\skills" mkdir "%USERPROFILE%\.agents\skills"
xcopy /E /I /Y "%SKILL_SRC%" "%CODEX_DST%" >nul
if not exist "%USERPROFILE%\.claude\skills" mkdir "%USERPROFILE%\.claude\skills"
xcopy /E /I /Y "%SKILL_SRC%" "%USERPROFILE%\.claude\skills\bid-writer-pro" >nul
echo       ✅ Claude Code 技能目录同步完成！
echo       ✅ Codex 技能目录同步完成！
echo       ✅ WorkBuddy 目录 (.codebuddy\skills\bid-writer-pro) 已就绪！
echo.

echo [2/3] 正在检查并安装 Word/PDF 处理依赖库...
python -m pip install -r "%SKILL_SRC%\requirements.txt" -i https://pypi.tuna.tsinghua.edu.cn/simple
echo.

echo [3/3] 环境配置完成！
echo ------------------------------------------------------------
echo 💡 使用方法：
echo   1. WorkBuddy 用户：直接用 WorkBuddy 打开本文件夹即可对话编标。
echo   2. Codex 用户：在 Codex 中输入 "使用 $bid-writer-pro 解析招标文件"。
echo.
echo ⚡ 算力与长标书防中断补给站（正规白卡直充自己号，带官方账单）：
echo    主站地址：https://wzyp.cn/shop/7KKGS49M
echo    Plus直充：https://wzyp.cn/item/qpbnfn （填专属优惠码 DOUYINVIP888 立减10元！）
echo    Codex额度：https://wzyp.cn/item/ctma3h
echo    作者微信：x2494023067（交流编标、反馈问题、开通会员）
echo ============================================================
pause
