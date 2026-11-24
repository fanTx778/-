@echo off
chcp 65001 >nul
cd /d "%~dp0"

rem 用法：
rem   push-tag.bat v1.3          推送指定标签
rem   不带参数运行会显示本提示
rem 用途：github.com 被 DNS 污染时，临时把 140.82.113.3 写进 hosts 完成推送，结束后还原。

if "%~1"=="" (
  echo.
  echo   用法：push-tag.bat ^<标签名^>
  echo   例如：push-tag.bat v1.3
  echo.
  echo   当前本地标签：
  git tag
  echo.
  pause
  exit /b 1
)
set "TAG=%~1"

net session >nul 2>&1
if %errorlevel% neq 0 (
  echo.
  echo   请右键本文件，选择「以管理员身份运行」。
  echo.
  pause
  exit /b 1
)

set "HOSTS=%SystemRoot%\System32\drivers\etc\hosts"
set "BAK=%TEMP%\hosts.bishe.bak"

echo ================================================
echo   临时修正 github.com 解析并推送标签 %TAG%
echo ================================================
echo.

copy /y "%HOSTS%" "%BAK%" >nul
findstr /c:"140.82.113.3 github.com" "%HOSTS%" >nul
if errorlevel 1 (
  echo 140.82.113.3 github.com>>"%HOSTS%"
  echo   已临时添加 140.82.113.3 github.com
) else (
  echo   已存在对应记录，跳过
)
ipconfig /flushdns >nul
echo.

echo   正在推送标签 %TAG% ...
echo ------------------------------------------------
git push origin %TAG%
set "RC=%errorlevel%"
echo ------------------------------------------------
echo.

echo   正在还原 hosts ...
copy /y "%BAK%" "%HOSTS%" >nul
ipconfig /flushdns >nul
echo   已还原
echo.

if "%RC%"=="0" (
  echo [成功] 标签 %TAG% 已推送到 https://github.com/fanTx778/-
) else (
  echo [失败] 请把上面的报错截图发给助手
)
echo.
pause
