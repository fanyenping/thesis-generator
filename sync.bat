@echo off
cd /d "%~dp0"
echo.
echo === 同步到 GitHub ===
echo.

git add .

:: 取得目前時間作為 commit message
for /f "tokens=1-3 delims=/ " %%a in ('date /t') do set DATE=%%a/%%b/%%c
for /f "tokens=1-2 delims=: " %%a in ('time /t') do set TIME=%%a:%%b

set MSG=%DATE% %TIME% 更新
if not "%~1"=="" set MSG=%~1

git commit -m "%MSG%"
git push origin master

echo.
echo === 同步完成 ✓ ===
echo https://github.com/fanyenping/thesis-generator
echo.
pause
