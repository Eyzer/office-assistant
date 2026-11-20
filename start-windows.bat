@echo off
setlocal
cd /d "%~dp0"
title office办公助手

set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY (
  where python >nul 2>nul && set "PY=python"
)
if not defined PY (
  echo.
  echo [错误] 没有找到 Python。
  echo 请先安装 Python 3.8 或更高版本，安装时勾选 "Add python.exe to PATH"。
  echo 下载地址：https://www.python.org/downloads/windows/
  echo.
  pause
  exit /b 1
)

%PY% -c "import openpyxl, docx" >nul 2>nul
if errorlevel 1 (
  echo [提示] 缺少必需依赖，正在安装 ...
  %PY% -m pip install -r requirements.txt
  if errorlevel 1 (
    echo.
    echo [错误] 依赖安装失败，请手动执行：%PY% -m pip install -r requirements.txt
    pause
    exit /b 1
  )
)

echo 正在启动 office办公助手，浏览器会自动打开 ...
echo 关闭本窗口或按 Ctrl+C 即退出程序。
echo.
%PY% main.py %*
set "RC=%ERRORLEVEL%"
if "%RC%"=="0" goto :eof
if "%RC%"=="130" goto :eof
echo.
echo 程序异常退出，代码 %RC%。上面最后几行是错误信息。
pause
endlocal
