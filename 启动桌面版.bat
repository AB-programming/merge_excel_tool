@echo off
cd /d "%~dp0"

rem 双击启动桌面版，需要本机装有 Python 3.11+。
rem 不想装 Python，就用「打包_exe.bat」生成 dist 里的免安装 exe。
rem 注意：本文件用 ANSI/GBK 保存，改成 UTF-8 会让 cmd 把中文行拆错。

set "PY=py -3"
where py >nul 2>nul || set "PY=python"

%PY% --version >nul 2>nul
if errorlevel 1 goto :no_python

%PY% merge_excel.py --gui
if errorlevel 1 pause
exit /b 0

:no_python
echo 没找到 Python，请先安装 Python 3.11 或更高版本，安装时勾选 Add to PATH。
echo 也可以先用「打包_exe.bat」生成免安装的 exe。
pause
exit /b 1
