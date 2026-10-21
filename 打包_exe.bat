@echo off
cd /d "%~dp0"

rem 打包成两个免安装的 exe（需要联网；PyInstaller 只在打包这一步需要）
rem   dist\Excel合并工具.exe   双击打开窗口，给不敲命令的人用
rem   dist\merge_excel.exe     命令行版，参数和 python merge_excel.py 完全一样
rem 注意：本文件用 ANSI/GBK 保存，改成 UTF-8 会让 cmd 把中文行拆错。

set "PY=py -3"
where py >nul 2>nul || set "PY=python"

%PY% --version >nul 2>nul
if errorlevel 1 goto :no_python

echo [1/3] 准备 PyInstaller…
%PY% -m pip install --quiet pyinstaller
if errorlevel 1 goto :fail

echo [2/3] 打包桌面版（无黑窗口）…
%PY% -m PyInstaller --noconfirm --clean --onefile --windowed --name "Excel合并工具" --hidden-import openpyxl --hidden-import xlrd gui.py
if errorlevel 1 goto :fail

echo [3/3] 打包命令行版…
%PY% -m PyInstaller --noconfirm --clean --onefile --console --name "merge_excel" --hidden-import openpyxl --hidden-import xlrd --hidden-import gui merge_excel.py
if errorlevel 1 goto :fail

echo.
echo 打包完成，两个文件都在 dist 文件夹里。
echo exe 放在哪个文件夹，双击后默认就扫描那个文件夹，结果写到它旁边的 output。
echo.
pause
exit /b 0

:no_python
echo 没找到 Python，请先安装 Python 3.11 或更高版本，安装时勾选 Add to PATH。
pause
exit /b 1

:fail
echo.
echo 打包失败，请把上面的报错信息复制出来看看。
pause
exit /b 1
