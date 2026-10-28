@echo off
chcp 65001 >nul
setlocal

cd /d "%~dp0"

echo [1/3] 检查 PyInstaller...
python -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo     未安装，正在安装 PyInstaller...
    python -m pip install pyinstaller
    if errorlevel 1 goto :fail
)

echo [2/3] 检查 openpyxl...
python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo     未安装，正在安装 openpyxl...
    python -m pip install -r requirements.txt
    if errorlevel 1 goto :fail
)

echo [3/3] 开始打包...
python -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --onefile ^
    --windowed ^
    --name "SVN日志导出" ^
    --version-file version_info.txt ^
    --paths . ^
    main.py
if errorlevel 1 goto :fail

echo.
echo 打包完成：%~dp0dist\SVN日志导出.exe
echo 注意：目标机器仍需安装 TortoiseSVN（提供 svn.exe），或把 svn.exe 放在 PATH 中。
pause
exit /b 0

:fail
echo.
echo 打包失败，请检查上方错误信息。
pause
exit /b 1
