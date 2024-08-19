@echo off
echo 開始打包 cylinder.py...

rem 檢查是否已安裝 PyInstaller，如果沒有，則安裝
pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo 安裝 PyInstaller...
    pip install pyinstaller
)

rem 使用 PyInstaller 打包
pyinstaller --onefile --windowed cylinder.py

rem 清理多餘的文件夾和文件
rd /s /q build
rd /s /q __pycache__
del cylinder.spec

echo 打包完成！
pause
