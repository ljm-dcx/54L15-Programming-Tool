@echo off
chcp 65001 >nul
echo ====================================
echo Nordic烧录工具 - 打包为EXE
echo ====================================
echo.

echo 正在安装依赖...
pip install openpyxl pyinstaller
echo.

echo 正在打包程序...
pyinstaller ^
  --onefile ^
  --windowed ^
  --name "NordicFlashTool" ^
  --collect-data openpyxl ^
  --collect-data et_xmlfile ^
  --hidden-import openpyxl ^
  --hidden-import openpyxl.cell._writer ^
  --hidden-import et_xmlfile ^
  nordic_flash_tool.py

echo.
echo 打包完成！
echo 可执行文件位置: dist\Nordic烧录工具.exe
echo.
pause