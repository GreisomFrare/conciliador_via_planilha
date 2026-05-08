@echo off
echo ============================================================
echo  Build - Conciliacao de Cartao
echo ============================================================
echo.

REM Instala dependencias
pip install oracledb openpyxl cryptography pyinstaller --quiet

echo Gerando executavel...
echo.

pyinstaller ^
  --onefile ^
  --noconsole ^
  --name "ConciliacaoCartao" ^
  --hidden-import=oracledb ^
  --hidden-import=oracledb.thin_impl ^
  --hidden-import=openpyxl ^
  --hidden-import=openpyxl.cell._writer ^
  --add-data "config_manager.py;." ^
  --add-data "database.py;." ^
  --add-data "excel_importer.py;." ^
  --add-data "app.py;." ^
  main.py

echo.
if exist "dist\ConciliacaoCartao.exe" (
    echo [OK] Executavel gerado: dist\ConciliacaoCartao.exe
) else (
    echo [ERRO] Falha ao gerar executavel.
)
echo.
pause
