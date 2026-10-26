@echo off
echo ========================================================
echo   Configurando entorno de Resolucion Estadistica UPT
echo ========================================================
python -m pip install --upgrade pip
pip install -r requirements.txt
playwright install chromium
echo.
echo ========================================================
echo   Entorno listo! Ya puedes entregar un Word a la IA.
echo ========================================================
pause
