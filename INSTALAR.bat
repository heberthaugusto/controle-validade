@echo off
echo Instalando dependencias do Controle de Validade...
python -m pip install --upgrade pip
python -m pip install openpyxl python-dotenv
if not exist "%~dp0.env" copy "%~dp0.env.exemplo" "%~dp0.env"
echo @echo off> "%~dp0EXECUTAR.bat"
echo python "%%~dp0lancador_validades.py">> "%~dp0EXECUTAR.bat"
echo.
echo Instalacao concluida. Edite o arquivo .env e use EXECUTAR.bat para abrir o Lancador.
pause
