@echo off
echo Actualizando normal.dotm en la carpeta de plantillas de Microsoft Word
echo.

set "ruta_descargas=%userprofile%\Downloads\QuicKual-main\QuicKual-main"
set "ruta_destino=%appdata%\Microsoft\Templates"
set "escritorio=%userprofile%\Desktop"

echo.
echo Moviendo o reemplazando normal.dotm en %ruta_destino%
copy /y "%ruta_descargas%\normal.dotm" "%ruta_destino%\normal.dotm"

echo.
echo Moviendo scripts de QuicKual al escritorio
move /y "%ruta_descargas%\quickual_con_voz_para_dictar(BETA).py" "%escritorio%"
move /y "%ruta_descargas%\quickual_sin_voz_para_dictar(BETA).py" "%escritorio%"

echo.
echo ¡Operación completada con éxito!

echo.
echo Eliminando la carpeta %userprofile%\Downloads\QuicKual-main
rmdir /s /q "%userprofile%\Downloads\QuicKual-main"

echo.
echo ¡Carpeta eliminada con éxito!
pause
