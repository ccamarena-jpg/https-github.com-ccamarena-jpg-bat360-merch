@echo off
rem Guarda un corte nuevo del Excel de Hardware en HW_CORTES.xlsx (solo si el contenido cambio).
rem Se puede ejecutar a mano o desde el Programador de tareas de Windows.
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
python guardar_corte_hw.py >> "%~dp0guardar_corte_hw.log" 2>&1
