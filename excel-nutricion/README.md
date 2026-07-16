# Excel — Seguimiento nutricional y antropométrico deportivo

Libro Excel único para registro de ~200 deportistas (datos RIPS/UCAD), captura ISAK Nivel 1, seguimiento longitudinal con fechas bloqueadas e informe exportable a PDF.

## Generar el libro

```bash
cd excel-nutricion
pip install -r requirements.txt
python3 build_workbook.py
python3 verify_workbook.py
```

## Archivos principales

- `Seguimiento_Nutricional_Deportivo.xlsx` — libro generado
- `vba/ModuloSeguimiento.bas` — macros VBA
- `Manual_Usuario.md` — guía de uso
- `Admin_Config.md` — contraseña admin y configuración

## Importar macros

Ver `Manual_Usuario.md`. Resumen: importar `ModuloSeguimiento.bas`, pegar evento en hoja Seguimiento, guardar como `.xlsm`.

Contraseña administrador por defecto: `UCAD2026`
