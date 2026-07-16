# Prompt para agente LOCAL (copiar y pegar)

Usa este texto completo en un **nuevo chat de Cursor en tu PC (modo local, NO Cloud Agent)**.

---

## PROMPT INICIO — copiar desde aquí

```
Eres un agente de Cursor ejecutándose EN MI PC LOCAL (Lenovo). Tienes acceso a mis archivos locales.

## Objetivo del proyecto

Completar y ajustar un libro Excel único (.xlsm) para seguimiento y valoración nutricional/antropométrica de ~200 deportistas del programa UCAD. El libro ya tiene una primera versión generada; debes refinarla usando mis archivos de referencia LOCALES.

## Repositorio

- Repo: ederacosta/ederacosta.github.io
- Rama con el trabajo: cursor/excel-seguimiento-nutricional-38bb
- Carpeta del proyecto Excel: excel-nutricion/
- PR existente: #2

Si no tienes la rama, ejecuta:
git fetch origin
git checkout cursor/excel-seguimiento-nutricional-38bb

## Archivos de referencia LOCALES (léelos tú directamente)

1. PDF antropometría ISAK Metry:
   C:\Users\LENOVO\Desktop\Isakmetry.pdf

2. Excel identificación pacientes RIPS UCAD:
   C:\Users\LENOVO\Desktop\Nueva carpeta\UCAD\pacientes\2026\JUNIO 2026\RIPS JUNIO 2026.xlsx

3. PDF informe de valoración (si existe en Desktop):
   C:\Users\LENOVO\Desktop\7a23fca2-7407-4cc3-9e13-08b0f7c42054-VALORACIONES_BLANCO_5C_ANDRES_OCT.pdf

IMPORTANTE: Abre y lee estos archivos desde mi disco. No improvises la estructura; mapea columnas y diseño reales.

## Lo que YA está implementado (no rehacer desde cero)

En excel-nutricion/ existe:

- Seguimiento_Nutricional_Deportivo.xlsx — libro generado
- build_workbook.py — script que regenera el Excel
- isak_config.py — variables ISAK y columnas RIPS
- vba/ModuloSeguimiento.bas — macros VBA
- vba/Hoja_Seguimiento.cls — evento de bloqueo en hoja Seguimiento
- Manual_Usuario.md, Admin_Config.md, README.md

### Hojas del libro actual

| Hoja | Función |
|---|---|
| Inicio | Guía rápida |
| Deportistas | ~200 filas, campos RIPS/UCAD (inferidos, no validados con mi Excel real) |
| Config_ISAK | 17 variables ISAK Nivel 1 |
| Captura_ISAK | Formulario con 3 repeticiones, media/mediana, alertas 5%/1% |
| Seguimiento | Historial longitudinal, una fila = una evaluación |
| Informe | Vista imprimible + gráficas hasta fecha de corte → Guardar como PDF manual |
| Auditoria | Log de desbloqueos admin |
| Listas | Listas desplegables (oculta) |

### Decisiones de negocio confirmadas

1. UN solo Excel para todo (~200 deportistas).
2. Seguimiento en formato largo: cada fila = una evaluación con SU propia fecha (no fecha global).
3. Fecha de evaluación se asigna automáticamente al día de captura y se bloquea (VBA).
4. Correcciones permitidas con contraseña de administrador (default: UCAD2026).
5. Informe en hoja Excel; el usuario exporta con Archivo → Guardar como → PDF (NO macro de PDF obligatoria).
6. Variables antropométricas según protocolo ISAK / Isakmetry.
7. Datos de identificación según RIPS JUNIO 2026.xlsx.

## Tu tarea principal

1. LEER Isakmetry.pdf y mapear TODAS las variables, secciones y layout exactos.
2. LEER RIPS JUNIO 2026.xlsx y mapear TODAS las columnas de identificación de pacientes/deportistas.
3. AJUSTAR build_workbook.py, isak_config.py y el Excel generado para que coincidan con esos archivos reales.
4. AJUSTAR hoja Captura_ISAK para replicar visualmente el formulario Isakmetry.
5. AJUSTAR hoja Informe para parecerse al PDF de valoración (si existe el PDF de Andrés).
6. EMBEBER o verificar que las macros VBA funcionen en .xlsm (registro, bloqueo fecha, modo admin).
7. Probar con datos ficticios + al menos 1 deportista real del RIPS si es posible.
8. Actualizar Manual_Usuario.md si cambia el flujo.

## Requisitos técnicos

- Excel 365 o 2021 (funciones LET, FILTER, SORT, MAXIFS).
- Formato .xlsm con macros VBA para bloqueo real de fechas.
- ~200 deportistas en tabla maestra Deportistas.
- Hasta 17+ variables ISAK con 3 repeticiones en captura.
- Gráficas de progreso filtradas por deportista y fecha de corte.

## Flujo de uso esperado

Captura_ISAK → botón "Registrar evaluación de hoy" → Seguimiento (fecha bloqueada)
Informe → elegir deportista + fecha de corte → Guardar como PDF

## Orden de trabajo sugerido

1. Inspeccionar archivos locales de referencia.
2. Documentar diferencias entre lo implementado y lo real.
3. Actualizar isak_config.py y build_workbook.py.
4. Regenerar: python build_workbook.py && python verify_workbook.py
5. Importar/verificar VBA en Excel local.
6. Ajustar Informe según PDF de valoración.
7. Commit, push y actualizar PR #2.

## Restricciones

- No eliminar la arquitectura de hojas sin justificación.
- Cambios mínimos y enfocados; no sobre-ingenierizar.
- Responder en español.
- Si una columna del RIPS no aplica a deportistas, déjala pero documenta por qué.

## Entregables esperados

- Seguimiento_Nutricional_Deportivo.xlsm funcional con macros
- build_workbook.py actualizado
- Manual actualizado
- Resumen de cambios vs. archivos de referencia

Empieza leyendo los 2 archivos locales obligatorios y dime qué columnas y secciones encontraste antes de modificar código.
```

## PROMPT FIN — copiar hasta aquí

---

## Cómo usarlo

1. Abre **Cursor Desktop** en tu Lenovo (no Cloud Agent).
2. Abre la carpeta del repositorio `ederacosta.github.io`.
3. Cambia a la rama `cursor/excel-seguimiento-nutricional-38bb`.
4. Crea un **nuevo chat**.
5. Verifica que el agente sea **local** (no nube).
6. Copia todo el bloque entre "PROMPT INICIO" y "PROMPT FIN".
7. Pégalo y envía.

## Tip adicional

Si el agente local tampoco lee las rutas, arrastra los archivos al chat y añade:

```
Archivos adjuntos:
- Isakmetry.pdf
- RIPS JUNIO 2026.xlsx
```
