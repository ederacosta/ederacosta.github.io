# Manual de usuario — Seguimiento Nutricional Deportivo

## Archivos entregados

| Archivo | Descripción |
|---|---|
| `Seguimiento_Nutricional_Deportivo.xlsx` | Libro Excel con todas las hojas, fórmulas y datos de prueba |
| `vba/ModuloSeguimiento.bas` | Macros VBA (registro, bloqueo, informe, admin) |
| `vba/Hoja_Seguimiento.cls` | Código de evento para bloqueo en hoja Seguimiento |

## Configuración inicial (una sola vez)

1. Abra `Seguimiento_Nutricional_Deportivo.xlsx` en Excel.
2. Pulse `Alt + F11` para abrir el editor VBA.
3. Menú **Archivo → Importar archivo** e importe `vba/ModuloSeguimiento.bas`.
4. En el Explorador de proyectos, doble clic en la hoja **Seguimiento** y pegue el contenido de `vba/Hoja_Seguimiento.cls` (o importe si su Excel lo permite).
5. Guarde como **`Seguimiento_Nutricional_Deportivo.xlsm`** (libro habilitado para macros).
6. Habilite macros al abrir el archivo.
7. Asigne macros a los botones:
   - En **Captura_ISAK**, celda verde `[ Registrar evaluación de hoy ]` → macro `RegistrarEvaluacionHoy`
   - En **Captura_ISAK**, celda naranja `[ Modo administrador ]` → macro `ModoAdministrador`
   - En **Informe**, celda `[ Actualizar informe ]` → macro `ActualizarInforme`

> Para asignar macro: clic derecho en la celda → **Asignar macro…**

## Hojas del libro

### Inicio
Resumen y guía rápida.

### Deportistas
Registro maestro de ~200 deportistas con campos RIPS/UCAD:
- Identificación (tipo y número de documento, nombres, apellidos)
- Datos RIPS (tipo usuario, sexo, municipio, zona, EPS…)
- Datos deportivos (deporte, equipo, categoría)

### Config_ISAK
Las 17 variables del protocolo ISAK Nivel 1 (perfil restringido).

### Captura_ISAK
Formulario diario de antropometría:
1. Seleccione deportista en la lista desplegable.
2. Diligencie Medida 1, 2 y 3 según protocolo ISAK.
3. La columna **Final** calcula media o mediana automáticamente.
4. La columna **Alerta** avisa si necesita 3.ª medida (tolerancia 5% pliegues / 1% resto).
5. Pulse **Registrar evaluación de hoy**.

La fecha se fija al día de registro y no puede modificarse sin contraseña de administrador.

### Seguimiento
Historial longitudinal: una fila por evaluación por deportista.
Incluye IMC, suma de pliegues y estado de bloqueo.

### Informe
1. Seleccione deportista y **fecha de corte**.
2. Pulse **Actualizar informe** (o `F9` / recalcular).
3. Revise tablas y gráficas de progreso.
4. **Archivo → Guardar como → PDF** para exportar.

### Auditoria
Registro automático de desbloqueos administrativos.

## Flujo diario

```
Captura_ISAK → Registrar evaluación → Seguimiento
Informe → elegir fecha → Guardar como PDF
```

## Datos de prueba incluidos

- 5 deportistas ficticios (DEP-001 a DEP-005)
- 3 evaluaciones para DEP-001 y DEP-002; 2-3 para el resto
- Listas desplegables precargadas

## Requisitos

- Microsoft Excel 365 o Excel 2021 (funciones `LET`, `FILTER`, `SORT`, `MAXIFS`)
- Macros habilitadas para bloqueo de fechas y registro automático
