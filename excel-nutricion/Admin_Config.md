# Configuración de administrador

## Contraseña por defecto

```
UCAD2026
```

Cambiar en `vba/ModuloSeguimiento.bas`:

```vba
Public Const ADMIN_PASSWORD As String = "UCAD2026"
```

Tras cambiarla, reimporte el módulo VBA en Excel.

## Modo administrador

1. Vaya a la hoja **Seguimiento**.
2. Seleccione la fila de evaluación a corregir.
3. Ejecute la macro **ModoAdministrador** (botón en Captura_ISAK o `Alt+F8`).
4. Ingrese la contraseña.
5. Indique el motivo de la corrección (queda en **Auditoria**).
6. La columna **Bloqueado** pasa a `No` temporalmente.
7. Corrija los datos y vuelva a poner **Bloqueado = Sí** al terminar.

## Agregar deportistas

En la hoja **Deportistas**, complete una fila nueva de la tabla:
- `ID_Deportista`: único (ej. `DEP-042`)
- `numDocumentoIdentificacion`: documento real
- `activo`: `Sí` para que aparezca en listas desplegables

Actualice la hoja **Listas**, columna J, con el formato:

```
DEP-042 - Apellido Nombre (documento)
```

O cree una fórmula dinámica en Excel 365 que filtre deportistas activos.

## Agregar o modificar variables ISAK

Edite la hoja **Config_ISAK** y las columnas correspondientes en:
- `Captura_ISAK` (filas de medición)
- `Seguimiento` (columnas de la tabla)
- `Informe` (tabla de valoración y gráficas)

> Cambios estructurales requieren ajustar también el módulo VBA (`RegistrarEvaluacionHoy`).

## Proteger hojas

Opcional: ejecute la macro `ProtegerHojas` (si se habilita en el módulo VBA) para proteger Seguimiento con la misma contraseña de administrador.

## Ecuación de % grasa

El informe usa una estimación simplificada basada en suma de 6 pliegues. Ajuste la fórmula en la hoja **Informe** según el protocolo clínico que utilice (Evans, Slaughter, ISAK Metry, etc.).

## Respaldo

Guarde una copia del `.xlsm` periódicamente. El historial en **Seguimiento** es la fuente de verdad para informes y curvas.
