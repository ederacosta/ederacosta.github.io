#!/usr/bin/env python3
"""Genera el libro Excel de seguimiento antropométrico ISAK + RIPS."""

from __future__ import annotations

import zipfile
from copy import copy
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.chart import LineChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

from isak_config import (
    ACTIVO_LIST,
    ADMIN_PASSWORD,
    DEPORTISTA_COLUMNS,
    ISAK_VARIABLES,
    LADO_LIST,
    METODO_LIST,
    SAMPLE_DEPORTISTAS,
    SAMPLE_SEGUIMIENTO,
    SEXO_LIST,
    SI_NO_LIST,
    TIPO_DOC_LIST,
    TIPO_USUARIO_LIST,
    ZONA_LIST,
)

ROOT = Path(__file__).resolve().parent
OUTPUT_XLSX = ROOT / "Seguimiento_Nutricional_Deportivo.xlsx"
OUTPUT_XLSM = ROOT / "Seguimiento_Nutricional_Deportivo.xlsm"
VBA_DIR = ROOT / "vba"

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
SUBHEADER_FILL = PatternFill("solid", fgColor="D9E2F3")
ACCENT_FILL = PatternFill("solid", fgColor="E2EFDA")
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")
LOCKED_FILL = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="B4B4B4")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def define_name(wb: Workbook, name: str, ref: str):
    wb.defined_names.add(DefinedName(name, attr_text=ref))


def style_range(ws, cell_range: str, fill=None, font=None, border=None, alignment=None):
    for row in ws[cell_range]:
        for cell in row:
            if fill:
                cell.fill = fill
            if font:
                cell.font = font
            if border:
                cell.border = border
            if alignment:
                cell.alignment = alignment


def set_col_widths(ws, widths: dict[int, float]):
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width


def add_list_validation(ws, cell: str, formula: str, prompt: str = ""):
    dv = DataValidation(type="list", formula1=f'"{formula}"', allow_blank=True)
    if prompt:
        dv.prompt = prompt
        dv.showInputMessage = True
    ws.add_data_validation(dv)
    dv.add(cell)


def build_listas_sheet(wb: Workbook):
    ws = wb.create_sheet("Listas")
    ws.sheet_state = "hidden"
    lists = {
        "A": TIPO_DOC_LIST.split(","),
        "B": TIPO_USUARIO_LIST.split(","),
        "C": SEXO_LIST.split(","),
        "D": ZONA_LIST.split(","),
        "E": SI_NO_LIST.split(","),
        "F": ACTIVO_LIST.split(","),
        "G": LADO_LIST.split(","),
        "H": METODO_LIST.split(","),
    }
    for col_letter, values in lists.items():
        for i, value in enumerate(values, start=1):
            ws[f"{col_letter}{i}"] = value
    ws["J1"] = "DeportistasActivos"
    j_row = 1
    for row in SAMPLE_DEPORTISTAS:
        if row[-1] == "Sí":
            j_row += 1
            ws[f"J{j_row}"] = f"{row[0]} - {row[3]} {row[5]} ({row[2]})"
    return ws


def build_config_sheet(wb: Workbook):
    ws = wb.create_sheet("Config_ISAK")
    headers = [
        "Codigo",
        "Nombre",
        "Unidad",
        "Categoria",
        "Tolerancia_Pct",
        "Incluir_Grafica",
        "Orden",
    ]
    ws.append(headers)
    style_range(ws, "A1:G1", fill=HEADER_FILL, font=HEADER_FONT, border=BORDER)
    for idx, var in enumerate(ISAK_VARIABLES, start=1):
        ws.append([var[0], var[1], var[2], var[3], var[4], "Sí" if var[5] else "No", idx])
    tab = Table(displayName="TablaConfigISAK", ref=f"A1:G{len(ISAK_VARIABLES)+1}")
    tab.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(tab)
    set_col_widths(ws, {1: 18, 2: 34, 3: 8, 4: 12, 5: 14, 6: 14, 7: 8})
    ws.freeze_panes = "A2"
    return ws


def build_deportistas_sheet(wb: Workbook):
    ws = wb.create_sheet("Deportistas")
    headers = [c[0] for c in DEPORTISTA_COLUMNS]
    ws.append(headers)
    style_range(ws, f"A1:{get_column_letter(len(headers))}1", fill=HEADER_FILL, font=HEADER_FONT, border=BORDER)
    for row in SAMPLE_DEPORTISTAS:
        ws.append(list(row))
    last_row = len(SAMPLE_DEPORTISTAS) + 1
    # Filas vacías para llegar a ~200
    for i in range(len(SAMPLE_DEPORTISTAS) + 1, 201):
        ws.append([f"DEP-{i:03d}"] + [""] * (len(headers) - 1))

    tab = Table(displayName="TablaDeportistas", ref=f"A1:{get_column_letter(len(headers))}201")
    tab.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(tab)

    # Validaciones en primeras filas de datos (se propagan manualmente en Excel)
    for r in range(2, 202):
        add_list_validation(ws, f"B{r}", TIPO_DOC_LIST)
        add_list_validation(ws, f"G{r}", TIPO_USUARIO_LIST)
        add_list_validation(ws, f"I{r}", SEXO_LIST)
        add_list_validation(ws, f"L{r}", ZONA_LIST)
        add_list_validation(ws, f"M{r}", "SI,NO")
        add_list_validation(ws, f"X{r}", ACTIVO_LIST)

    set_col_widths(ws, {i: 16 for i in range(1, len(headers) + 1)})
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["F"].width = 20
    ws.freeze_panes = "A2"
    return ws


def seguimiento_headers():
    base = [
        "ID_Deportista",
        "Nombre_Completo",
        "Fecha_Evaluacion",
        "Hora_Evaluacion",
        "Antropometrista",
        "Anotador",
        "Lado_Medido",
        "Notas",
    ]
    base.extend([v[0] for v in ISAK_VARIABLES])
    base.extend(["IMC", "Suma_6_Pliegues", "Suma_8_Pliegues", "Bloqueado", "ID_Evaluacion"])
    return base


def build_seguimiento_sheet(wb: Workbook):
    ws = wb.create_sheet("Seguimiento")
    headers = seguimiento_headers()
    ws.append(headers)
    style_range(ws, f"A1:{get_column_letter(len(headers))}1", fill=HEADER_FILL, font=HEADER_FONT, border=BORDER)

    masa_col = headers.index("MASA") + 1
    talla_col = headers.index("TALLA") + 1
    pl_cols = [headers.index(v[0]) + 1 for v in ISAK_VARIABLES if v[3] == "pliegue"]

    for i, row in enumerate(SAMPLE_SEGUIMIENTO, start=2):
        values = {
            "ID_Deportista": row[0],
            "Fecha_Evaluacion": row[1],
            "MASA": row[2],
            "TALLA": row[3],
            "PL_TRICEPS": row[4],
            "PL_SUBESCAPULAR": row[5],
            "PL_BICEPS": row[6],
            "PL_CRESTA_ILIACA": row[7],
            "PL_SUPRAESPINAL": row[8],
            "PL_ABDOMINAL": row[9],
            "PL_MUSLO_ANT": row[10],
            "PL_PIERNA_MED": row[11],
            "PR_BRAZO_REL": row[12],
            "PR_BRAZO_FLEX": row[13],
            "PR_CINTURA": row[14],
            "PR_CADERA": row[15],
            "PR_PANTORRILLA": row[16],
            "D_HUMERO": row[17],
            "D_FEMUR": row[18],
            "Antropometrista": row[19],
            "Bloqueado": row[20],
            "Hora_Evaluacion": "08:30",
            "Lado_Medido": "D",
        }
        line = []
        for h in headers:
            if h == "Nombre_Completo":
                line.append(
                    f'=IFERROR(INDEX(TablaDeportistas[primerApellido],MATCH(A{i},TablaDeportistas[ID_Deportista],0))&" "&INDEX(TablaDeportistas[nombres],MATCH(A{i},TablaDeportistas[ID_Deportista],0)),"")'
                )
            elif h == "IMC":
                mc = get_column_letter(masa_col)
                tc = get_column_letter(talla_col)
                line.append(f'=IF(OR({mc}{i}="",{tc}{i}=""),"",ROUND({mc}{i}/(({tc}{i}/100)^2),2))')
            elif h == "Suma_6_Pliegues":
                # triceps, subescapular, supraespinal, abdominal, muslo ant, pierna med
                refs = [
                    get_column_letter(headers.index("PL_TRICEPS") + 1),
                    get_column_letter(headers.index("PL_SUBESCAPULAR") + 1),
                    get_column_letter(headers.index("PL_SUPRAESPINAL") + 1),
                    get_column_letter(headers.index("PL_ABDOMINAL") + 1),
                    get_column_letter(headers.index("PL_MUSLO_ANT") + 1),
                    get_column_letter(headers.index("PL_PIERNA_MED") + 1),
                ]
                parts = "+".join([f'IF({c}{i}="",0,{c}{i})' for c in refs])
                line.append(f"=IF({get_column_letter(masa_col)}{i}=\"\",\"\",ROUND({parts},1))")
            elif h == "Suma_8_Pliegues":
                parts = "+".join(
                    [f'IF({get_column_letter(c)}{i}="",0,{get_column_letter(c)}{i})' for c in pl_cols]
                )
                line.append(f"=IF({get_column_letter(masa_col)}{i}=\"\",\"\",ROUND({parts},1))")
            elif h == "ID_Evaluacion":
                line.append(f"=IF(A{i}=\"\",\"\",A{i}&\"-\"&TEXT(C{i},\"yyyymmdd\"))")
            else:
                line.append(values.get(h, ""))
        ws.append(line)

    last_row = max(500, len(SAMPLE_SEGUIMIENTO) + 50)
    tab = Table(displayName="TablaSeguimiento", ref=f"A1:{get_column_letter(len(headers))}{last_row}")
    tab.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium4",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(tab)

    # Fórmulas para filas vacías restantes
    for r in range(len(SAMPLE_SEGUIMIENTO) + 2, last_row + 1):
        ws[f"B{r}"] = (
            f'=IFERROR(INDEX(TablaDeportistas[primerApellido],MATCH(A{r},TablaDeportistas[ID_Deportista],0))&" "&INDEX(TablaDeportistas[nombres],MATCH(A{r},TablaDeportistas[ID_Deportista],0)),"")'
        )
        mc = get_column_letter(masa_col)
        tc = get_column_letter(talla_col)
        ws[f"{get_column_letter(headers.index('IMC')+1)}{r}"] = (
            f'=IF(OR({mc}{r}="",{tc}{r}=""),"",ROUND({mc}{r}/(({tc}{r}/100)^2),2))'
        )
        refs6 = [
            get_column_letter(headers.index("PL_TRICEPS") + 1),
            get_column_letter(headers.index("PL_SUBESCAPULAR") + 1),
            get_column_letter(headers.index("PL_SUPRAESPINAL") + 1),
            get_column_letter(headers.index("PL_ABDOMINAL") + 1),
            get_column_letter(headers.index("PL_MUSLO_ANT") + 1),
            get_column_letter(headers.index("PL_PIERNA_MED") + 1),
        ]
        parts6 = "+".join([f'IF({c}{r}="",0,{c}{r})' for c in refs6])
        ws[f"{get_column_letter(headers.index('Suma_6_Pliegues')+1)}{r}"] = (
            f'=IF({mc}{r}="","",ROUND({parts6},1))'
        )
        parts8 = "+".join(
            [f'IF({get_column_letter(c)}{r}="",0,{get_column_letter(c)}{r})' for c in pl_cols]
        )
        ws[f"{get_column_letter(headers.index('Suma_8_Pliegues')+1)}{r}"] = (
            f'=IF({mc}{r}="","",ROUND({parts8},1))'
        )
        ws[f"{get_column_letter(headers.index('ID_Evaluacion')+1)}{r}"] = (
            f'=IF(A{r}="","",A{r}&"-"&TEXT(C{r},"yyyymmdd"))'
        )

    set_col_widths(ws, {1: 14, 2: 22, 3: 14, 4: 10})
    ws.freeze_panes = "A2"
    return ws


def build_captura_sheet(wb: Workbook):
    ws = wb.create_sheet("Captura_ISAK")
    ws["A1"] = "CAPTURA ANTROPOMÉTRICA ISAK NIVEL 1"
    ws["A1"].font = Font(size=16, bold=True, color="1F4E79")
    ws.merge_cells("A1:F1")

    labels = [
        ("A3", "Deportista (ID):"),
        ("A4", "Documento:"),
        ("A5", "Nombre:"),
        ("A6", "Sexo / Nacimiento:"),
        ("A7", "Deporte / Equipo:"),
        ("D3", "Fecha evaluación:"),
        ("D4", "Hora:"),
        ("D5", "Antropometrista:"),
        ("D6", "Anotador:"),
        ("D7", "Lado medido:"),
        ("A8", "Método muslo:"),
        ("D8", "Notas variaciones:"),
    ]
    for cell, text in labels:
        ws[cell] = text
        ws[cell].font = Font(bold=True)

    # Controles
    ws["B3"] = "DEP-001"
    dv_dep = DataValidation(
        type="list",
        formula1="=OFFSET(Listas!$J$1,0,0,COUNTA(Listas!$J:$J),1)",
        allow_blank=False,
    )
    dv_dep.prompt = "Seleccione deportista activo"
    dv_dep.showInputMessage = True
    ws.add_data_validation(dv_dep)
    dv_dep.add("B3")

    ws["B4"] = '=IFERROR(INDEX(TablaDeportistas[tipoDocumentoIdentificacion],MATCH(LEFT(B3,FIND(" -",B3&" -")-1),TablaDeportistas[ID_Deportista],0))&" "&INDEX(TablaDeportistas[numDocumentoIdentificacion],MATCH(LEFT(B3,FIND(" -",B3&" -")-1),TablaDeportistas[ID_Deportista],0)),"")'
    ws["B5"] = '=IFERROR(INDEX(TablaDeportistas[primerApellido],MATCH(LEFT(B3,FIND(" -",B3&" -")-1),TablaDeportistas[ID_Deportista],0))&" "&INDEX(TablaDeportistas[nombres],MATCH(LEFT(B3,FIND(" -",B3&" -")-1),TablaDeportistas[ID_Deportista],0)),"")'
    ws["B6"] = '=IFERROR(INDEX(TablaDeportistas[codSexo],MATCH(LEFT(B3,FIND(" -",B3&" -")-1),TablaDeportistas[ID_Deportista],0))&" | "&TEXT(INDEX(TablaDeportistas[fechaNacimiento],MATCH(LEFT(B3,FIND(" -",B3&" -")-1),TablaDeportistas[ID_Deportista],0)),"dd/mm/yyyy"),"")'
    ws["B7"] = '=IFERROR(INDEX(TablaDeportistas[deporte],MATCH(LEFT(B3,FIND(" -",B3&" -")-1),TablaDeportistas[ID_Deportista],0))&" / "&INDEX(TablaDeportistas[equipo],MATCH(LEFT(B3,FIND(" -",B3&" -")-1),TablaDeportistas[ID_Deportista],0)),"")'

    ws["E3"] = "=TODAY()"
    ws["E3"].number_format = "DD/MM/YYYY"
    ws["E3"].fill = LOCKED_FILL
    ws["E4"] = "=NOW()"
    ws["E4"].number_format = "HH:MM"
    ws["E5"] = "Eder Acosta"
    ws["E6"] = ""
    add_list_validation(ws, "E7", LADO_LIST)
    ws["E7"] = "D"
    add_list_validation(ws, "B8", METODO_LIST)
    ws["B8"] = "A"

    style_range(ws, "B4:B7", fill=LOCKED_FILL)
    style_range(ws, "A3:F8", border=BORDER)

    # Tabla de mediciones
    start_row = 11
    headers = ["#", "Variable", "Unidad", "Medida 1", "Medida 2", "Medida 3", "Final", "Alerta"]
    for col, h in enumerate(headers, start=1):
        c = ws.cell(row=start_row, column=col, value=h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.border = BORDER

    for i, var in enumerate(ISAK_VARIABLES, start=1):
        r = start_row + i
        ws.cell(r, 1, i)
        ws.cell(r, 2, var[1])
        ws.cell(r, 3, var[2])
        for mc in range(4, 7):
            ws.cell(r, mc).number_format = "0.0"
            ws.cell(r, mc).border = BORDER
        # Final: media o mediana ISAK
        m1, m2, m3 = f"D{r}", f"E{r}", f"F{r}"
        final_formula = (
            f'=LET(vals,FILTER({m1}:{m3},{m1}:{m3}<>""),n,ROWS(vals),'
            f'IF(n=0,"",IF(n=1,INDEX(vals,1),IF(n=2,AVERAGE(vals),MEDIAN(vals)))))'
        )
        ws.cell(r, 7, final_formula)
        ws.cell(r, 7).number_format = "0.0"
        ws.cell(r, 7).fill = ACCENT_FILL
        ws.cell(r, 7).border = BORDER
        # Alerta 5%/1%
        tol = var[4]
        alert_formula = (
            f'=IF(OR(D{r}="",E{r}=""),"",IF(ABS(E{r}-D{r})/D{r}>{tol},"Revisar 3ª medida",""))'
        )
        ws.cell(r, 8, alert_formula)
        ws.cell(r, 8).border = BORDER

    # Resaltar alertas
    alert_range = f"H{start_row+1}:H{start_row+len(ISAK_VARIABLES)}"
    ws.conditional_formatting.add(
        alert_range,
        FormulaRule(formula=[f'H{start_row+1}="Revisar 3ª medida"'], fill=WARN_FILL),
    )

    # Instrucciones y botones (texto; VBA asigna macros)
    instr_row = start_row + len(ISAK_VARIABLES) + 3
    ws.cell(instr_row, 1, "INSTRUCCIONES").font = Font(bold=True, size=12)
    ws.cell(instr_row + 1, 1, "1. Seleccione deportista y diligencie las medidas.")
    ws.cell(instr_row + 2, 1, "2. Use el botón 'Registrar evaluación de hoy' (requiere macros VBA).")
    ws.cell(instr_row + 3, 1, "3. La fecha se fija automáticamente al registrar.")
    ws.cell(instr_row + 4, 1, f"4. Contraseña administrador: ver Admin_Config.md (default: {ADMIN_PASSWORD})")

    btn_row = instr_row + 6
    ws.cell(btn_row, 1, "[ Registrar evaluación de hoy ]").font = Font(bold=True, color="FFFFFF")
    ws.cell(btn_row, 1).fill = PatternFill("solid", fgColor="548235")
    ws.cell(btn_row, 3, "[ Modo administrador ]").font = Font(bold=True, color="FFFFFF")
    ws.cell(btn_row, 3).fill = PatternFill("solid", fgColor="C55A11")

    # Named ranges for VBA
    define_name(wb, "Captura_Deportista", "Captura_ISAK!$B$3")
    define_name(wb, "Captura_Fecha", "Captura_ISAK!$E$3")
    define_name(wb, "Captura_Hora", "Captura_ISAK!$E$4")
    define_name(wb, "Captura_Antropometrista", "Captura_ISAK!$E$5")
    define_name(wb, "Captura_Anotador", "Captura_ISAK!$E$6")
    define_name(wb, "Captura_Lado", "Captura_ISAK!$E$7")
    define_name(wb, "Captura_Notas", "Captura_ISAK!$D$8")
    define_name(wb, "Captura_Metodo", "Captura_ISAK!$B$8")
    define_name(wb, "Captura_Mediciones", f"Captura_ISAK!$D${start_row+1}:$G${start_row+len(ISAK_VARIABLES)}")

    set_col_widths(ws, {1: 6, 2: 34, 3: 8, 4: 10, 5: 10, 6: 10, 7: 10, 8: 16})
    return ws, start_row


def build_auditoria_sheet(wb: Workbook):
    ws = wb.create_sheet("Auditoria")
    ws.append(["Fecha_Hora", "Usuario", "Accion", "ID_Evaluacion", "Detalle"])
    style_range(ws, "A1:E1", fill=HEADER_FILL, font=HEADER_FONT, border=BORDER)
    tab = Table(displayName="TablaAuditoria", ref="A1:E200")
    tab.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(tab)
    set_col_widths(ws, {1: 20, 2: 18, 3: 22, 4: 18, 5: 40})
    return ws


def build_informe_sheet(wb: Workbook):
    ws = wb.create_sheet("Informe")
    ws["A1"] = "INFORME DE VALORACIÓN ANTROPOMÉTRICA"
    ws["A1"].font = Font(size=18, bold=True, color="1F4E79")
    ws.merge_cells("A1:H1")

    # Controles (fuera del área de impresión)
    ws["A3"] = "Deportista:"
    ws["A3"].font = Font(bold=True)
    ws["B3"] = "DEP-001 - García Andrés (1020304050)"
    dv = DataValidation(
        type="list",
        formula1="=OFFSET(Listas!$J$1,0,0,COUNTA(Listas!$J:$J),1)",
        allow_blank=False,
    )
    ws.add_data_validation(dv)
    dv.add("B3")

    ws["D3"] = "Fecha de corte:"
    ws["D3"].font = Font(bold=True)
    ws["E3"] = "=TODAY()"
    ws["E3"].number_format = "DD/MM/YYYY"

    ws["G3"] = "[ Actualizar informe ]"
    ws["G3"].font = Font(bold=True, color="FFFFFF")
    ws["G3"].fill = PatternFill("solid", fgColor="1F4E79")

    # === ÁREA DE IMPRESIÓN (fila 6 en adelante) ===
    print_start = 6
    r = print_start
    ws.cell(r, 1, "VALORACIÓN ANTROPOMÉTRICA DEPORTIVA - UCAD").font = Font(size=14, bold=True)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    r += 1
    ws.cell(r, 1, "Nutrición Deportiva | Protocolo ISAK Nivel 1")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    r += 2

    id_formula = 'LEFT(B3,FIND(" -",B3&" -")-1)'

    ws.cell(r, 1, "Datos de identificación").font = Font(bold=True, size=12)
    r += 1
    info_fields = [
        ("Nombre:", f'=IFERROR(INDEX(TablaDeportistas[primerApellido],MATCH({id_formula},TablaDeportistas[ID_Deportista],0))&" "&INDEX(TablaDeportistas[segundoApellido],MATCH({id_formula},TablaDeportistas[ID_Deportista],0))&" "&INDEX(TablaDeportistas[nombres],MATCH({id_formula},TablaDeportistas[ID_Deportista],0)),"")'),
        ("Documento:", f'=IFERROR(INDEX(TablaDeportistas[tipoDocumentoIdentificacion],MATCH({id_formula},TablaDeportistas[ID_Deportista],0))&" "&INDEX(TablaDeportistas[numDocumentoIdentificacion],MATCH({id_formula},TablaDeportistas[ID_Deportista],0)),"")'),
        ("Sexo:", f'=IFERROR(INDEX(TablaDeportistas[codSexo],MATCH({id_formula},TablaDeportistas[ID_Deportista],0)),"")'),
        ("Fecha nacimiento:", f'=IFERROR(TEXT(INDEX(TablaDeportistas[fechaNacimiento],MATCH({id_formula},TablaDeportistas[ID_Deportista],0)),"dd/mm/yyyy"),"")'),
        ("Deporte:", f'=IFERROR(INDEX(TablaDeportistas[deporte],MATCH({id_formula},TablaDeportistas[ID_Deportista],0)),"")'),
        ("Equipo:", f'=IFERROR(INDEX(TablaDeportistas[equipo],MATCH({id_formula},TablaDeportistas[ID_Deportista],0)),"")'),
        ("EPS:", f'=IFERROR(INDEX(TablaDeportistas[eps],MATCH({id_formula},TablaDeportistas[ID_Deportista],0)),"")'),
        ("Informe al:", "=TEXT(E3,\"dd/mm/yyyy\")"),
    ]
    for label, formula in info_fields:
        ws.cell(r, 1, label).font = Font(bold=True)
        ws.cell(r, 2, formula)
        r += 1

    r += 1
    ws.cell(r, 1, "Valoración actual (última evaluación ≤ fecha de corte)").font = Font(bold=True, size=12)
    r += 1

    # Tabla 5 columnas estilo informe
    table_headers = ["Variable", "Valor", "Unidad", "Fecha eval.", "Antropometrista"]
    for col, h in enumerate(table_headers, start=1):
        c = ws.cell(r, col, h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.border = BORDER
    header_row = r
    r += 1

    # Datos auxiliares para última evaluación
    aux_row = 2
    ws.cell(aux_row, 10, "Aux_ID").value = f"={id_formula}"
    ws.cell(aux_row, 11, "Aux_FechaCorte").value = "=E3"
    # Última fecha <= corte usando MAXIFS
    ws.cell(aux_row, 12, "UltimaFecha").value = (
        '=IFERROR(MAXIFS(TablaSeguimiento[Fecha_Evaluacion],TablaSeguimiento[ID_Deportista],J2,'
        'TablaSeguimiento[Fecha_Evaluacion],"<="&K2),"")'
    )

    report_vars = [
        ("Masa corporal", "MASA", "kg"),
        ("Talla", "TALLA", "cm"),
        ("IMC", "IMC", "kg/m²"),
        ("Suma 6 pliegues", "Suma_6_Pliegues", "mm"),
        ("Suma 8 pliegues", "Suma_8_Pliegues", "mm"),
        ("Cintura", "PR_CINTURA", "cm"),
        ("Cadera", "PR_CADERA", "cm"),
        ("% grasa estimada*", "MASA", "—"),
    ]
    data_start = r
    for label, code, unit in report_vars:
        ws.cell(r, 1, label).border = BORDER
        if code == "MASA" and label.startswith("%"):
            # Ecuación simplificada Evans (adultos) usando suma 6 pliegues - placeholder
            val_formula = (
                '=IFERROR(ROUND(0.8*INDEX(TablaSeguimiento[Suma_6_Pliegues],'
                'MATCH(1,INDEX((TablaSeguimiento[ID_Deportista]=J2)*'
                '(TablaSeguimiento[Fecha_Evaluacion]=L2),0),0)),1)&" % (estimado)",'
                '"")'
            )
        else:
            val_formula = (
                f'=IFERROR(INDEX(TablaSeguimiento[{code}],MATCH(1,INDEX(('
                f'TablaSeguimiento[ID_Deportista]=J2)*(TablaSeguimiento[Fecha_Evaluacion]=L2),0),0)),"")'
            )
        ws.cell(r, 2, val_formula).border = BORDER
        ws.cell(r, 3, unit).border = BORDER
        ws.cell(r, 4, '=IF(L2="","",TEXT(L2,"dd/mm/yyyy"))').border = BORDER
        ws.cell(r, 5, (
            '=IFERROR(INDEX(TablaSeguimiento[Antropometrista],MATCH(1,INDEX(('
            'TablaSeguimiento[ID_Deportista]=J2)*(TablaSeguimiento[Fecha_Evaluacion]=L2),0),0)),"")'
        )).border = BORDER
        r += 1
    data_end = r - 1

    r += 1
    ws.cell(r, 1, "* % grasa estimado con ecuación simplificada; ajustar según protocolo clínico.").font = Font(italic=True, size=9)
    r += 2

    # Tabla histórica para gráficas
    ws.cell(r, 1, "Historial para curvas de progreso").font = Font(bold=True, size=12)
    r += 1
    hist_headers = ["Fecha", "Masa", "Talla", "IMC", "Suma6P", "Cintura"]
    for col, h in enumerate(hist_headers, start=1):
        c = ws.cell(r, col, h)
        c.fill = SUBHEADER_FILL
        c.font = Font(bold=True)
        c.border = BORDER
    hist_header_row = r
    r += 1
    hist_data_start = r
    for i in range(12):
        row_num = r + i
        ws.cell(row_num, 1, (
            f'=IFERROR(INDEX(SORT(FILTER(TablaSeguimiento[Fecha_Evaluacion],'
            f'(TablaSeguimiento[ID_Deportista]=J2)*(TablaSeguimiento[Fecha_Evaluacion]<=K2))),{i+1}),"")'
        ))
        ws.cell(row_num, 1).number_format = "DD/MM/YYYY"
        for col_idx, field in enumerate(["MASA", "TALLA", "IMC", "Suma_6_Pliegues", "PR_CINTURA"], start=2):
            ws.cell(row_num, col_idx, (
                f'=IF($A{row_num}="","",IFERROR(INDEX(TablaSeguimiento[{field}],'
                f'MATCH(1,INDEX((TablaSeguimiento[ID_Deportista]=$J$2)*'
                f'(TablaSeguimiento[Fecha_Evaluacion]=$A{row_num}),0),0)),""))'
            ))
    hist_data_end = hist_data_start + 11

    chart_row = hist_data_end + 3
    # Gráfico masa
    chart1 = LineChart()
    chart1.title = "Evolución masa corporal (kg)"
    chart1.y_axis.title = "kg"
    chart1.x_axis.title = "Evaluación"
    chart1.style = 10
    data1 = Reference(ws, min_col=2, min_row=hist_header_row, max_row=hist_data_end)
    cats1 = Reference(ws, min_col=1, min_row=hist_data_start, max_row=hist_data_end)
    chart1.add_data(data1, titles_from_data=True)
    chart1.set_categories(cats1)
    chart1.height = 8
    chart1.width = 16
    ws.add_chart(chart1, f"A{chart_row}")

    chart2 = LineChart()
    chart2.title = "Evolución suma 6 pliegues (mm)"
    chart2.y_axis.title = "mm"
    chart2.style = 10
    data2 = Reference(ws, min_col=5, min_row=hist_header_row, max_row=hist_data_end)
    chart2.add_data(data2, titles_from_data=True)
    chart2.set_categories(cats1)
    chart2.height = 8
    chart2.width = 16
    ws.add_chart(chart2, f"E{chart_row}")

    footer_row = chart_row + 18
    ws.cell(footer_row, 1, "Elaborado por: Nutrición Deportiva UCAD | Exportar: Archivo → Guardar como → PDF")
    ws.cell(footer_row + 1, 1, f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}")

    # Ocultar columnas auxiliares
    ws.column_dimensions["J"].hidden = True
    ws.column_dimensions["K"].hidden = True
    ws.column_dimensions["L"].hidden = True

    # Área de impresión
    ws.print_area = f"A{print_start}:H{footer_row + 1}"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.5

    define_name(wb, "Informe_Deportista", "Informe!$B$3")
    define_name(wb, "Informe_FechaCorte", "Informe!$E$3")

    set_col_widths(ws, {1: 28, 2: 22, 3: 10, 4: 14, 5: 20, 6: 12, 7: 14, 8: 12})
    return ws, print_start, footer_row


def build_instrucciones_sheet(wb: Workbook):
    ws = wb.create_sheet("Inicio")
    ws["A1"] = "Seguimiento Nutricional y Antropométrico Deportivo"
    ws["A1"].font = Font(size=20, bold=True, color="1F4E79")
    lines = [
        "",
        "Hojas del libro:",
        "• Deportistas — Registro maestro (~200) con datos RIPS/UCAD",
        "• Config_ISAK — Variables antropométricas ISAK Nivel 1",
        "• Captura_ISAK — Formulario diario con 3 repeticiones por medida",
        "• Seguimiento — Historial longitudinal (fecha bloqueada con VBA)",
        "• Informe — Vista para imprimir / Guardar como PDF",
        "• Auditoria — Registro de desbloqueos administrativos",
        "",
        "Pasos iniciales:",
        "1. Importar macros VBA desde excel-nutricion/vba/ModuloSeguimiento.bas",
        "2. Guardar como .xlsm y habilitar macros",
        "3. Asignar botones en Captura_ISAK e Informe a las macros indicadas",
        "",
        "Flujo diario:",
        "Captura_ISAK → Registrar evaluación → Seguimiento",
        "Informe → elegir deportista y fecha → Guardar como PDF",
        "",
        f"Contraseña administrador por defecto: {ADMIN_PASSWORD}",
    ]
    for i, line in enumerate(lines, start=2):
        ws.cell(i, 1, line)
    ws.column_dimensions["A"].width = 80
    return ws


def reorder_sheets(wb: Workbook):
    order = [
        "Inicio",
        "Deportistas",
        "Config_ISAK",
        "Captura_ISAK",
        "Seguimiento",
        "Informe",
        "Auditoria",
        "Listas",
    ]
    for i, name in enumerate(order):
        wb.move_sheet(name, offset=i - wb.sheetnames.index(name))


def package_xlsm(xlsx_path: Path, xlsm_path: Path, vba_path: Path):
    """Empaqueta xlsx + vba en xlsm usando plantilla mínima OLE."""
    # Crear xlsm copiando xlsx y ajustando content types
    import shutil
    import tempfile

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        with zipfile.ZipFile(xlsx_path, "r") as zin:
            zin.extractall(tmp_path)

        # Actualizar [Content_Types].xml
        ct = (tmp_path / "[Content_Types].xml").read_text(encoding="utf-8")
        if "vbaProject" not in ct:
            override = (
                '<Override PartName="/vbaProject.bin" '
                'ContentType="application/vnd.ms-office.vbaProject"/>'
            )
            ct = ct.replace("</Types>", f"{override}</Types>")
            (tmp_path / "[Content_Types].xml").write_text(ct, encoding="utf-8")

        # Relación vba en workbook
        rels_path = tmp_path / "xl" / "_rels" / "workbook.xml.rels"
        rels = rels_path.read_text(encoding="utf-8")
        if "vbaProject" not in rels:
            rels = rels.replace(
                "</Relationships>",
                '<Relationship Id="rIdVBA" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>',
            )
            rels_path.write_text(rels, encoding="utf-8")

        # Copiar vbaProject.bin compilado
        vba_bin = vba_path / "vbaProject.bin"
        if vba_bin.exists():
            shutil.copy(vba_bin, tmp_path / "xl" / "vbaProject.bin")

        with zipfile.ZipFile(xlsm_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for file in sorted(tmp_path.rglob("*")):
                if file.is_file():
                    zout.write(file, file.relative_to(tmp_path).as_posix())


def main():
    wb = Workbook()
    wb.remove(wb.active)

    build_listas_sheet(wb)
    build_config_sheet(wb)
    build_deportistas_sheet(wb)
    build_seguimiento_sheet(wb)
    build_captura_sheet(wb)
    build_auditoria_sheet(wb)
    build_informe_sheet(wb)
    build_instrucciones_sheet(wb)
    reorder_sheets(wb)

    wb.properties.title = "Seguimiento Nutricional Deportivo"
    wb.properties.creator = "UCAD / Eder Acosta"
    wb.properties.description = "Seguimiento antropométrico ISAK + identificación RIPS"

    OUTPUT_XLSX.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUTPUT_XLSX)
    print(f"Generado: {OUTPUT_XLSX}")

    vba_bin = VBA_DIR / "vbaProject.bin"
    if vba_bin.exists():
        package_xlsm(OUTPUT_XLSX, OUTPUT_XLSM, VBA_DIR)
        print(f"Generado: {OUTPUT_XLSM}")
    else:
        print("Nota: vbaProject.bin no encontrado; importar ModuloSeguimiento.bas manualmente.")


if __name__ == "__main__":
    main()
