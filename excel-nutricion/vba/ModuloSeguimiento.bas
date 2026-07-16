Attribute VB_Name = "ModuloSeguimiento"
' Seguimiento Nutricional Deportivo - UCAD
' Importar: Alt+F11 > Archivo > Importar archivo
' Asignar macros: RegistrarEvaluacionHoy, ModoAdministrador, ActualizarInforme

Option Explicit

Public Const ADMIN_PASSWORD As String = "UCAD2026"
Public Const CAPTURA_SHEET As String = "Captura_ISAK"
Public Const SEGUIMIENTO_SHEET As String = "Seguimiento"
Public Const AUDITORIA_SHEET As String = "Auditoria"
Public Const INFORME_SHEET As String = "Informe"
Public Const CAPTURA_START_ROW As Long = 12
Public Const CAPTURA_VAR_COUNT As Long = 17

Public Sub RegistrarEvaluacionHoy()
    Dim wsC As Worksheet
    Dim wsS As Worksheet
    Dim tbl As ListObject
    Dim depSel As String
    Dim depId As String
    Dim newRow As ListRow
    Dim i As Long
    Dim hasData As Boolean
    Dim fechaHoy As Date
    Dim rowNum As Long
    Dim colBloqueado As Long
    Dim colIdEval As Long

    On Error GoTo ErrHandler

    Set wsC = ThisWorkbook.Worksheets(CAPTURA_SHEET)
    Set wsS = ThisWorkbook.Worksheets(SEGUIMIENTO_SHEET)
    Set tbl = wsS.ListObjects("TablaSeguimiento")

    depSel = Trim(CStr(wsC.Range("B3").Value))
    If depSel = "" Then
        MsgBox "Seleccione un deportista.", vbExclamation, "Captura ISAK"
        Exit Sub
    End If

    depId = ExtractDeportistaId(depSel)
    If depId = "" Then
        MsgBox "No se pudo identificar al deportista.", vbExclamation, "Captura ISAK"
        Exit Sub
    End If

    hasData = False
    For i = 0 To CAPTURA_VAR_COUNT - 1
        If IsNumeric(wsC.Cells(CAPTURA_START_ROW + i, 4).Value) _
            Or IsNumeric(wsC.Cells(CAPTURA_START_ROW + i, 5).Value) _
            Or IsNumeric(wsC.Cells(CAPTURA_START_ROW + i, 6).Value) Then
            hasData = True
            Exit For
        End If
    Next i

    If Not hasData Then
        MsgBox "Ingrese al menos una medida antropométrica.", vbExclamation, "Captura ISAK"
        Exit Sub
    End If

    fechaHoy = Date
    If EvaluacionExiste(depId, fechaHoy) Then
        MsgBox "Ya existe una evaluación para este deportista hoy." & vbCrLf & _
            "Use Modo administrador para corregir.", vbExclamation, "Captura ISAK"
        Exit Sub
    End If

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Set newRow = tbl.ListRows.Add
    rowNum = newRow.Range.Row

    tbl.ListColumns("ID_Deportista").DataBodyRange.Rows(tbl.ListRows.Count).Value = depId
    tbl.ListColumns("Nombre_Completo").DataBodyRange.Rows(tbl.ListRows.Count).Formula = _
        "=IFERROR(INDEX(TablaDeportistas[primerApellido],MATCH(""" & depId & """," & _
        "TablaDeportistas[ID_Deportista],0))&"" ""&INDEX(TablaDeportistas[nombres],MATCH(""" & depId & """," & _
        "TablaDeportistas[ID_Deportista],0)),"""")"

    tbl.ListColumns("Fecha_Evaluacion").DataBodyRange.Rows(tbl.ListRows.Count).Value = fechaHoy
    tbl.ListColumns("Fecha_Evaluacion").DataBodyRange.Rows(tbl.ListRows.Count).NumberFormat = "dd/mm/yyyy"

    tbl.ListColumns("Hora_Evaluacion").DataBodyRange.Rows(tbl.ListRows.Count).Value = Format(Time, "HH:MM")
    tbl.ListColumns("Antropometrista").DataBodyRange.Rows(tbl.ListRows.Count).Value = wsC.Range("E5").Value
    tbl.ListColumns("Anotador").DataBodyRange.Rows(tbl.ListRows.Count).Value = wsC.Range("E6").Value
    tbl.ListColumns("Lado_Medido").DataBodyRange.Rows(tbl.ListRows.Count).Value = wsC.Range("E7").Value
    tbl.ListColumns("Notas").DataBodyRange.Rows(tbl.ListRows.Count).Value = wsC.Range("D8").Value

    Dim codes As Variant
    codes = Array("MASA", "TALLA", "PL_TRICEPS", "PL_SUBESCAPULAR", "PL_BICEPS", _
        "PL_CRESTA_ILIACA", "PL_SUPRAESPINAL", "PL_ABDOMINAL", "PL_MUSLO_ANT", "PL_PIERNA_MED", _
        "PR_BRAZO_REL", "PR_BRAZO_FLEX", "PR_CINTURA", "PR_CADERA", "PR_PANTORRILLA", _
        "D_HUMERO", "D_FEMUR")

    For i = 0 To CAPTURA_VAR_COUNT - 1
        tbl.ListColumns(codes(i)).DataBodyRange.Rows(tbl.ListRows.Count).Value = _
            wsC.Cells(CAPTURA_START_ROW + i, 7).Value
    Next i

    ' Copiar fórmulas derivadas de fila plantilla (fila 2)
    If tbl.ListRows.Count > 1 Then
        wsS.Range("Z2:AD2").Copy
        wsS.Range("Z" & rowNum).PasteSpecial xlPasteFormulas
        Application.CutCopyMode = False
    End If

    tbl.ListColumns("Bloqueado").DataBodyRange.Rows(tbl.ListRows.Count).Value = "Sí"
    tbl.ListColumns("ID_Evaluacion").DataBodyRange.Rows(tbl.ListRows.Count).Formula = _
        "=""" & depId & """&""-""&TEXT(C" & rowNum & ",""yyyymmdd"")"

    For i = 0 To CAPTURA_VAR_COUNT - 1
        wsC.Cells(CAPTURA_START_ROW + i, 4).ClearContents
        wsC.Cells(CAPTURA_START_ROW + i, 5).ClearContents
        wsC.Cells(CAPTURA_START_ROW + i, 6).ClearContents
    Next i

    wsC.Range("E3").Value = Date

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Evaluación registrada para " & depId & " (" & Format(fechaHoy, "dd/mm/yyyy") & ").", _
        vbInformation, "Captura ISAK"
    Exit Sub

ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Error al registrar: " & Err.Description, vbCritical, "Captura ISAK"
End Sub

Private Function ExtractDeportistaId(depSel As String) As String
    Dim pos As Long
    pos = InStr(depSel, " -")
    If pos > 0 Then
        ExtractDeportistaId = Trim(Left(depSel, pos - 1))
    Else
        ExtractDeportistaId = Trim(depSel)
    End If
End Function

Private Function EvaluacionExiste(depId As String, fecha As Date) As Boolean
    Dim tbl As ListObject
    Dim i As Long
    Dim f As Variant

    Set tbl = ThisWorkbook.Worksheets(SEGUIMIENTO_SHEET).ListObjects("TablaSeguimiento")
    EvaluacionExiste = False

    For i = 1 To tbl.ListRows.Count
        If Trim(CStr(tbl.ListColumns("ID_Deportista").DataBodyRange.Cells(i, 1).Value)) = depId Then
            f = tbl.ListColumns("Fecha_Evaluacion").DataBodyRange.Cells(i, 1).Value
            If IsDate(f) Then
                If CDate(f) = fecha Then
                    EvaluacionExiste = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Public Sub Worksheet_Change_Seguimiento(ByVal Target As Range)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rowNum As Long
    Dim fechaEval As Date
    Dim bloqueado As String
    Dim colFecha As Long
    Dim colBloq As Long

    Set ws = ThisWorkbook.Worksheets(SEGUIMIENTO_SHEET)
    Set tbl = ws.ListObjects("TablaSeguimiento")

    If tbl.DataBodyRange Is Nothing Then Exit Sub
    If Intersect(Target, tbl.DataBodyRange) Is Nothing Then Exit Sub

    Application.EnableEvents = False

    rowNum = Target.Row
    colFecha = tbl.ListColumns("Fecha_Evaluacion").Range.Column
    colBloq = tbl.ListColumns("Bloqueado").Range.Column

    If Target.Column = colFecha Then
        Application.Undo
        MsgBox "La fecha de evaluación no se puede modificar.", vbExclamation
        GoTo CleanExit
    End If

    bloqueado = CStr(ws.Cells(rowNum, colBloq).Value)
    If IsDate(ws.Cells(rowNum, colFecha).Value) Then
        fechaEval = CDate(ws.Cells(rowNum, colFecha).Value)
        If (UCase(bloqueado) = "SÍ" Or UCase(bloqueado) = "SI") And fechaEval < Date Then
            Application.Undo
            MsgBox "Evaluación bloqueada. Use Modo administrador.", vbExclamation
        End If
    End If

CleanExit:
    Application.EnableEvents = True
End Sub

Public Sub ModoAdministrador()
    Dim pwd As String
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim detalle As String
    Dim idEval As String

    pwd = InputBox("Contraseña de administrador:", "Modo administrador")
    If pwd = "" Then Exit Sub
    If pwd <> ADMIN_PASSWORD Then
        MsgBox "Contraseña incorrecta.", vbCritical
        Exit Sub
    End If

    Set ws = ActiveSheet
    If ws.Name <> SEGUIMIENTO_SHEET Then
        MsgBox "Seleccione una fila en la hoja Seguimiento.", vbExclamation
        Exit Sub
    End If

    Set tbl = ws.ListObjects("TablaSeguimiento")
    If Intersect(Selection, tbl.DataBodyRange) Is Nothing Then
        MsgBox "Seleccione una fila de evaluación.", vbExclamation
        Exit Sub
    End If

    detalle = InputBox("Motivo de la corrección:", "Auditoría", "Corrección de datos")
    idEval = CStr(ws.Cells(Selection.Row, tbl.ListColumns("ID_Evaluacion").Range.Column).Value)

    ws.Cells(Selection.Row, tbl.ListColumns("Bloqueado").Range.Column).Value = "No"
    Call RegistrarAuditoria("DESBLOQUEO", idEval, detalle)

    MsgBox "Fila desbloqueada. Corrija y cambie Bloqueado a Sí al terminar.", vbInformation
End Sub

Private Sub RegistrarAuditoria(accion As String, idEval As String, detalle As String)
    Dim tbl As ListObject
    Dim lr As ListRow

    Set tbl = ThisWorkbook.Worksheets(AUDITORIA_SHEET).ListObjects("TablaAuditoria")
    Set lr = tbl.ListRows.Add
    lr.Range.Cells(1, 1).Value = Now
    lr.Range.Cells(1, 2).Value = Environ("USERNAME")
    lr.Range.Cells(1, 3).Value = accion
    lr.Range.Cells(1, 4).Value = idEval
    lr.Range.Cells(1, 5).Value = detalle
End Sub

Public Sub ActualizarInforme()
    Application.CalculateFull
    ThisWorkbook.Worksheets(INFORME_SHEET).Calculate
    MsgBox "Informe actualizado. Use Vista previa de impresión y Guardar como PDF.", vbInformation
End Sub

Public Sub Auto_Open()
    ' Habilitar macros al abrir
End Sub
