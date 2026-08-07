Attribute VB_Name = "PivotCorreosPorHora"
' ---------------------------------------------------------------------------
' Crea una tabla dinamica con el conteo de correos por hora del dia (0-23)
' a partir de la columna "Received" del CSV exportado desde Microsoft Graph.
'
' Uso:
'   1. Abrir el CSV en Excel.
'   2. Alt+F11 -> Insertar -> Modulo -> pegar este codigo.
'   3. Volver a Excel, Alt+F8 -> CrearPivotCorreosPorHora -> Ejecutar.
'
' La extraccion de la hora soporta ambos casos:
'   * Excel parseo "Received" como fecha/hora real  -> se usa Hour() directo.
'   * "Received" quedo como texto ("7/23/2026 10:16") -> se corta despues del
'     espacio y se lee la parte anterior a los dos puntos.
' ---------------------------------------------------------------------------
Option Explicit

Private Const HDR_RECEIVED  As String = "Received"
Private Const HDR_HORA      As String = "Hora"
Private Const HDR_RANGO     As String = "Rango horario"
Private Const HDR_FOLDER    As String = "Folder"
Private Const HOJA_PIVOT    As String = "Correos por hora"
Private Const NOMBRE_PIVOT  As String = "ptCorreosPorHora"

Public Sub CrearPivotCorreosPorHora()

    Dim ws As Worksheet
    Dim wb As Workbook
    Dim colRec As Long, colHora As Long, colRango As Long, colFolder As Long
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, sinHora As Long
    Dim h As Variant

    Set ws = ActiveSheet
    Set wb = ws.Parent

    colRec = BuscarColumna(ws, HDR_RECEIVED)
    If colRec = 0 Then
        MsgBox "No encontre una columna llamada '" & HDR_RECEIVED & "' en la fila 1 de la hoja '" & _
               ws.Name & "'." & vbCrLf & vbCrLf & _
               "Abri la hoja con el CSV exportado y volve a ejecutar la macro.", _
               vbExclamation, "Falta la columna Received"
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, colRec).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "La columna '" & HDR_RECEIVED & "' no tiene datos.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' --- Columnas auxiliares (se reutilizan si ya existen de una corrida previa)
    colHora = BuscarColumna(ws, HDR_HORA)
    If colHora = 0 Then
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        colHora = lastCol + 1
        ws.Cells(1, colHora).Value = HDR_HORA
    End If

    colRango = BuscarColumna(ws, HDR_RANGO)
    If colRango = 0 Then
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        colRango = lastCol + 1
        ws.Cells(1, colRango).Value = HDR_RANGO
    End If

    ' --- Rellenar hora y rango fila por fila
    sinHora = 0
    For i = 2 To lastRow
        h = ExtraerHora(ws.Cells(i, colRec).Value)
        If IsEmpty(h) Then
            ws.Cells(i, colHora).ClearContents
            ws.Cells(i, colRango).Value = "(sin hora)"
            sinHora = sinHora + 1
        Else
            ws.Cells(i, colHora).Value = CLng(h)
            ws.Cells(i, colRango).Value = Format$(h, "00") & ":00 - " & Format$(h, "00") & ":59"
        End If
    Next i

    ws.Columns(colHora).NumberFormat = "0"
    ws.Columns(colRango).NumberFormat = "@"

    ' --- Hoja de destino limpia
    Dim wsPivot As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets(HOJA_PIVOT).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set wsPivot = wb.Worksheets.Add(After:=ws)
    wsPivot.Name = HOJA_PIVOT

    ' --- Tabla dinamica
    Dim rngData As Range, pc As PivotCache, pt As PivotTable
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Set rngData = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    Set pc = wb.PivotCaches.Create(SourceType:=xlDatabase, _
                                   SourceData:="'" & ws.Name & "'!" & rngData.Address(True, True, xlA1))
    Set pt = pc.CreatePivotTable(TableDestination:=wsPivot.Range("A3"), TableName:=NOMBRE_PIVOT)

    With pt
        .PivotFields(HDR_RANGO).Orientation = xlRowField
        .PivotFields(HDR_RANGO).Position = 1

        ' Si el CSV trae la carpeta, se abre una columna por carpeta.
        colFolder = BuscarColumna(ws, HDR_FOLDER)
        If colFolder > 0 Then
            .PivotFields(HDR_FOLDER).Orientation = xlColumnField
            .PivotFields(HDR_FOLDER).Position = 1
        End If

        .AddDataField .PivotFields(HDR_RECEIVED), "Correos", xlCount
        .RowAxisLayout xlTabularRow
        .ColumnGrand = True
        .RowGrand = True
    End With

    wsPivot.Range("A1").Value = "Correos recibidos por hora del dia (formato 24h)"
    wsPivot.Range("A1").Font.Bold = True
    wsPivot.Range("A1").Font.Size = 13
    wsPivot.Columns("A:Z").AutoFit

    ' --- Grafico de barras sobre la dinamica
    Dim ch As ChartObject
    Set ch = wsPivot.ChartObjects.Add(Left:=wsPivot.Range("F3").Left, Top:=wsPivot.Range("F3").Top, _
                                      Width:=520, Height:=300)
    ch.Chart.SetSourceData Source:=pt.TableRange1
    ch.Chart.ChartType = xlColumnClustered
    ch.Chart.HasTitle = True
    ch.Chart.ChartTitle.Text = "Correos por hora"

    Application.ScreenUpdating = True
    wsPivot.Activate

    Dim msg As String
    msg = "Listo. Tabla dinamica creada en la hoja '" & HOJA_PIVOT & "'." & vbCrLf & _
          "Filas procesadas: " & (lastRow - 1)
    If sinHora > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Atencion: " & sinHora & " fila(s) sin hora reconocible " & _
              "quedaron agrupadas como '(sin hora)'."
    End If
    MsgBox msg, vbInformation, "Pivot de correos por hora"

End Sub

' ---------------------------------------------------------------------------
' Devuelve la hora (0-23) de un valor, o Empty si no se puede determinar.
' ---------------------------------------------------------------------------
Private Function ExtraerHora(ByVal v As Variant) As Variant

    Dim s As String, parteHora As String
    Dim posEspacio As Long, posDosPuntos As Long
    Dim hh As Long

    ExtraerHora = Empty

    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function

    ' Caso 1: Excel lo parseo como fecha/hora real.
    If VarType(v) = vbDate Then
        ExtraerHora = Hour(CDate(v))
        Exit Function
    End If

    ' Caso 2: numero de serie de Excel.
    If IsNumeric(v) And Not VarType(v) = vbString Then
        ExtraerHora = Hour(CDate(CDbl(v)))
        Exit Function
    End If

    ' Caso 3: texto "7/23/2026 10:16" o "2026-07-23 10:16:00".
    s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Function

    posEspacio = InStr(s, " ")
    If posEspacio = 0 Then Exit Function

    parteHora = Trim$(Mid$(s, posEspacio + 1))
    If Len(parteHora) = 0 Then Exit Function

    ' Si VBA sabe leerlo (incluido AM/PM), que lo haga el.
    If IsDate(parteHora) Then
        ExtraerHora = Hour(CDate(parteHora))
        Exit Function
    End If

    ' Ultimo recurso: lo que hay antes de los dos puntos.
    posDosPuntos = InStr(parteHora, ":")
    If posDosPuntos < 2 Then Exit Function

    parteHora = Left$(parteHora, posDosPuntos - 1)
    If Not IsNumeric(parteHora) Then Exit Function

    hh = CLng(parteHora)
    If hh >= 0 And hh <= 23 Then ExtraerHora = hh

End Function

' ---------------------------------------------------------------------------
' Numero de columna cuyo encabezado (fila 1) coincide con el nombre dado.
' Ignora mayusculas y espacios sobrantes. 0 si no existe.
' ---------------------------------------------------------------------------
Private Function BuscarColumna(ByVal ws As Worksheet, ByVal nombre As String) As Long

    Dim ultima As Long, c As Long

    BuscarColumna = 0
    ultima = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To ultima
        If StrComp(Trim$(CStr(ws.Cells(1, c).Value)), Trim$(nombre), vbTextCompare) = 0 Then
            BuscarColumna = c
            Exit Function
        End If
    Next c

End Function
