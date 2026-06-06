Attribute VB_Name = "Rounding"
Option Explicit

Private gCutDecision As String  ' "MATERIAL" o "SO"
Private Const LIGHT_TABLE_STYLE As String = "TableStyleLight1"


'========================================================
' Módulo: ZVORDROUND_Auto
' Propósito: Forzar el FROM (pgwmordering) al crear el draft de rounding.
' Dependencias: Outlook (late bound)
'========================================================

Private Const DEFAULT_FROM_SMTP As String = "pgwmordering.im@pg.com"
'==================== PARTE 1 ====================
' Módulo: ZVORDROUND_Auto
' Propósito: Actualizar la tabla ZVORDROUND con el último Excel en Descargas (OneDrive personal) y crear hojas ASMs/Reg/Family con % y marcado rojo en tablas.
' Dependencias: Microsoft Scripting Runtime (opcional). Si no está, usa CreateObject("Scripting.Dictionary") igual funciona.


Private Function GetDraftsFolderForMailbox(ByVal olApp As Object, ByVal smtp As String) As Object
    ' Devuelve Drafts del mailbox aunque sea Shared (aunque no aparezca en Session.Accounts)
    Dim ns As Object, recip As Object, f As Object

    Set ns = olApp.Session

    On Error Resume Next
    Set recip = ns.CreateRecipient(smtp)
    recip.Resolve
    On Error GoTo 0

    If Not recip Is Nothing Then
        If recip.Resolved Then
            On Error Resume Next
            ' 16 = olFolderDrafts
            Set f = ns.GetSharedDefaultFolder(recip, 16)
            On Error GoTo 0
            If Not f Is Nothing Then
                Set GetDraftsFolderForMailbox = f
                Exit Function
            End If
        End If
    End If

    ' Fallback: Drafts del default store
    On Error Resume Next
    Set GetDraftsFolderForMailbox = ns.GetDefaultFolder(16)
    On Error GoTo 0
End Function

Private Function CreateDraftInFolder(ByVal draftsFolder As Object) As Object
    On Error Resume Next
    Set CreateDraftInFolder = draftsFolder.Items.Add("IPM.Note")
    On Error GoTo 0
End Function

Public Sub ActualizarZVORDROUND_Y_CrearVistas()

    Dim fLatest As String
    Dim wbExp As Workbook
    Dim wsData As Worksheet
    Dim loMain As ListObject
    Dim ok As Boolean
    Dim finalMsg As String

    On Error GoTo Fail

    ok = False
    finalMsg = "Proceso falló."

    '1) Encontrar el último EXPORT en Descargas
    fLatest = GetLatestExcelFromDownloads()
    If Len(Trim$(fLatest)) = 0 Then GoTo CleanExit
    If Len(Dir$(fLatest)) = 0 Then GoTo CleanExit

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    'IMPORTANTE: NO apagar DisplayAlerts aquí (para que aparezca el Sensitivity Label)

    '2) Abrir EXPORT (aquí te deja seleccionar el Sensitivity Label)
    Set wbExp = Workbooks.Open(fileName:=fLatest, ReadOnly:=False)
    If wbExp Is Nothing Then GoTo CleanExit

    '3) Ahora sí: correr "en background" (sin popups de Excel)
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    '4) Crear / recrear la tabla ZVORDROUND DENTRO DEL EXPORT (en hoja 1)
    Set loMain = CreateOrReplaceZVORDROUNDTableInWorkbook(wbExp, wsData)
    If loMain Is Nothing Then GoTo CleanExit

    '5) Limpiar formatos de toda la hoja y aplicar marcado SOLO en la tabla
    ClearListObjectRowFill loMain
    MarkMainFirst loMain
    ApplyFamilyStyle loMain
    DecideAndSortMain loMain

    '6) Ordenar tabla principal por color (rojos primero, luego verdes)
    SortTableByFontColor loMain, RGB(156, 0, 6)

    CreateViewSheet wbMain:=wbExp, sourceLO:=loMain, targetSheetName:="ASMs (R1)", shipToMode:="INCLUDE", shipToPattern1:="as", shipToPattern2:="", pctThreshold:=1#, dictOriginalRows:=Nothing, markOriginalThreshold:=1#, pctHeader:="Percentage"
    CreateViewSheet wbMain:=wbExp, sourceLO:=loMain, targetSheetName:="Reg (R1)", shipToMode:="EXCLUDE_2", shipToPattern1:="as", shipToPattern2:="HVDC", pctThreshold:=2#, dictOriginalRows:=Nothing, markOriginalThreshold:=2#, pctHeader:="Percentage"
    CreateViewSheet wbMain:=wbExp, sourceLO:=loMain, targetSheetName:="Family (HA)", shipToMode:="INCLUDE", shipToPattern1:="HVDC", shipToPattern2:="", pctThreshold:=2#, dictOriginalRows:=Nothing, markOriginalThreshold:=2#, pctHeader:="Percentage"

    '7) Draft (adjunta 1ra hoja del EXPORT, no del template)
    CreateRoundingEmailDraft wbExp

    '8) Dejarte en la primera hoja del EXPORT
    wbExp.Activate
    wbExp.Worksheets(1).Activate

    ok = True
    finalMsg = "Proceso exitoso." & vbCrLf & _
               "Archivo: " & fLatest & vbCrLf & _
               "Cut decision: " & IIf(UCase$(gCutDecision) = "MATERIAL", "Material", IIf(UCase$(gCutDecision) = "SO", "SO", "N/A"))

CleanExit:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox finalMsg, IIf(ok, vbInformation, vbCritical), "ZVORDROUND"
    Exit Sub

Fail:
    finalMsg = "Proceso falló: " & Err.Description
    Resume CleanExit

End Sub

'==================== PARTE 2 ====================
' Módulo: ZVORDROUND_Auto
' Propósito: Encontrar la tabla ZVORDROUND en el libro y devolver hoja y ListObject
' Dependencias: Ninguna

Private Function FindListObjectByName(ByVal wb As Workbook, ByVal loName As String, ByRef wsOut As Worksheet) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.name, loName, vbTextCompare) = 0 Then
                Set wsOut = ws
                Set FindListObjectByName = lo
                Exit Function
            End If
        Next lo
    Next ws

    Set FindListObjectByName = Nothing
End Function

'==================== PARTE 3 ====================
' Módulo: ZVORDROUND_Auto
' Propósito: Encontrar el Excel más reciente en Descargas (Downloads + OneDrive personal)
' Dependencias: Scripting.FileSystemObject (late bound)

Private Function GetLatestExcelFromDownloads() As String
    Dim candidates(1 To 3) As String
    Dim pUser As String
    Dim bestPath As String
    Dim bestDT As Date
    Dim i As Long

    pUser = Environ$("USERPROFILE")

    candidates(1) = pUser & "\Downloads"
    candidates(2) = pUser & "\OneDrive\Downloads"
    candidates(3) = pUser & "\OneDrive - Personal\Downloads"

    bestPath = ""
    bestDT = #1/1/1900#

    For i = 1 To 3
        If FolderExists(candidates(i)) Then
            PickLatestExcelInFolderStartingWith candidates(i), bestPath, bestDT, "EXPORT"
        End If
    Next i

    GetLatestExcelFromDownloads = bestPath
End Function

Private Sub PickLatestExcelInFolderStartingWith( _
    ByVal folderPath As String, _
    ByRef bestPath As String, _
    ByRef bestDT As Date, _
    ByVal startsWithText As String _
)
    Dim fso As Object
    Dim fol As Object
    Dim fil As Object
    Dim ext As String
    Dim dt As Date
    Dim nm As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fol = fso.GetFolder(folderPath)

    For Each fil In fol.Files
        nm = fil.name
        If LCase$(Left$(nm, Len(startsWithText))) = LCase$(startsWithText) Then
            ext = LCase$(fso.GetExtensionName(nm))
            If ext = "xlsx" Or ext = "xlsm" Or ext = "xls" Then
                dt = fil.DateLastModified
                If dt > bestDT Then
                    bestDT = dt
                    bestPath = fil.path
                End If
            End If
        End If
    Next fil
End Sub

Private Function FolderExists(ByVal folderPath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FolderExists = fso.FolderExists(folderPath)
End Function

'==================== PARTE 4 ====================
' Módulo: ZVORDROUND_Auto
' Propósito: Helpers para encabezados, detección de header, slicing de matrices 2D
' Dependencias: Ninguna

Private Function GetListObjectHeaders(ByVal lo As ListObject) As Variant
    Dim n As Long
    Dim i As Long
    Dim arr() As String

    n = lo.ListColumns.count
    ReDim arr(1 To 1, 1 To n)

    For i = 1 To n
        arr(1, i) = CStr(lo.ListColumns(i).name)
    Next i

    GetListObjectHeaders = arr
End Function

Private Function LooksLikeHeaderRow(ByVal v As Variant, ByVal mainHeaders As Variant) As Boolean
    Dim c As Long
    Dim hits As Long
    Dim maxC As Long
    Dim cellTxt As String

    hits = 0
    maxC = UBound(mainHeaders, 2)

    For c = 1 To Application.Min(UBound(v, 2), maxC)
        cellTxt = CStr(v(1, c))
        If StrComp(Trim$(cellTxt), Trim$(CStr(mainHeaders(1, c))), vbTextCompare) = 0 Then
            hits = hits + 1
        End If
    Next c

    'Si calza con la mayoría, asumimos encabezados presentes
    LooksLikeHeaderRow = (hits >= Application.Max(3, CLng(maxC * 0.5)))
End Function

Private Function Slice2D(ByVal v As Variant, ByVal r1 As Long, ByVal r2 As Long, ByVal c1 As Long, ByVal c2 As Long) As Variant
    Dim r As Long, c As Long
    Dim outV As Variant
    Dim rr As Long, cc As Long

    rr = r2 - r1 + 1
    cc = c2 - c1 + 1

    ReDim outV(1 To rr, 1 To cc)

    For r = 1 To rr
        For c = 1 To cc
            outV(r, c) = v(r1 + r - 1, c1 + c - 1)
        Next c
    Next r

    Slice2D = outV
End Function

'==================== PARTE 5 ====================
' Módulo: ZVORDROUND_Auto
' Propósito: Reemplazar data de ListObject (ZVORDROUND) con matriz y ajustar tamaño
' Dependencias: Ninguna

Private Sub ReplaceListObjectData(ByVal lo As ListObject, ByVal dataOnly As Variant, ByVal srcColCount As Long)
    Dim ws As Worksheet
    Dim headerRow As Long
    Dim startCell As Range
    Dim newRows As Long
    Dim newCols As Long
    Dim newRange As Range

    Set ws = lo.Parent
    Set startCell = lo.Range.Cells(1, 1) 'Esquina superior izquierda (incluye encabezado)
    headerRow = startCell.Row

    newRows = UBound(dataOnly, 1)
    newCols = lo.ListColumns.count

    'Si la fuente trae más/menos columnas, solo copiamos hasta lo que existe en la tabla
    If srcColCount < newCols Then newCols = srcColCount

    'Borrar data actual
    If Not lo.DataBodyRange Is Nothing Then
        lo.DataBodyRange.ClearContents
    End If

    'Redimensionar la tabla al nuevo tamaño (encabezado + filas)
    Set newRange = ws.Range(startCell, startCell.Offset(newRows, lo.ListColumns.count - 1))
    lo.Resize newRange

    'Pegar valores (solo las columnas que vamos a usar/caben)
    ws.Range(startCell.Offset(1, 0), startCell.Offset(newRows, newCols - 1)).Value = dataOnly

    'Si la tabla tiene más columnas que el archivo fuente, esas quedan en blanco en las filas nuevas
End Sub

Private Sub ClearListObjectRowFill(ByVal lo As ListObject)
    If Not lo.DataBodyRange Is Nothing Then
        With lo.DataBodyRange
            .Interior.pattern = xlNone
            .Font.ColorIndex = xlAutomatic
            .FormatConditions.Delete
        End With
    End If
End Sub

'==================== PARTE 6 ====================
' Módulo: ZVORDROUND_Auto
' Propósito: Construir diccionario key(Order Number|Item No.) -> Range de fila en tabla original
' Dependencias: Scripting.Dictionary (late bound)

Private Function BuildKeyToRowRange(ByVal lo As ListObject, ByVal colOrder As String, ByVal colItem As String) As Object
    Dim dict As Object
    Dim idxOrder As Long
    Dim idxItem As Long
    Dim r As Long
    Dim vOrder As Variant
    Dim vItem As Variant
    Dim key As String
    Dim rowRng As Range

    If lo.DataBodyRange Is Nothing Then
        Set BuildKeyToRowRange = Nothing
        Exit Function
    End If

    idxOrder = GetColumnIndex(lo, colOrder)
    idxItem = GetColumnIndex(lo, colItem)

    If idxOrder = 0 Or idxItem = 0 Then
        Set BuildKeyToRowRange = Nothing
        Exit Function
    End If

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    For r = 1 To lo.DataBodyRange.Rows.count
        vOrder = lo.DataBodyRange.Cells(r, idxOrder).Value
        vItem = lo.DataBodyRange.Cells(r, idxItem).Value

        key = CStr(vOrder) & "|" & CStr(vItem)
        Set rowRng = lo.DataBodyRange.Rows(r) 'Range de la fila dentro de la tabla

        'Guardar como objeto Range (sin error 424)
        If dict.Exists(key) Then dict.Remove key
        dict.Add key, rowRng
    Next r

    Set BuildKeyToRowRange = dict
End Function

Private Function GetColumnIndex(ByVal lo As ListObject, ByVal headerName As String) As Long
    Dim i As Long
    For i = 1 To lo.ListColumns.count
        If StrComp(lo.ListColumns(i).name, headerName, vbTextCompare) = 0 Then
            GetColumnIndex = i
            Exit Function
        End If
    Next i
    GetColumnIndex = 0
End Function

'==================== PARTE 7 ====================
' Módulo: ZVORDROUND_Auto
' Propósito: Crear hoja (ASMs/Reg/Family), filtrar por Ship To Name, insertar Percentage, marcar filas y marcar original
' Dependencias: Ninguna

Private Sub CreateViewSheet( _
    ByVal wbMain As Workbook, _
    ByVal sourceLO As ListObject, _
    ByVal targetSheetName As String, _
    ByVal shipToMode As String, _
    ByVal shipToPattern1 As String, _
    ByVal shipToPattern2 As String, _
    ByVal pctThreshold As Double, _
    ByVal dictOriginalRows As Object, _
    ByVal markOriginalThreshold As Double, _
    ByVal pctHeader As String _
)
    ' Usa el ListObject como fuente (no la hoja completa)
    CreateViewByCopyFromLO sourceLO, targetSheetName, shipToMode, shipToPattern1, shipToPattern2, pctHeader
End Sub

Private Sub CreateViewByCopyFromLO( _
    ByVal loSource As ListObject, _
    ByVal newSheetName As String, _
    ByVal shipToMode As String, _
    ByVal p1 As String, _
    ByVal p2 As String, _
    ByVal pctHeader As String _
)
    Dim wb As Workbook
    Dim wsNew As Worksheet
    Dim loNew As ListObject
    Dim srcRng As Range
    Dim usedRng As Range
    Dim idxShip As Long
    Dim r As Long
    Dim shipName As String
    Dim i As Long

    If loSource Is Nothing Then Exit Sub
    Set wb = loSource.Parent.Parent

    ' Si el workbook está protegido a nivel estructura, ningún método que agregue/borre hojas va a servir
    If wb.ProtectStructure Then Err.Raise 1001, , "Workbook structure is protected. Unprotect the workbook structure and retry."

    ' Fuente: solo el rango de la tabla (incluye headers)
    Set srcRng = loSource.Range
    If srcRng Is Nothing Then Exit Sub
    If srcRng.Rows.count < 2 Then Exit Sub

    Application.ScreenUpdating = False

    DeleteSheetIfExists wb, newSheetName

    Set wsNew = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
    wsNew.name = newSheetName

    ' Copiar ancho de columnas (solo las columnas de la tabla)
    For i = 1 To srcRng.Columns.count
        wsNew.Columns(i).ColumnWidth = srcRng.Columns(i).ColumnWidth
    Next i

    ' Copiar valores + formatos (sin Worksheet.Copy)
    srcRng.Copy
    With wsNew.Range("A1")
        .PasteSpecial xlPasteValues
        .PasteSpecial xlPasteFormats
    End With
    Application.CutCopyMode = False

    Set usedRng = wsNew.Range("A1").CurrentRegion
    If usedRng.Rows.count < 2 Then Exit Sub

    ' Crear tabla nueva
    Set loNew = wsNew.ListObjects.Add(xlSrcRange, usedRng, , xlYes)
    On Error Resume Next
    loNew.name = "ZVORDROUND_" & Replace(newSheetName, " ", "_")
    On Error GoTo 0

    ' Forzar estilo light
    On Error Resume Next
    loNew.TableStyle = LIGHT_TABLE_STYLE
    On Error GoTo 0

    idxShip = GetColumnIndex(loNew, "Ship To Name")
    If idxShip = 0 Then Exit Sub

    ' Limpiar filtros/FC
    On Error Resume Next
    If loNew.AutoFilter.FilterMode Then loNew.AutoFilter.ShowAllData
    On Error GoTo 0
    If Not loNew.DataBodyRange Is Nothing Then loNew.DataBodyRange.FormatConditions.Delete

    ' Aplicar regla ShipTo borrando filas que no cumplen
    If Not loNew.DataBodyRange Is Nothing Then
        For r = loNew.DataBodyRange.Rows.count To 1 Step -1
            shipName = CStr(loNew.DataBodyRange.Cells(r, idxShip).Value)
            If ShipToRule(shipName, shipToMode, p1, p2) = False Then
                loNew.ListRows(r).Delete
            End If
        Next r
    End If

    ' Ordenar por color: filas rojas primero, luego azules (Family), luego verdes
    If InStr(1, newSheetName, "Family", vbTextCompare) > 0 Then
        SortTableByFontColor loNew, RGB(31, 78, 121)
        ApplyBlueOnlyFilter loNew, pctHeader
    Else
        SortTableByFontColor loNew, RGB(156, 0, 6)
        ApplyRedOnlyFilter loNew, pctHeader
    End If

    ' Re-forzar estilo light por si deletes lo afectaron
    On Error Resume Next
    loNew.TableStyle = LIGHT_TABLE_STYLE
    On Error GoTo 0
End Sub

Private Sub CreateViewByCopy( _
    ByVal wsSource As Worksheet, _
    ByVal newSheetName As String, _
    ByVal shipToMode As String, _
    ByVal p1 As String, _
    ByVal p2 As String _
)
    Dim wb As Workbook
    Dim wsNew As Worksheet
    Dim loNew As ListObject
    Dim idxShip As Long
    Dim r As Long
    Dim shipName As String

    Set wb = wsSource.Parent

    DeleteSheetIfExists wb, newSheetName

    wsSource.Copy After:=wb.Worksheets(wb.Worksheets.count)
    Set wsNew = wb.Worksheets(wb.Worksheets.count)
    wsNew.name = newSheetName

    If wsNew.ListObjects.count = 0 Then Exit Sub
    Set loNew = wsNew.ListObjects(1)

    'Forzar estilo light en la tabla copiada
    On Error Resume Next
    loNew.TableStyle = LIGHT_TABLE_STYLE
    On Error GoTo 0

    idxShip = GetColumnIndex(loNew, "Ship To Name")
    If idxShip = 0 Then Exit Sub

    On Error Resume Next
    If loNew.AutoFilter.FilterMode Then loNew.AutoFilter.ShowAllData
    On Error GoTo 0

    If Not loNew.DataBodyRange Is Nothing Then
        loNew.DataBodyRange.FormatConditions.Delete
    End If

    If Not loNew.DataBodyRange Is Nothing Then
        For r = loNew.DataBodyRange.Rows.count To 1 Step -1
            shipName = CStr(loNew.DataBodyRange.Cells(r, idxShip).Value)
            If ShipToRule(shipName, shipToMode, p1, p2) = False Then
                loNew.ListRows(r).Delete
            End If
        Next r
    End If

    SortTableByFontColor loNew, RGB(156, 0, 6)
    ApplyRedOnlyFilter loNew, "Percentage"

    'Re-forzar estilo light (por si deletes lo afectaron)
    On Error Resume Next
    loNew.TableStyle = LIGHT_TABLE_STYLE
    On Error GoTo 0
End Sub

Private Sub MarkMainFirst(ByVal loMain As ListObject)

    Dim ws As Worksheet
    Dim idxShip As Long, idxInit As Long, idxOrig As Long, idxPct As Long
    Dim pctColPos As Long
    Dim nRows As Long, r As Long
    Dim shipName As String
    Dim initQty As Double, origQty As Double, pct As Double

    Dim vShip As Variant, vInit As Variant, vOrig As Variant
    Dim rngRed As Range, rngRow As Range

    Dim calcMode As XlCalculation
    Dim scr As Boolean, ev As Boolean, da As Boolean

    If loMain Is Nothing Then Exit Sub
    If loMain.Range Is Nothing Then Exit Sub

    Set ws = loMain.Parent

    'Guardar estado y apagar UI
    scr = Application.ScreenUpdating
    ev = Application.EnableEvents
    da = Application.DisplayAlerts
    calcMode = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    '====================================================
    '1) LIMPIEZA TOTAL DE LA HOJA A "DEFAULT"
    '   - Fondo: sin relleno (transparente)
    '   - Fuente: color automático
    '   - Quitar formatos condicionales (hoja completa)
    '====================================================
    With ws.Cells
        .Interior.pattern = xlNone
        .Font.ColorIndex = xlAutomatic
    End With

    On Error Resume Next
    ws.Cells.FormatConditions.Delete
    On Error GoTo 0

    On Error Resume Next
    loMain.TableStyle = LIGHT_TABLE_STYLE
    On Error GoTo 0
    '====================================================
    '2) VALIDAR COLUMNAS + CREAR Percentage
    '====================================================
    If loMain.DataBodyRange Is Nothing Then GoTo CleanExit

    idxShip = GetColumnIndex(loMain, "Ship To Name")
    idxInit = GetColumnIndex(loMain, "Initial Order Qty (Before Rounding)")
    idxOrig = GetColumnIndex(loMain, "Original Order Qty (After Rounding)")

    If idxShip = 0 Or idxInit = 0 Or idxOrig = 0 Then GoTo CleanExit

    idxPct = GetColumnIndex(loMain, "Percentage")
    pctColPos = idxInit + 1

    If idxPct = 0 Then
        loMain.ListColumns.Add Position:=pctColPos
        loMain.HeaderRowRange.Cells(1, pctColPos).Value = "Percentage"
        idxPct = pctColPos
    End If

    'Fórmula + formato numérico
    loMain.ListColumns(idxPct).DataBodyRange.Formula = _
        "=[@[Original Order Qty (After Rounding)]]/[@[Initial Order Qty (Before Rounding)]]"
    loMain.ListColumns(idxPct).DataBodyRange.NumberFormat = "0.00"

    '====================================================
    '3) BORRAR FC SOLO EN LA TABLA (por si quedaba algo)
    '====================================================
    On Error Resume Next
    loMain.Range.FormatConditions.Delete
    On Error GoTo 0

    '====================================================
    '4) CARGAR ARRAYS Y DEFINIR QUÉ FILAS VAN ROJAS
    '   REGLAS:
    '   - ASM  ("as"):   pct > 1.0  → rojo   |  pct = 1.0 exacto → verde
    '   - HVDC (Family): pct > 2.0  → rojo   |  pct = 2.0 exacto → verde
    '   - Reg  (otros):  pct > 2.0  → rojo   |  pct = 2.0 exacto → verde
    '====================================================
    vShip = loMain.ListColumns(idxShip).DataBodyRange.Value
    vInit = loMain.ListColumns(idxInit).DataBodyRange.Value
    vOrig = loMain.ListColumns(idxOrig).DataBodyRange.Value

    nRows = UBound(vShip, 1)
    Set rngRed = Nothing

    For r = 1 To nRows

        shipName = CStr(vShip(r, 1))

        If IsNumeric(vInit(r, 1)) Then initQty = CDbl(vInit(r, 1)) Else initQty = 0#
        If IsNumeric(vOrig(r, 1)) Then origQty = CDbl(vOrig(r, 1)) Else origQty = 0#

        If initQty = 0 Then
            pct = 0#
        Else
            pct = origQty / initQty
        End If

        If InStr(1, shipName, "as", vbTextCompare) > 0 Then
            ' ASM: exactamente 1.0 = verde (no se marca), más de 1.0 = rojo
            If pct > 1# Then
                Set rngRow = loMain.DataBodyRange.Rows(r)
                If rngRed Is Nothing Then Set rngRed = rngRow Else Set rngRed = Union(rngRed, rngRow)
            End If

        ElseIf InStr(1, shipName, "HVDC", vbTextCompare) > 0 Then
            ' Family Care: exactamente 2.0 = verde (no se marca), más de 2.0 = rojo
            If pct > 2# Then
                Set rngRow = loMain.DataBodyRange.Rows(r)
                If rngRed Is Nothing Then Set rngRed = rngRow Else Set rngRed = Union(rngRed, rngRow)
            End If

        Else
            ' Reg: exactamente 2.0 = verde (no se marca), más de 2.0 = rojo
            If pct > 2# Then
                Set rngRow = loMain.DataBodyRange.Rows(r)
                If rngRed Is Nothing Then Set rngRed = rngRow Else Set rngRed = Union(rngRed, rngRow)
            End If
        End If

    Next r

    '====================================================
    '5) APLICAR FORMATO SOLO A ZVORDROUND (RANGO TABLA)
    '   - Primero verde a todo el DataBodyRange
    '   - Luego rojo solo a filas marcadas (pct estrictamente mayor al umbral)
    '====================================================
    With loMain.DataBodyRange
        .Interior.Color = RGB(198, 239, 206) 'verde claro
        .Font.Color = RGB(0, 97, 0)          'verde oscuro
    End With

    If Not rngRed Is Nothing Then
        With rngRed
            .Interior.Color = RGB(255, 199, 206) 'rosa
            .Font.Color = RGB(156, 0, 6)         'rojo oscuro
        End With
    End If

CleanExit:
    Application.Calculation = calcMode
    Application.DisplayAlerts = da
    Application.EnableEvents = ev
    Application.ScreenUpdating = scr

End Sub

Private Sub ApplyPinkRed(ByVal rngRow As Range)
    With rngRow
        .Interior.Color = RGB(255, 199, 206) 'rosa
        .Font.Color = RGB(156, 0, 6)         'rojo oscuro
    End With
End Sub

Private Sub AddPinkRule(ByVal loMain As ListObject, ByVal rngApplies As Range, ByVal formulaStr As String)
    Dim fc As FormatCondition
    Set fc = rngApplies.FormatConditions.Add(Type:=xlExpression, Formula1:=formulaStr)
    With fc
        .Interior.Color = RGB(255, 199, 206) 'rosa
        .Font.Color = RGB(156, 0, 6)         'rojo oscuro
        .StopIfTrue = False
    End With
End Sub


Private Function ShipToRule(ByVal shipName As String, ByVal mode As String, ByVal p1 As String, ByVal p2 As String) As Boolean
    Dim hasP1 As Boolean
    Dim hasP2 As Boolean

    hasP1 = (InStr(1, shipName, p1, vbTextCompare) > 0)
    hasP2 = (InStr(1, shipName, p2, vbTextCompare) > 0)

    Select Case UCase$(mode)
        Case "INCLUDE"
            ShipToRule = hasP1
        Case "EXCLUDE_2"
            ShipToRule = (Not hasP1) And (Not hasP2)
        Case Else
            ShipToRule = False
    End Select
End Function

Private Sub MarkRowsByPercentage(ByVal lo As ListObject, ByVal pctColIndex As Long, ByVal threshold As Double)

    Dim rng As Range
    Dim fc As FormatCondition
    Dim firstDataCell As Range
    Dim formulaStr As String

    If lo.DataBodyRange Is Nothing Then Exit Sub

    Set rng = lo.DataBodyRange
    Set firstDataCell = lo.ListColumns(pctColIndex).DataBodyRange.Cells(1, 1)

    'Eliminar reglas previas
    rng.FormatConditions.Delete

    'Usar referencia A1 (NO structured reference)
    formulaStr = "=" & firstDataCell.Address(False, False) & ">=" & threshold

    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:=formulaStr)

    With fc
        .Interior.Color = RGB(255, 199, 206) 'rosa
        .Font.Color = RGB(156, 0, 6)         'rojo oscuro
        .StopIfTrue = False
    End With

End Sub

Private Function ToDoubleSafe(ByVal v As Variant) As Double
    If IsNumeric(v) Then
        ToDoubleSafe = CDbl(v)
    Else
        ToDoubleSafe = 0#
    End If
End Function

'==================== PARTE 8 ====================
' Módulo: ZVORDROUND_Auto
' Propósito: Borrar hoja si existe
' Dependencias: Ninguna

Private Sub DeleteSheetIfExists(ByVal wb As Workbook, ByVal sheetName As String)
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.name, sheetName, vbTextCompare) = 0 Then
            ws.Delete
            Exit Sub
        End If
    Next ws
End Sub

'==================== PARTE 9 ====================
' Módulo: ZVORDROUND_Auto
' Propósito: Crear borrador de correo en Outlook con attachment (solo 1ra hoja del workbook principal)
' Dependencias: Outlook (late bound)

Private Sub CreateRoundingEmailDraft(ByVal wbMain As Workbook)

    Dim olApp As Object
    Dim draftsFolder As Object
    Dim olMail As Object

    Dim attachPath As String
    Dim htmlBody As String

    '1) Crear archivo adjunto: solo primera hoja del Excel principal
    attachPath = ExportFirstSheetAsXlsx(wbMain)

    '2) Armar HTML exacto
    htmlBody = ""
    htmlBody = htmlBody & "<p>Hello Team,</p>"
    htmlBody = htmlBody & "<p>We had Rounding today on <span style='color:red; font-weight:bold;'>Insert Categories</span>.</p>"
    htmlBody = htmlBody & "<ul>"
    htmlBody = htmlBody & "<li>Assembly items that round up and warehouse items that round above 2X the original amount were cut using R1.</li>"
    htmlBody = htmlBody & "<li>Assembly items that round down and warehouse items that round under 2X were released.</li>"
    htmlBody = htmlBody & "</ul>"
    htmlBody = htmlBody & "<p><b>Any Family Care rounding is cut with HA.</b></p>"
    htmlBody = htmlBody & "<p><b>BTDs were excluded while working rounding.</b></p>"
    htmlBody = htmlBody & "<p>This does not include any rounding that may occur for Fabric/Fem DS.</p>"

    '3) Obtener Outlook
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    If olApp Is Nothing Then Exit Sub

    '4) Drafts del mailbox pgwmordering (shared o account)
    Set draftsFolder = GetDraftsFolderForMailbox(olApp, DEFAULT_FROM_SMTP)
    If draftsFolder Is Nothing Then Exit Sub

    '5) Crear draft DIRECTO en Drafts de ese mailbox
    Set olMail = CreateDraftInFolder(draftsFolder)
    If olMail Is Nothing Then Exit Sub

    With olMail
        .To = "cnf-wgct_csoall@groups.pg.com"
        .cc = "parlow.or@pg.com; brown.mw@pg.com; butler.kl@pg.com"
        .Subject = "WM Stores Rounding " & Format(Date, "mm.dd.yyyy")
        .htmlBody = htmlBody

        If Len(Dir$(attachPath)) > 0 Then
            .Attachments.Add attachPath
        End If

        On Error Resume Next
        .SentOnBehalfOfName = DEFAULT_FROM_SMTP
        On Error GoTo 0

        .Save
        .Display
    End With

End Sub

'==================== PARTE 10 ====================
' Módulo: ZVORDROUND_Auto
' Propósito: Exportar solo la 1ra hoja del workbook principal a un .xlsx con nombre ZVROUND MM.DD.YYYY
' Dependencias: Ninguna

Private Function ExportFirstSheetAsXlsx(ByVal wbMain As Workbook) As String

    Dim wbTemp As Workbook
    Dim wsFirst As Worksheet
    Dim fName As String
    Dim fPath As String

    Set wsFirst = wbMain.Worksheets(1)

    fName = "ZVROUND " & Format(Date, "mm.dd.yyyy") & ".xlsx"
    fPath = Environ$("TEMP") & "\" & fName

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    'Crear workbook temporal y copiar solo la primera hoja
    Set wbTemp = Workbooks.Add(xlWBATWorksheet)
    wsFirst.Copy Before:=wbTemp.Worksheets(1)

    'Eliminar la hoja vacía que creó Workbooks.Add
    On Error Resume Next
    wbTemp.Worksheets(wbTemp.Worksheets.count).Delete
    On Error GoTo 0

    'Guardar como xlsx (sin macros)
    wbTemp.SaveAs fileName:=fPath, FileFormat:=51 'xlOpenXMLWorkbook
    wbTemp.Close SaveChanges:=False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    ExportFirstSheetAsXlsx = fPath

End Function

'==================== PARTE X ====================
' Módulo: ZVORDROUND_Auto
' Propósito: Decidir y ordenar antes del split (por Material o por Order Number)
' Dependencias: Ninguna

Private Sub DecideAndSortMain(ByVal lo As ListObject)

    Dim nOrders As Long
    Dim nMaterials As Long

    gCutDecision = "" 'reset

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    nOrders = CountDistinctInColumn(lo, "Order Number")
    nMaterials = CountDistinctInColumn(lo, "Material")

    If nOrders > nMaterials Then
        SortListObjectByColumn lo, "Material"
        gCutDecision = "MATERIAL"
    Else
        SortListObjectByColumn lo, "Order Number"
        gCutDecision = "SO"
    End If

End Sub

Private Function CountDistinctInColumn(ByVal lo As ListObject, ByVal headerName As String) As Long

    Dim dict As Object
    Dim idx As Long
    Dim v As Variant
    Dim i As Long
    Dim key As String

    If lo.DataBodyRange Is Nothing Then
        CountDistinctInColumn = 0
        Exit Function
    End If

    idx = GetColumnIndex(lo, headerName)
    If idx = 0 Then
        CountDistinctInColumn = 0
        Exit Function
    End If

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    v = lo.ListColumns(idx).DataBodyRange.Value

    For i = 1 To UBound(v, 1)
        key = Trim$(CStr(v(i, 1)))
        If Len(key) > 0 Then
            If Not dict.Exists(key) Then dict.Add key, 1
        End If
    Next i

    CountDistinctInColumn = dict.count

End Function

Private Sub SortListObjectByColumn(ByVal lo As ListObject, ByVal headerName As String)

    Dim idx As Long

    idx = GetColumnIndex(lo, headerName)
    If idx = 0 Then Exit Sub

    With lo.Sort
        .SortFields.Clear
        .SortFields.Add key:=lo.ListColumns(idx).Range, _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

End Sub

'==================== PARTE Y ====================
' Módulo: ZVORDROUND_Auto
' Propósito: Dejar la tabla filtrada mostrando solo valores en rojo (por color de fuente)
' Dependencias: Ninguna

Private Sub ApplyRedOnlyFilter(ByVal lo As ListObject, ByVal pctHeader As String)

    Dim idx As Long

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    idx = GetColumnIndex(lo, pctHeader)
    If idx = 0 Then Exit Sub

    'Quitar cualquier filtro previo
    On Error Resume Next
    If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    On Error GoTo 0

    'Filtrar por color de fuente rojo (el que usa MarkMainFirst)
    lo.Range.AutoFilter Field:=idx, Criteria1:=RGB(156, 0, 6), Operator:=xlFilterFontColor

End Sub

Private Sub ShowCutDecisionMessage()

    Dim msg As String

    If UCase$(gCutDecision) = "MATERIAL" Then
        msg = "Today you will have to cut by Material."
    ElseIf UCase$(gCutDecision) = "SO" Then
        msg = "Today you will have to cut by SO."
    Else
        msg = "Cut decision could not be determined today."
    End If

    MsgBox msg, vbInformation, "Rounding Cut Decision"

End Sub

Private Sub ApplyLightStyleToAllTables(ByVal wb As Workbook)

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim s As String

    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects

            '1) Intentar el estilo configurado
            If TrySetTableStyle(lo, LIGHT_TABLE_STYLE) Then GoTo NextTable

            '2) Fallbacks "Light" comunes
            If TrySetTableStyle(lo, "TableStyleLight1") Then GoTo NextTable
            If TrySetTableStyle(lo, "TableStyleLight9") Then GoTo NextTable
            If TrySetTableStyle(lo, "TableStyleLight11") Then GoTo NextTable
            If TrySetTableStyle(lo, "TableStyleLight15") Then GoTo NextTable

            '3) Último fallback (si por política no están los Light)
            Call TrySetTableStyle(lo, "TableStyleMedium2")

NextTable:
        Next lo
    Next ws

End Sub

Private Function CreateOrReplaceZVORDROUNDTableInWorkbook(ByVal wb As Workbook, ByRef wsOut As Worksheet) As ListObject

    Dim ws As Worksheet
    Dim rng As Range
    Dim lo As ListObject
    Dim i As Long
    Dim missing As String
    Dim baseName As String

    Set ws = wb.Worksheets(1)
    Set wsOut = ws

    '====================================================
    '1) Validar nombre del archivo = "EXPORT"
    '====================================================
    baseName = wb.name
    If InStrRev(baseName, ".") > 0 Then
        baseName = Left$(baseName, InStrRev(baseName, ".") - 1)
    End If

    If StrComp(baseName, "EXPORT", vbTextCompare) <> 0 Then
        MsgBox "Este documento no es el EXPORT que necesito." & vbCrLf & _
               "Nombre detectado: " & wb.name, vbExclamation, "ZVORDROUND"
        Set CreateOrReplaceZVORDROUNDTableInWorkbook = Nothing
        Exit Function
    End If

    '====================================================
    '2) Validar rango usable
    '====================================================
    If ws.UsedRange Is Nothing Then GoTo NotReport
    Set rng = ws.UsedRange
    If rng.Rows.count < 2 Or rng.Columns.count < 1 Then GoTo NotReport

    '====================================================
    '3) Validar columnas obligatorias
    '====================================================
    missing = MissingHeadersInRow(rng.Rows(1), Array( _
        "Ship To Name", _
        "Initial Order Qty (Before Rounding)", _
        "Original Order Qty (After Rounding)", _
        "Order Number", _
        "Material" _
    ))

    If Len(missing) > 0 Then
        MsgBox "Este documento no es el reporte que necesito." & vbCrLf & _
               "Faltan columnas: " & missing, vbExclamation, "ZVORDROUND"
        Set CreateOrReplaceZVORDROUNDTableInWorkbook = Nothing
        Exit Function
    End If

    '====================================================
    '4) Quitar tablas existentes
    '====================================================
    On Error Resume Next
    For i = ws.ListObjects.count To 1 Step -1
        ws.ListObjects(i).Unlist
    Next i
    On Error GoTo 0

    '====================================================
    '5) Crear tabla ZVORDROUND
    '====================================================
    Set lo = ws.ListObjects.Add(SourceType:=xlSrcRange, _
                                Source:=rng, _
                                XlListObjectHasHeaders:=xlYes)

    On Error Resume Next
    lo.name = "ZVORDROUND"
    lo.TableStyle = LIGHT_TABLE_STYLE
    On Error GoTo 0

    Set CreateOrReplaceZVORDROUNDTableInWorkbook = lo
    Exit Function

NotReport:
    MsgBox "Este documento no es el reporte que necesito.", vbExclamation, "ZVORDROUND"
    Set CreateOrReplaceZVORDROUNDTableInWorkbook = Nothing

End Function

Private Function TrySetTableStyle(ByVal lo As ListObject, ByVal styleName As String) As Boolean
    On Error GoTo EH
    lo.TableStyle = styleName
    TrySetTableStyle = True
    Exit Function
EH:
    TrySetTableStyle = False
End Function


Private Function MissingHeadersInRow(ByVal headerRow As Range, ByVal requiredHeaders As Variant) As String
    Dim dict As Object, c As Range
    Dim i As Long, h As String
    Dim miss As String

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    For Each c In headerRow.Cells
        h = Trim$(CStr(c.Value))
        If h <> vbNullString Then
            If Not dict.Exists(h) Then dict.Add h, True
        End If
    Next c

    For i = LBound(requiredHeaders) To UBound(requiredHeaders)
        h = Trim$(CStr(requiredHeaders(i)))
        If h <> vbNullString Then
            If Not dict.Exists(h) Then
                If miss = vbNullString Then miss = h Else miss = miss & ", " & h
            End If
        End If
    Next i

    MissingHeadersInRow = miss
End Function

Private Sub ApplyFamilyStyle(ByVal loMain As ListObject)

    Dim idxShip As Long
    Dim r As Long
    Dim shipName As String
    Dim rngFamily As Range
    Dim rngRow As Range

    If loMain Is Nothing Then Exit Sub
    If loMain.DataBodyRange Is Nothing Then Exit Sub

    idxShip = GetColumnIndex(loMain, "Ship To Name")
    If idxShip = 0 Then Exit Sub

    Set rngFamily = Nothing

    For r = 1 To loMain.DataBodyRange.Rows.count
        shipName = CStr(loMain.DataBodyRange.Cells(r, idxShip).Value)

        If InStr(1, shipName, "HVDC", vbTextCompare) > 0 Then
            Set rngRow = loMain.DataBodyRange.Rows(r)
            If rngFamily Is Nothing Then
                Set rngFamily = rngRow
            Else
                Set rngFamily = Union(rngFamily, rngRow)
            End If
        End If
    Next r

    If Not rngFamily Is Nothing Then
        With rngFamily
            .Interior.Color = RGB(221, 235, 247) ' celeste profesional
            .Font.Color = RGB(31, 78, 121)       ' azul profesional
        End With
    End If

End Sub

Private Sub ApplyBlueOnlyFilter(ByVal lo As ListObject, ByVal pctHeader As String)

    Dim idx As Long

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    idx = GetColumnIndex(lo, pctHeader)
    If idx = 0 Then Exit Sub

    On Error Resume Next
    If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    On Error GoTo 0

    ' Filtrar por color de fuente azul de Family
    lo.Range.AutoFilter Field:=idx, Criteria1:=RGB(31, 78, 121), Operator:=xlFilterFontColor

End Sub

'==================== PARTE Z ====================
' Módulo: ZVORDROUND_Auto
' Propósito: Ordenar tabla por color de fuente — el color prioritario queda arriba
' Dependencias: Ninguna

Private Sub SortTableByFontColor(ByVal lo As ListObject, ByVal priorityColor As Long)
    ' Compatibilidad con todas las versiones: columna auxiliar 0/1, sort, borrar columna.
    Dim ws As Worksheet
    Dim helperCol As ListColumn
    Dim helperRng As Range
    Dim sortRng As Range
    Dim r As Long
    Dim nRows As Long

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Set ws = lo.Parent
    nRows = lo.DataBodyRange.Rows.count

    ' 1) Agregar columna auxiliar al final de la tabla
    Set helperCol = lo.ListColumns.Add
    helperCol.name = "_SortHelper_"
    Set helperRng = helperCol.DataBodyRange

    ' 2) Marcar 0 = color prioritario (va arriba), 1 = resto
    Dim cellColor As Long
    For r = 1 To nRows
        cellColor = lo.DataBodyRange.Cells(r, 1).Font.Color
        helperRng.Cells(r, 1).Value = IIf(cellColor = priorityColor, 0, 1)
    Next r

    ' 3) Ordenar por la columna auxiliar
    Set sortRng = helperRng
    With lo.Sort
        .SortFields.Clear
        .SortFields.Add key:=sortRng, SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    ' 4) Borrar columna auxiliar
    helperCol.Delete

End Sub
