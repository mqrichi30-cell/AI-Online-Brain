Attribute VB_Name = "ApptsWakefern"
'==================== MACRO WAKEFERN APPT SCHED ====================
' Comentarios en español. Todos los mensajes al usuario en inglés.

' ==================== Destinatarios ====================
Private Const GROCERY_TO As String = _
    "Steve.Salotti@wakefern.com; Grocery_Special_PO_Group@wakefern.com; Mark.Kielczynski@wakefern.com; anne.mucchiello@wakefern.com; Al.DAgostino@wakefern.com"

Private Const HABA_TO As String = _
    "GMScheduling@wakefern.com; sharper.d@pg.com; Irma.Lang@wakefern.com;"

Private Const CC_DEFAULT As String = "jackson.vs@pg.com; rocas.cr@pg.com"

' ==================== Firma y template ====================
Private Const FIRMA_W2 As String = _
    "C:\Users\marin.c\AppData\Roaming\Microsoft\Signatures\AA - Cristhofer (NO BORRAR) (pgcustservw2.im@pg.com).htm"

Private Const TEMPLATE_CHANGES_DT As String = _
    "C:\Users\marin.c\OneDrive - Procter and Gamble\Desktop\Templates\Changes DT.xlsx"

Private Const MARK_SENDER As String = "mark.kielczynski@wakefern.com"

' ==================== ENTRADA PRINCIPAL ====================
' Com: Esta Sub DEBE ser llamada por la regla Application_NewMailEx/ItemAdd pasando el MailItem que disparó el evento.
Public Sub DetectarTablasYColumnas(ByVal mail As Outlook.MailItem)
    On Error GoTo ErrHandler

    If mail Is Nothing Then
        MsgBox "No source email was provided to the macro.", vbExclamation
        Exit Sub
    End If

    ' ===== GUARDIA RE:/FW: =====
    ' Com: Si el asunto empieza con RE: o FW: verificamos el remitente.
    '      - Si viene de Mark Kielczynski → flujo especial MOC desde texto plano.
    '      - Si viene de cualquier otro → ignorar (es respuesta al hilo principal).
    Dim subjUpper As String
    subjUpper = UCase$(Trim$(mail.Subject))
    If Left$(subjUpper, 3) = "RE:" Or Left$(subjUpper, 3) = "FW:" Then
        Dim senderEmail As String
        senderEmail = ""
        On Error Resume Next
        senderEmail = LCase$(mail.SenderEmailAddress)
        ' Com: En Exchange el SenderEmailAddress puede ser la dirección SMTP o un EX DN;
        '      intentamos también con el campo Reply-To via PropertyAccessor.
        If InStr(1, senderEmail, "@") = 0 Then
            senderEmail = LCase$(mail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x5D01001E"))
        End If
        On Error GoTo ErrHandler

        If InStr(1, senderEmail, MARK_SENDER, vbTextCompare) > 0 Then
            ' Com: RE/FW de Mark con datos de PO → construir MOC Request
            ProcesarCorreoMark mail
        End If
        ' Com: RE/FW de cualquier otro → no hacer nada
        Exit Sub
    End If

    ' --- HTML / tablas ---
    Dim htmlDoc As Object, tables As Object, table As Object
    Dim row As Object, cell As Object
    Dim htmlBody As String
    Dim seDetectoAppt As Boolean
    Dim i As Long

    htmlBody = mail.htmlBody
    Set htmlDoc = CreateObject("htmlfile")
    htmlDoc.Open
    htmlDoc.Write htmlBody
    htmlDoc.Close
    Set tables = htmlDoc.getElementsByTagName("table")

    If (tables Is Nothing) Or tables.Length = 0 Then
        MsgBox "No tables were found in this email.", vbInformation
        Exit Sub
    End If

    seDetectoAppt = False
    For Each table In tables
        If table.Rows.Length > 0 Then
            For Each cell In table.Rows(0).cells
                If Trim$(cell.innerText) = "Delivery  Time" Then
                    seDetectoAppt = True
                    Exit For
                End If
            Next cell
        End If
        If seDetectoAppt Then Exit For
    Next table

    ' ===== SIN APPT: autodetectar Grocery/HABA por Subject =====
    If Not seDetectoAppt Then
        Dim grupo As String, isGrocery As Boolean, resp As VbMsgBoxResult
        grupo = DetectarGrupoPorSubject(mail.Subject)

        If grupo = "GROCERY" Then
            isGrocery = True
        ElseIf grupo = "HABA" Then
            isGrocery = False
        Else
            resp = MsgBox("No 'Delivery  Time' column was found and the subject didn't indicate GROCERY or HABA." & vbCrLf & _
                          "Is this GROCERY? (Yes = Grocery, No = HABA)", vbQuestion + vbYesNo, "Wakefern Routing")
            isGrocery = (resp = vbYes)
        End If

        CrearCorreoWakefernDesdeHTML mail, isGrocery, False
        Exit Sub
    End If

    ' ===== CON APPT: generar Excel + correo Wakefern (Grocery) + draft OSSGenAI =====
    Dim excelApp As Object, wbNew As Object, wsNew As Object
    Dim wbTemplate As Object, wsTemplate As Object
    Dim r As Long, c As Long

    Set excelApp = CreateObject("Excel.Application")
    With excelApp
        .Visible = False: .ScreenUpdating = False: .DisplayAlerts = False: .EnableEvents = False
    End With

    Set wbNew = excelApp.Workbooks.Add
    Set wsNew = wbNew.Sheets(1)

    r = 1
    For Each table In tables
        c = 1
        For Each row In table.Rows
            For Each cell In row.cells
                wsNew.cells(r, c).Value = Trim$(cell.innerText)
                c = c + 1
            Next cell
            r = r + 1: c = 1
        Next row
        r = r + 1
    Next table

    Dim colPO As Long, colShipPOs As Long, colDelivery As Long, colDeliveryTime As Long
    colPO = 0: colShipPOs = 0: colDelivery = 0: colDeliveryTime = 0

    For i = 1 To wsNew.UsedRange.Columns.Count
        Select Case Trim$(wsNew.cells(1, i).Value)
            Case "PO":               colPO = i
            Case "Ship with POs":    colShipPOs = i
            Case "Delivery":         colDelivery = i
            Case "Delivery  Time":   colDeliveryTime = i
        End Select
    Next i

    If colPO = 0 Or colShipPOs = 0 Or colDelivery = 0 Or colDeliveryTime = 0 Then
        MsgBox "Required columns missing: PO, Ship with POs, Delivery, Delivery  Time.", vbCritical
        GoTo Limpieza
    End If

    If Dir$(TEMPLATE_CHANGES_DT, vbNormal) = "" Then
        MsgBox "Template not found: " & TEMPLATE_CHANGES_DT, vbCritical
        GoTo Limpieza
    End If

    Set wbTemplate = excelApp.Workbooks.Open(TEMPLATE_CHANGES_DT, ReadOnly:=False)
    Set wsTemplate = wbTemplate.Sheets(1)

    Dim colSchedDate As Long, colSchedTime As Long, colPONumber As Long
    colSchedDate = 0: colSchedTime = 0: colPONumber = 0

    For i = 1 To wsTemplate.UsedRange.Columns.Count
        Select Case Trim$(wsTemplate.cells(1, i).Value)
            Case "SCHED_DATE": colSchedDate = i
            Case "SCHED_TIME": colSchedTime = i
            Case "PO_NUMBER":  colPONumber = i
        End Select
    Next i

    If colSchedDate = 0 Or colSchedTime = 0 Or colPONumber = 0 Then
        MsgBox "Template columns not found: SCHED_DATE, SCHED_TIME, PO_NUMBER.", vbCritical
        GoTo LimpiezaWB
    End If

    Dim lastDataRow As Long
    lastDataRow = wsTemplate.cells(wsTemplate.Rows.Count, colPONumber).End(-4162).row
    If lastDataRow >= 2 Then
        wsTemplate.Range(wsTemplate.Rows(2), wsTemplate.Rows(lastDataRow)).ClearContents
    End If

    Dim fila As Long, otrosPOs As String, poList As Variant, po As Variant
    Dim fecha As String, horaRaw As String, horaFinal As Variant, poPrincipal As String
    fila = 2

    For i = 2 To wsNew.cells(wsNew.Rows.Count, colPO).End(-4162).row
        fecha = CStr(wsNew.cells(i, colDelivery).Value)
        horaRaw = CStr(wsNew.cells(i, colDeliveryTime).Value)
        horaFinal = FormatoHora(horaRaw)
        If IsEmpty(horaFinal) Or horaFinal = "" Then GoTo Siguiente

        poPrincipal = CStr(wsNew.cells(i, colPO).Value)
        otrosPOs = CStr(wsNew.cells(i, colShipPOs).Value)

        If Len(poPrincipal) > 0 Then
            wsTemplate.cells(fila, colSchedDate).Value = fecha
            wsTemplate.cells(fila, colSchedTime).Value = horaFinal
            wsTemplate.cells(fila, colPONumber).Value = poPrincipal
            fila = fila + 1
        End If

        If Len(otrosPOs) > 0 Then
            poList = Split(otrosPOs, "/")
            For Each po In poList
                po = Trim$(po)
                If Len(po) > 0 Then
                    wsTemplate.cells(fila, colSchedDate).Value = fecha
                    wsTemplate.cells(fila, colSchedTime).Value = horaFinal
                    wsTemplate.cells(fila, colPONumber).Value = po
                    fila = fila + 1
                End If
            Next po
        End If
Siguiente:
    Next i

    If fila > 2 Then
        wsTemplate.Range(wsTemplate.cells(2, colSchedTime), wsTemplate.cells(fila - 1, colSchedTime)).NumberFormat = "hh:mm:ss"
    End If

    Dim ultimaFilaPO As Long, poVal As String, j As Long
    ultimaFilaPO = wsTemplate.cells(wsTemplate.Rows.Count, colPONumber).End(-4162).row
    For j = 2 To ultimaFilaPO
        poVal = CStr(wsTemplate.cells(j, colPONumber).Value)
        If IsNumeric(poVal) And Len(poVal) > 1 Then
            wsTemplate.cells(j, colPONumber).Value = Mid$(poVal, 2)
        End If
    Next j

    wbTemplate.Save
    wbTemplate.Close SaveChanges:=False

    CrearCorreoWakefernDesdeHTML mail, True, True
    CreateDraftFromPG2WithSignatureAndAttachment TEMPLATE_CHANGES_DT

Limpieza:
    On Error Resume Next
    If Not wbNew Is Nothing Then wbNew.Close False
    If Not excelApp Is Nothing Then
        excelApp.DisplayAlerts = True
        excelApp.ScreenUpdating = True
        excelApp.EnableEvents = True
        excelApp.Quit
    End If
    On Error GoTo 0
    Exit Sub

LimpiezaWB:
    On Error Resume Next
    If Not wbTemplate Is Nothing Then wbTemplate.Close False
    GoTo Limpieza

ErrHandler:
    MsgBox "An unexpected error occurred: " & Err.Description, vbCritical
End Sub

' ==================== Procesador de correo de Mark Kielczynski (RE:/FW:) ====================
' Com: Parsea las líneas de texto plano que Mark envía con datos de PO.
' Formato de cada línea significativa:
'   736911 CC PROC 006 PL 1432 PROCTER & GAMBLE 06 09 26 06 09 26 04 00 P
'   tok(0)=PO   tok(1..8)=ignorar (CC PROC ### PL #### PROCTER & GAMBLE)
'   tok(9..11)=fecha desconocida mm dd yy  (ignorar)
'   tok(12..14)=RDD mm dd yy  (SCHED_DATE)
'   tok(15)=hora  tok(16)=minutos  tok(17)=A o P (AM/PM)
Private Sub ProcesarCorreoMark(ByVal mail As Outlook.MailItem)
    On Error GoTo ErrHandler

    If Dir$(TEMPLATE_CHANGES_DT, vbNormal) = "" Then
        MsgBox "Template not found: " & TEMPLATE_CHANGES_DT, vbCritical
        Exit Sub
    End If

    ' Com: Extraer texto plano del cuerpo del correo
    Dim bodyText As String
    bodyText = mail.Body
    Dim lines() As String
    lines = Split(bodyText, vbCrLf)
    ' Com: Algunos clientes usan solo vbLf
    If UBound(lines) < 1 Then lines = Split(bodyText, vbLf)

    ' Com: Abrir Excel y template
    Dim excelApp As Object, wbTemplate As Object, wsTemplate As Object
    Set excelApp = CreateObject("Excel.Application")
    With excelApp
        .Visible = False: .ScreenUpdating = False: .DisplayAlerts = False: .EnableEvents = False
    End With

    Set wbTemplate = excelApp.Workbooks.Open(TEMPLATE_CHANGES_DT, ReadOnly:=False)
    Set wsTemplate = wbTemplate.Sheets(1)

    ' Com: Localizar columnas del template
    Dim i As Long
    Dim colSchedDate As Long, colSchedTime As Long, colPONumber As Long
    colSchedDate = 0: colSchedTime = 0: colPONumber = 0

    For i = 1 To wsTemplate.UsedRange.Columns.Count
        Select Case Trim$(wsTemplate.cells(1, i).Value)
            Case "SCHED_DATE": colSchedDate = i
            Case "SCHED_TIME": colSchedTime = i
            Case "PO_NUMBER":  colPONumber = i
        End Select
    Next i

    If colSchedDate = 0 Or colSchedTime = 0 Or colPONumber = 0 Then
        MsgBox "Template columns not found: SCHED_DATE, SCHED_TIME, PO_NUMBER.", vbCritical
        GoTo LimpiezaMark
    End If

    ' Com: Limpiar datos previos del template (conservar encabezados)
    Dim lastDataRow As Long
    lastDataRow = wsTemplate.cells(wsTemplate.Rows.Count, colPONumber).End(-4162).row
    If lastDataRow >= 2 Then
        wsTemplate.Range(wsTemplate.Rows(2), wsTemplate.Rows(lastDataRow)).ClearContents
    End If

    ' Com: Parsear cada línea
    Dim fila As Long
    fila = 2
    Dim lin As Variant
    For Each lin In lines
        Dim lineStr As String
        lineStr = Trim$(CStr(lin))
        If Len(lineStr) = 0 Then GoTo SiguienteLinea

        ' Com: Dividir por espacios y filtrar tokens vacíos
        Dim rawTokens() As String
        rawTokens = Split(lineStr, " ")
        Dim tokens() As String
        Dim tCount As Long
        tCount = 0
        Dim t As Variant
        For Each t In rawTokens
            If Len(Trim$(CStr(t))) > 0 Then
                ReDim Preserve tokens(tCount)
                tokens(tCount) = Trim$(CStr(t))
                tCount = tCount + 1
            End If
        Next t

        ' Com: Necesitamos exactamente 18 tokens para el formato de Mark
        '      tok(0)=PO, tok(1)=CC, tok(2)=PROC, tok(3)=###, tok(4)=PL,
        '      tok(5)=####, tok(6)=PROCTER, tok(7)=&, tok(8)=GAMBLE,
        '      tok(9..11)=fecha1 mm dd yy, tok(12..14)=RDD mm dd yy,
        '      tok(15)=hh, tok(16)=mm, tok(17)=A/P
        If tCount < 18 Then GoTo SiguienteLinea
        If Not IsNumeric(tokens(0)) Then GoTo SiguienteLinea

        Dim poNum As String
        poNum = tokens(0)

        ' Com: Construir fecha RDD desde tok(12) tok(13) tok(14) → MM/DD/YY
        Dim rddMM As String, rddDD As String, rddYY As String
        rddMM = tokens(12)
        rddDD = tokens(13)
        rddYY = tokens(14)
        ' Com: Validar que los tres sean numéricos
        If Not (IsNumeric(rddMM) And IsNumeric(rddDD) And IsNumeric(rddYY)) Then GoTo SiguienteLinea

        Dim rddDate As Date
        On Error Resume Next
        rddDate = DateSerial(2000 + CInt(rddYY), CInt(rddMM), CInt(rddDD))
        On Error GoTo ErrHandler
        If rddDate = 0 Then GoTo SiguienteLinea

        ' Com: Construir hora de appt desde tok(15) (hh), tok(16) (mm), tok(17) (A/P)
        Dim apptHH As Integer, apptMM As Integer
        Dim ampm As String
        If Not IsNumeric(tokens(15)) Then GoTo SiguienteLinea
        If Not IsNumeric(tokens(16)) Then GoTo SiguienteLinea
        apptHH = CInt(tokens(15))
        apptMM = CInt(tokens(16))
        ampm = UCase$(Left$(Trim$(tokens(17)), 1))

        ' Com: Convertir a 24 horas
        If ampm = "P" Then
            If apptHH <> 12 Then apptHH = apptHH + 12
        ElseIf ampm = "A" Then
            If apptHH = 12 Then apptHH = 0
        End If

        If apptHH > 23 Or apptMM > 59 Then GoTo SiguienteLinea

        Dim apptTime As Date
        apptTime = TimeSerial(apptHH, apptMM, 0)

        ' Com: Ajuste de PO: remover primer dígito si aplica (igual que flujo Grocery)
        Dim poFinal As String
        poFinal = poNum
        If IsNumeric(poFinal) And Len(poFinal) > 1 Then
            poFinal = Mid$(poFinal, 2)
        End If

        ' Com: Escribir en template
        wsTemplate.cells(fila, colSchedDate).Value = rddDate
        wsTemplate.cells(fila, colSchedTime).Value = apptTime
        wsTemplate.cells(fila, colPONumber).Value = poFinal
        fila = fila + 1

SiguienteLinea:
    Next lin

    If fila = 2 Then
        MsgBox "No valid PO lines were found in Mark's email.", vbExclamation
        GoTo LimpiezaMark
    End If

    ' Com: Aplicar formato de hora a la columna SCHED_TIME
    wsTemplate.Range(wsTemplate.cells(2, colSchedTime), wsTemplate.cells(fila - 1, colSchedTime)).NumberFormat = "hh:mm:ss"

    ' Com: Aplicar formato de fecha a la columna SCHED_DATE
    wsTemplate.Range(wsTemplate.cells(2, colSchedDate), wsTemplate.cells(fila - 1, colSchedDate)).NumberFormat = "mm/dd/yyyy"

    wbTemplate.Save
    wbTemplate.Close SaveChanges:=False
    Set wsTemplate = Nothing
    Set wbTemplate = Nothing

    excelApp.Quit
    Set excelApp = Nothing

    ' Com: Crear draft RPA EMAIL REQUEST TO CHANGES con el template adjunto
    CreateDraftFromPG2WithSignatureAndAttachment TEMPLATE_CHANGES_DT

    MsgBox "Mark's PO data processed: " & (fila - 2) & " row(s) written. Draft RPA email created.", vbInformation
    Exit Sub

LimpiezaMark:
    On Error Resume Next
    If Not wbTemplate Is Nothing Then wbTemplate.Close False
    If Not excelApp Is Nothing Then
        excelApp.DisplayAlerts = True
        excelApp.ScreenUpdating = True
        excelApp.EnableEvents = True
        excelApp.Quit
    End If
    On Error GoTo 0
    Exit Sub

ErrHandler:
    MsgBox "Error processing Mark's email: " & Err.Description, vbCritical
End Sub

' ==================== Detección de grupo por Subject ====================
Private Function DetectarGrupoPorSubject(ByVal subj As String) As String
    Dim u As String
    u = UCase$(Trim$(subj))
    If InStr(1, u, "GROCERY", vbTextCompare) > 0 Then
        DetectarGrupoPorSubject = "GROCERY"
    ElseIf InStr(1, u, "HABA", vbTextCompare) > 0 Then
        DetectarGrupoPorSubject = "HABA"
    Else
        DetectarGrupoPorSubject = ""
    End If
End Function

' ==================== Construcción del correo Wakefern ====================
Private Sub CrearCorreoWakefernDesdeHTML(ByVal mail As Outlook.MailItem, ByVal isGrocery As Boolean, ByVal conAppt As Boolean)
    On Error GoTo ErrHandler

    Dim olApp As Object, olMail As Object, olAccount As Object
    Dim regex As Object, matches As Object
    Dim htmlOriginal As String, subjectOriginal As String
    Dim tabla1 As String, tabla2 As String
    Dim firmaHTML As String

    subjectOriginal = mail.Subject
    If Len(Trim$(subjectOriginal)) = 0 Then subjectOriginal = "No subject"
    htmlOriginal = mail.htmlBody

    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Pattern = "<table[\s\S]*?</table>"
        .Global = True
        .IgnoreCase = True
    End With

    If Not regex.Test(htmlOriginal) Then
        MsgBox "No tables were found in this email.", vbExclamation
        Exit Sub
    End If

    Set matches = regex.Execute(htmlOriginal)
    tabla1 = matches(0).Value
    Dim reLink As Object
    Set reLink = CreateObject("VBScript.RegExp")
    With reLink
        .Pattern = "please\s+link"
        .Global = True
        .IgnoreCase = True
    End With
    tabla1 = reLink.Replace(tabla1, "Linked")

    If matches.Count >= 2 Then
        tabla2 = matches(1).Value
    Else
        tabla2 = ""
    End If

    firmaHTML = LeerFirmaHTML(FIRMA_W2)

    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)

    With olMail
        If isGrocery Then
            .To = GROCERY_TO
        Else
            .To = HABA_TO
        End If
        .CC = CC_DEFAULT
        .Subject = "RE: " & subjectOriginal

        If conAppt Then
            If Len(tabla2) > 0 Then
                .htmlBody = _
                    "<p>Hello team,</p>" & _
                    "<p>Please set delivery appt time for each PO on each truck:</p>" & _
                    tabla1 & _
                    "<br><br><p>Please have Wakefern GROCERY schedule:</p>" & _
                    tabla2 & _
                    "<br><br>" & firmaHTML
            Else
                .htmlBody = _
                    "<p>Hello team,</p>" & _
                    "<p>Please set delivery appt time for the PO(s) below:</p>" & _
                    tabla1 & _
                    "<br><br>" & firmaHTML
            End If

        ElseIf isGrocery Then
            If Len(tabla2) > 0 Then
                .htmlBody = _
                    "<p>Hello team,</p>" & _
                    "<p>Please process the PO(s) below accordingly:</p>" & _
                    tabla1 & _
                    "<br><br><p>Additionally, please handle the following:</p>" & _
                    tabla2 & _
                    "<br><br>" & firmaHTML
            Else
                .htmlBody = _
                    "<p>Hello team,</p>" & _
                    "<p>Please process the following PO(s):</p>" & _
                    tabla1 & _
                    "<br><br>" & firmaHTML
            End If

        Else
            If Len(tabla2) > 0 Then
                .htmlBody = _
                    "<p>Hello team,</p>" & _
                    "<p>Please schedule these trucks:</p>" & _
                    tabla1 & _
                    "<br><br><p>Please schedule:</p>" & _
                    tabla2 & _
                    "<br><br>" & firmaHTML
            Else
                .htmlBody = _
                    "<p>Hello team,</p>" & _
                    "<p>Please schedule these trucks:</p>" & _
                    tabla1 & _
                    "<br><br>" & firmaHTML
            End If
        End If

        For Each olAccount In olApp.Session.Accounts
            If InStr(1, olAccount.smtpAddress, "pgcustservw2.im@pg.com", vbTextCompare) > 0 Then
                Set .SendUsingAccount = olAccount
                Exit For
            End If
        Next olAccount

        .Display
    End With

    MsgBox "Draft created with " & matches.Count & " table(s).", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error creating the draft: " & Err.Description, vbCritical
End Sub

' ==================== Utilidades ====================
Private Function LeerFirmaHTML(ByVal ruta As String) As String
    On Error GoTo ErrHandler
    Dim fso As Object, ts As Object
    LeerFirmaHTML = ""
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(ruta) Then
        Set ts = fso.OpenTextFile(ruta, 1)
        LeerFirmaHTML = ts.ReadAll
        ts.Close
    End If
    Exit Function
ErrHandler:
    LeerFirmaHTML = ""
End Function

Private Function FormatoHora(ByVal strHora As String) As Variant
    strHora = Trim$(strHora)
    If Len(strHora) = 5 Then strHora = "0" & strHora
    If Len(strHora) <> 6 Or Not IsNumeric(strHora) Then
        FormatoHora = ""
        Exit Function
    End If

    Dim h As Integer, m As Integer, s As Integer
    h = CInt(Left$(strHora, 2))
    m = CInt(Mid$(strHora, 3, 2))
    s = CInt(Right$(strHora, 2))

    If h > 23 Or m > 59 Or s > 59 Then
        FormatoHora = ""
    Else
        FormatoHora = TimeSerial(h, m, s)
    End If
End Function

Private Sub CreateDraftFromPG2WithSignatureAndAttachment(ByVal rutaArchivo As String)
    On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(rutaArchivo) Then
        MsgBox "The file does not exist: " & rutaArchivo, vbCritical
        Exit Sub
    End If

    Dim firmaHTML As String
    firmaHTML = LeerFirmaHTML(FIRMA_W2)

    Dim olApp As Object, correo As Object, olAccount As Object
    Set olApp = CreateObject("Outlook.Application")
    Set correo = olApp.CreateItem(0)

    With correo
        .Subject = "RPA EMAIL REQUEST TO CHANGES"
        .To = "nacsoshared.im@pg.com"
        .htmlBody = "<p>Please find attached the updated appointment load file.</p>" & firmaHTML
        .Attachments.Add rutaArchivo

        For Each olAccount In olApp.Session.Accounts
            If InStr(1, olAccount.smtpAddress, "pgcustservw2.im@pg.com", vbTextCompare) > 0 Then
                Set .SendUsingAccount = olAccount
                Exit For
            End If
        Next olAccount

        .Display
    End With

    MsgBox "Draft RPA EMAIL REQUEST TO CHANGES created with attachment.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error creating attachment draft: " & Err.Description, vbCritical
End Sub

' ==================== Lanzador manual (solo pruebas) ====================
Public Sub DetectarTablasYColumnasDesdeSeleccion()
    On Error Resume Next
    Dim mail As Outlook.MailItem
    Set mail = Application.ActiveExplorer.selection.item(1)
    On Error GoTo 0

    If mail Is Nothing Then
        MsgBox "Please select one email first.", vbExclamation
        Exit Sub
    End If

    DetectarTablasYColumnas mail
End Sub
'==================== FIN ====================
