Attribute VB_Name = "SincronizzaBilancioSupabase"
Option Explicit

' Sincronizzazione diretta Excel <-> Supabase per Bilancio Casa.
' Non usa CSV esportati manualmente. Non cancella mai movimenti. (ramo di test)
' In assenza di snapshot, rileva i conflitti e si ferma senza sovrascrivere i dati.
' Richiede Excel 2013 e WinHTTP (presente in Windows).
' La password viene richiesta ad ogni avvio e non viene salvata nel file.

Private Const SUPABASE_URL As String = "https://bstcydrgyspgofckzybl.supabase.co"
Private Const SUPABASE_KEY As String = "sb_publishable_LURu7IVWMgryAlEq6upJKQ_wrTMfBJ6"
Private Const TABLE_NAME As String = "bilancio_movements"

Public Sub SincronizzaBilancioCasaDiretto()
    Dim email As String, password As String, token As String, userId As String
    Dim ws As Worksheet, hdr As Object, rowById As Object, apiRows As Object, newIds As Object
    Dim headerRow As Long, lastRow As Long, idCol As Long, c As Long, r As Long
    Dim id As String, body As String, url As String, response As String, remoteCsv As String
    Dim sent As Long, added As Long, downloaded As Long, updated As Long
    Dim oldCalc As XlCalculation, oldEvents As Boolean, oldScreen As Boolean
    Dim lo As ListObject, targetTable As ListObject, newRow As ListRow
    Dim data As Object, tmpWb As Workbook, tmpWs As Worksheet
    Dim tempPath As String, stm As Object, httpStatus As Long
    Dim errText As String

    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets("Movimenti")
    Set newIds = CreateObject("Scripting.Dictionary"): newIds.CompareMode = vbTextCompare
    email = InputBox("Inserisci l'email usata per accedere a Bilancio Casa:", "Sincronizzazione diretta")
    If Len(Trim$(email)) = 0 Then Exit Sub
    password = InputBox("Inserisci la password di Bilancio Casa. Non verrà salvata.", "Sincronizzazione diretta")
    If Len(password) = 0 Then Exit Sub

    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    response = HttpRequest("POST", SUPABASE_URL & "/auth/v1/token?grant_type=password", _
        "{""email"":" & JsonString(email) & ",""password"":" & JsonString(password) & "}", _
        "application/json", "", httpStatus)
    password = vbNullString
    If httpStatus < 200 Or httpStatus >= 300 Then Err.Raise vbObjectError + 701, , "Accesso Supabase non riuscito (" & httpStatus & "). Verifica email e password."
    token = JsonGetString(response, "access_token")
    userId = JsonGetObjectString(response, "user", "id")
    If Len(token) = 0 Or Len(userId) = 0 Then Err.Raise vbObjectError + 702, , "Risposta di accesso incompleta."

    headerRow = FindHeaderRow(ws, "idmovimento")
    If headerRow = 0 Then Err.Raise vbObjectError + 703, , "Non trovo la colonna ID movimento nel foglio Movimenti."
    Set hdr = CreateObject("Scripting.Dictionary"): hdr.CompareMode = vbTextCompare
    For c = 1 To ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
        If Len(NormalizeHeader(CStr(ws.Cells(headerRow, c).Value))) > 0 Then hdr(NormalizeHeader(CStr(ws.Cells(headerRow, c).Value))) = c
    Next c
    If Not hdr.Exists("data") Or Not hdr.Exists("concetto") Or Not hdr.Exists("categoria") Or Not hdr.Exists("tipo") Or Not hdr.Exists("importo") Or Not hdr.Exists("idmovimento") Then
        Err.Raise vbObjectError + 704, , "Mancano una o più colonne obbligatorie: ID movimento, Data, Concetto, Categoria, Tipo, Importo."
    End If
    idCol = hdr("idmovimento")
    Set targetTable = Nothing
    For Each lo In ws.ListObjects
        If lo.HeaderRowRange.Row = headerRow Then Set targetTable = lo: Exit For
    Next lo
    If Not targetTable Is Nothing Then
        If targetTable.DataBodyRange Is Nothing Then lastRow = headerRow Else lastRow = targetTable.DataBodyRange.Row + targetTable.DataBodyRange.Rows.Count - 1
    Else
        lastRow = ws.Cells(ws.Rows.Count, idCol).End(xlUp).Row
    End If

    ' Prima scarica il database: non inviare righe Excel con ID esistente senza confronto con uno snapshot.
    url = SUPABASE_URL & "/rest/v1/" & TABLE_NAME & "?select=id,movement_date,concept,category,movement_type,amount,notes&user_id=eq." & userId & "&order=movement_date.asc"
    response = HttpRequest("GET", url, "", "text/csv", token, httpStatus)
    If httpStatus < 200 Or httpStatus >= 300 Then Err.Raise vbObjectError + 710, , "Errore scaricando i movimenti (HTTP " & httpStatus & "): " & Left$(response, 300)
    remoteCsv = response

    For r = headerRow + 1 To lastRow
        If Len(Trim$(CStr(ws.Cells(r, hdr("concetto")).Value))) > 0 And IsDate(ws.Cells(r, hdr("data")).Value) Then
            If Not IsNumeric(ws.Cells(r, hdr("importo")).Value) Or CDbl(ws.Cells(r, hdr("importo")).Value) <= 0 Then
                Err.Raise vbObjectError + 705, , "Importo non valido alla riga " & r & ". Nessuna altra riga verrà sincronizzata."
            End If
            Dim movementType As String
            movementType = Trim$(CStr(ws.Cells(r, hdr("tipo")).Value))
            If movementType <> "Entrata" And movementType <> "Uscita" Then Err.Raise vbObjectError + 706, , "Tipo non valido alla riga " & r & ": usare Entrata o Uscita."
            body = "{""user_id"":" & JsonString(userId) & _
                ",""movement_date"":" & JsonString(Format$(CDate(ws.Cells(r, hdr("data")).Value), "yyyy-mm-dd")) & _
                ",""concept"":" & JsonString(CStr(ws.Cells(r, hdr("concetto")).Value)) & _
                ",""category"":" & JsonString(CStr(ws.Cells(r, hdr("categoria")).Value)) & _
                ",""movement_type"":" & JsonString(movementType) & _
                ",""amount"":" & Replace(Format$(CDbl(ws.Cells(r, hdr("importo")).Value), "0.00"), ",", ".") & _
                ",""notes"":" & JsonString(IIf(hdr.Exists("note"), CStr(ws.Cells(r, hdr("note")).Value), "")) & "}"
            id = Trim$(CStr(ws.Cells(r, idCol).Value))
            If Len(id) > 0 Then
                ' Rimandato: prima serve confronto con lo snapshot dell'ultima sincronizzazione riuscita.
            Else
                response = HttpRequest("POST", SUPABASE_URL & "/rest/v1/" & TABLE_NAME & "?select=id", body, "application/json", token, httpStatus, "return=representation")
                If httpStatus < 200 Or httpStatus >= 300 Then Err.Raise vbObjectError + 708, , "Errore inserendo la riga " & r & " (HTTP " & httpStatus & "): " & Left$(response, 300)
                id = JsonGetString(response, "id")
                If Len(id) = 0 Then Err.Raise vbObjectError + 709, , "Il server non ha restituito l'ID del nuovo movimento alla riga " & r
                ws.Cells(r, idCol).Value = id
                newIds(id) = True
                sent = sent + 1: added = added + 1
            End If
        End If
    Next r

    ' Il CSV remoto è già stato scaricato prima di qualsiasi invio.
    tempPath = Environ$("TEMP") & "\BilancioCasa_api_sync.csv"
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "utf-8": stm.Open: stm.WriteText remoteCsv: stm.SaveToFile tempPath, 2: stm.Close
    Workbooks.OpenText Filename:=tempPath, Origin:=65001, DataType:=xlDelimited, TextQualifier:=xlTextQualifierDoubleQuote, _
        ConsecutiveDelimiter:=False, Comma:=True, Semicolon:=False, Tab:=False, Local:=True
    Set tmpWb = ActiveWorkbook: Set tmpWs = tmpWb.Worksheets(1)
    Set apiRows = CreateObject("Scripting.Dictionary"): apiRows.CompareMode = vbTextCompare
    Dim apiHdr As Object, apiLast As Long, apiIdCol As Long, ac As Long, ar As Long, dest As Long, k As Variant
    Set apiHdr = CreateObject("Scripting.Dictionary"): apiHdr.CompareMode = vbTextCompare
    For ac = 1 To tmpWs.Cells(1, tmpWs.Columns.Count).End(xlToLeft).Column
        apiHdr(LCase$(Trim$(CStr(tmpWs.Cells(1, ac).Value)))) = ac
    Next ac
    If Not apiHdr.Exists("id") Then Err.Raise vbObjectError + 711, , "La risposta del database non contiene la colonna ID."
    apiIdCol = apiHdr("id"): apiLast = tmpWs.Cells(tmpWs.Rows.Count, apiIdCol).End(xlUp).Row

    ' Indicizza i record remoti prima di modificare il foglio Excel.
    For ar = 2 To apiLast
        id = Trim$(CStr(tmpWs.Cells(ar, apiIdCol).Value))
        If Len(id) > 0 Then apiRows(id) = ar
    Next ar

    Set rowById = CreateObject("Scripting.Dictionary"): rowById.CompareMode = vbTextCompare
    If Not targetTable Is Nothing Then
        If Not targetTable.DataBodyRange Is Nothing Then
            For r = targetTable.DataBodyRange.Row To targetTable.DataBodyRange.Row + targetTable.DataBodyRange.Rows.Count - 1
                id = Trim$(CStr(ws.Cells(r, idCol).Value)): If Len(id) > 0 Then If Not rowById.Exists(id) Then rowById.Add id, r
            Next r
        End If
    Else
        For r = headerRow + 1 To ws.Cells(ws.Rows.Count, idCol).End(xlUp).Row
            id = Trim$(CStr(ws.Cells(r, idCol).Value)): If Len(id) > 0 Then If Not rowById.Exists(id) Then rowById.Add id, r
        Next r
    End If

    ' Protezione anti-perdita: senza snapshot non si può decidere quale copia sia più recente.
    ' Se Excel e database differiscono su un ID già esistente, interrompe PRIMA di sovrascrivere Excel.
    For Each k In rowById.Keys
        id = CStr(k)
        If apiRows.Exists(id) Then
            If Not newIds.Exists(id) Then
                r = CLng(rowById(id)): ar = CLng(apiRows(id))
                If Not ExcelRowMatchesRemote(ws, r, hdr, tmpWs, ar, apiHdr) Then
                    Err.Raise vbObjectError + 712, , "Conflitto sul movimento ID " & id & ". Excel e app contengono dati diversi. Nessuna delle due versioni è stata sovrascritta; serve il confronto con lo snapshot."
                End If
            End If
        ElseIf Not newIds.Exists(id) Then
            Err.Raise vbObjectError + 713, , "L'ID " & id & " presente in Excel non risulta nel database. Sincronizzazione interrotta per evitare perdita di dati."
        End If
    Next k

    For ar = 2 To apiLast
        id = Trim$(CStr(tmpWs.Cells(ar, apiIdCol).Value))
        If Len(id) > 0 Then
            If rowById.Exists(id) Then
                dest = CLng(rowById(id)): updated = updated + 1
            Else
                If Not targetTable Is Nothing Then
                    Set newRow = targetTable.ListRows.Add: dest = newRow.Range.Row
                Else
                    dest = ws.Cells(ws.Rows.Count, idCol).End(xlUp).Row + 1
                End If
                rowById.Add id, dest: added = added + 1
            End If
            PutValue ws, dest, hdr, "idmovimento", id
            PutValue ws, dest, hdr, "data", IsoDateValue(CStr(tmpWs.Cells(ar, apiHdr("movement_date")).Value))
            PutValue ws, dest, hdr, "concetto", tmpWs.Cells(ar, apiHdr("concept")).Value
            PutValue ws, dest, hdr, "categoria", tmpWs.Cells(ar, apiHdr("category")).Value
            PutValue ws, dest, hdr, "tipo", tmpWs.Cells(ar, apiHdr("movement_type")).Value
            PutValue ws, dest, hdr, "importo", Val(Replace(CStr(tmpWs.Cells(ar, apiHdr("amount")).Value), ",", "."))
            PutValue ws, dest, hdr, "note", tmpWs.Cells(ar, apiHdr("notes")).Value
            downloaded = downloaded + 1
        End If
    Next ar

    tmpWb.Close SaveChanges:=False
    If Len(Dir$(tempPath)) > 0 Then Kill tempPath
    Application.Calculation = oldCalc: Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    MsgBox "Sincronizzazione diretta completata." & vbCrLf & _
        "Nuovi movimenti inviati da Excel: " & sent & vbCrLf & "Nuovi ID assegnati: " & added & vbCrLf & _
        "Movimenti ricevuti dal database: " & downloaded & vbCrLf & _
        "Nessun movimento è stato cancellato." & vbCrLf & _
        "Nota: le modifiche su ID esistenti non vengono ancora propagate. Se Excel e app differiscono, la macro si ferma per proteggere entrambe le versioni.", vbInformation
    Exit Sub

Fail:
    errText = "Errore " & Err.Number & ": " & Err.Description
    On Error Resume Next
    password = vbNullString
    If Not tmpWb Is Nothing Then tmpWb.Close SaveChanges:=False
    If Len(tempPath) > 0 Then If Len(Dir$(tempPath)) > 0 Then Kill tempPath
    Application.Calculation = oldCalc: Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    MsgBox errText, vbExclamation, "Sincronizzazione Bilancio Casa"
End Sub

Private Function HttpRequest(ByVal method As String, ByVal url As String, ByVal body As String, _
    ByVal acceptType As String, ByVal bearer As String, ByRef status As Long, Optional ByVal prefer As String = "") As String
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open method, url, False
    http.SetRequestHeader "apikey", SUPABASE_KEY
    If Len(bearer) > 0 Then http.SetRequestHeader "Authorization", "Bearer " & bearer
    http.SetRequestHeader "Accept", acceptType
    If method = "POST" Or method = "PATCH" Or method = "PUT" Then http.SetRequestHeader "Content-Type", "application/json"
    If Len(prefer) > 0 Then http.SetRequestHeader "Prefer", prefer
    http.Send body
    status = http.Status
    HttpRequest = http.ResponseText
End Function

Private Function JsonString(ByVal s As String) As String
    s = Replace(s, "\", "\\"): s = Replace(s, """", "\"""")
    s = Replace(s, vbCr, "\r"): s = Replace(s, vbLf, "\n"): s = Replace(s, vbTab, "\t")
    JsonString = """" & s & """"
End Function

Private Function JsonGetString(ByVal json As String, ByVal key As String) As String
    Dim p As Long, q As Long, i As Long, ch As String, out As String
    p = InStr(1, json, """" & key & """", vbTextCompare): If p = 0 Then Exit Function
    p = InStr(p, json, ":"): If p = 0 Then Exit Function
    p = InStr(p, json, """", vbBinaryCompare): If p = 0 Then Exit Function
    i = p + 1
    Do While i <= Len(json)
        ch = Mid$(json, i, 1)
        If ch = """" Then Exit Do
        If ch = "\" And i < Len(json) Then
            i = i + 1: ch = Mid$(json, i, 1)
            Select Case ch
                Case "n": out = out & vbLf
                Case "r": out = out & vbCr
                Case "t": out = out & vbTab
                Case Else: out = out & ch
            End Select
        Else
            out = out & ch
        End If
        i = i + 1
    Loop
    JsonGetString = out
End Function

Private Function JsonGetObjectString(ByVal json As String, ByVal objectName As String, ByVal key As String) As String
    Dim p As Long, q As Long, piece As String
    p = InStr(1, json, """" & objectName & """", vbTextCompare): If p = 0 Then Exit Function
    p = InStr(p, json, "{"): If p = 0 Then Exit Function
    q = InStr(p, json, "}"): If q = 0 Then Exit Function
    piece = Mid$(json, p, q - p + 1)
    JsonGetObjectString = JsonGetString(piece, key)
End Function

Private Function FindHeaderRow(ByVal ws As Worksheet, ByVal wanted As String) As Long
    Dim lo As ListObject, r As Long, c As Long
    For Each lo In ws.ListObjects
        For c = 1 To lo.ListColumns.Count
            If NormalizeHeader(CStr(lo.HeaderRowRange.Cells(1, c).Value)) = wanted Then FindHeaderRow = lo.HeaderRowRange.Row: Exit Function
        Next c
    Next lo
    For r = 1 To 20
        For c = 1 To ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
            If NormalizeHeader(CStr(ws.Cells(r, c).Value)) = wanted Then FindHeaderRow = r: Exit Function
        Next c
    Next r
End Function

Private Function NormalizeHeader(ByVal s As String) As String
    s = Replace(s, ChrW(65279), ""): s = Replace(s, ChrW(160), " ")
    s = LCase$(Trim$(s)): s = Replace(s, " ", "")
    s = Replace(s, "+", "piu"): s = Replace(s, "-", ""): s = Replace(s, "/", "")
    NormalizeHeader = s
End Function

Private Sub PutValue(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal hdr As Object, ByVal key As String, ByVal value As Variant)
    If hdr.Exists(key) Then ws.Cells(rowNum, hdr(key)).Value = value
End Sub

Private Function ExcelRowMatchesRemote(ByVal ws As Worksheet, ByVal excelRow As Long, ByVal hdr As Object, _
    ByVal remoteWs As Worksheet, ByVal remoteRow As Long, ByVal apiHdr As Object) As Boolean
    On Error GoTo NotMatch
    If Format$(CDate(ws.Cells(excelRow, hdr("data")).Value), "yyyy-mm-dd") <> Left$(CStr(remoteWs.Cells(remoteRow, apiHdr("movement_date")).Value), 10) Then GoTo NotMatch
    If CStr(ws.Cells(excelRow, hdr("concetto")).Value) <> CStr(remoteWs.Cells(remoteRow, apiHdr("concept")).Value) Then GoTo NotMatch
    If CStr(ws.Cells(excelRow, hdr("categoria")).Value) <> CStr(remoteWs.Cells(remoteRow, apiHdr("category")).Value) Then GoTo NotMatch
    If CStr(ws.Cells(excelRow, hdr("tipo")).Value) <> CStr(remoteWs.Cells(remoteRow, apiHdr("movement_type")).Value) Then GoTo NotMatch
    If Abs(CDbl(ws.Cells(excelRow, hdr("importo")).Value) - Val(Replace(CStr(remoteWs.Cells(remoteRow, apiHdr("amount")).Value), ",", "."))) > 0.005 Then GoTo NotMatch
    If hdr.Exists("note") Then
        If CStr(ws.Cells(excelRow, hdr("note")).Value) <> CStr(remoteWs.Cells(remoteRow, apiHdr("notes")).Value) Then GoTo NotMatch
    End If
    ExcelRowMatchesRemote = True
    Exit Function
NotMatch:
    ExcelRowMatchesRemote = False
End Function

Private Function IsoDateValue(ByVal s As String) As Variant
    s = Trim$(s)
    If Len(s) >= 10 And Mid$(s, 5, 1) = "-" And Mid$(s, 8, 1) = "-" Then
        IsoDateValue = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Mid$(s, 9, 2)))
    ElseIf IsDate(s) Then
        IsoDateValue = CDate(s)
    Else
        IsoDateValue = s
    End If
End Function

