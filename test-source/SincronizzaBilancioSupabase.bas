Option Explicit

' Stato temporaneo: resta solo in memoria mentre Excel è aperto.
Private gRefreshToken As String
Private gCachedAccessToken As String
Private gCachedUserId As String
Private gCachedEmail As String
Private gAutoSyncEnabled As Boolean
Private gSyncBusy As Boolean
Private gTimerTriggered As Boolean
Private gNextSyncTime As Date

' Sincronizzazione diretta Excel <-> Supabase per Bilancio Casa.
' Non usa CSV esportati manualmente. Non cancella mai movimenti. (ramo di test)
' In assenza di snapshot, rileva i conflitti e si ferma senza sovrascrivere i dati.
' Richiede Excel 2013 e WinHTTP (presente in Windows).
' Password non salvata; il refresh token resta solo in memoria durante la sessione Excel.

Private Const SUPABASE_URL As String = "https://bstcydrgyspgofckzybl.supabase.co"
Private Const SUPABASE_KEY As String = "sb_publishable_LURu7IVWMgryAlEq6upJKQ_wrTMfBJ6"
Private Const TABLE_NAME As String = "bilancio_movements"

Public Sub SincronizzaBilancioCasaDiretto()
    Dim email As String, password As String, token As String, userId As String
    Dim ws As Worksheet, snapshotWs As Worksheet, snapshot As Object, hdr As Object, rowById As Object, apiRows As Object, newIds As Object
    Dim headerRow As Long, lastRow As Long, idCol As Long, c As Long, r As Long
    Dim id As String, body As String, url As String, response As String, remoteCsv As String
    Dim sent As Long, added As Long, downloaded As Long, updated As Long
    Dim oldCalc As XlCalculation, oldEvents As Boolean, oldScreen As Boolean, settingsSaved As Boolean
    Dim lo As ListObject, targetTable As ListObject, newRow As ListRow
    Dim data As Object, tmpWb As Workbook, tmpWs As Worksheet
    Dim excelChanged As Boolean, remoteChanged As Boolean, snapVals As Variant
    Dim tempPath As String, stm As Object, httpStatus As Long
    Dim errText As String, wasAutoSync As Boolean, workbookWasSaved As Boolean
    Dim movementType As String
    Dim apiHdr As Object, apiLast As Long, apiIdCol As Long, ac As Long, ar As Long, dest As Long, k As Variant

    On Error GoTo Fail
    workbookWasSaved = ThisWorkbook.Saved
    Set ws = ThisWorkbook.Worksheets("Movimenti")
    Set snapshotWs = GetOrCreateSyncSheet(ThisWorkbook)
    Set snapshot = LoadSyncSnapshot(snapshotWs)
    Set newIds = CreateObject("Scripting.Dictionary"): newIds.CompareMode = vbTextCompare
    If gSyncBusy Then Exit Sub
    gSyncBusy = True

    If Len(gRefreshToken) > 0 Then
        response = HttpRequest("POST", SUPABASE_URL & "/auth/v1/token?grant_type=refresh_token", _
            "{""refresh_token"":" & JsonString(gRefreshToken) & "}", _
            "application/json", "", httpStatus)
        If httpStatus >= 200 And httpStatus < 300 Then
            token = JsonGetString(response, "access_token")
            If Len(JsonGetString(response, "refresh_token")) > 0 Then gRefreshToken = JsonGetString(response, "refresh_token")
            If Len(JsonGetObjectString(response, "user", "id")) > 0 Then gCachedUserId = JsonGetObjectString(response, "user", "id")
            If Len(token) > 0 And Len(gCachedUserId) > 0 Then
                userId = gCachedUserId
                gCachedAccessToken = token
            Else
                If gTimerTriggered Then Err.Raise vbObjectError + 703, , "La sessione automatica non è più valida. Esegui una sincronizzazione manuale per accedere di nuovo."
                gRefreshToken = vbNullString
            End If
        Else
            If gTimerTriggered Then Err.Raise vbObjectError + 704, , "La sessione automatica è scaduta o non è rinnovabile. Esegui una sincronizzazione manuale per accedere di nuovo."
            gRefreshToken = vbNullString
        End If
    End If

    If Len(gRefreshToken) = 0 Then
        If gTimerTriggered Then Err.Raise vbObjectError + 705, , "Per riprendere la sincronizzazione automatica, esegui prima una sincronizzazione manuale e accedi di nuovo."
        email = InputBox("Inserisci l'email usata per accedere a Bilancio Casa:", "Sincronizzazione diretta")
        If Len(Trim$(email)) = 0 Then GoTo Cancelled
        password = InputBox("Inserisci la password di Bilancio Casa. Non verrà salvata.", "Sincronizzazione diretta")
        If Len(password) = 0 Then GoTo Cancelled
        response = HttpRequest("POST", SUPABASE_URL & "/auth/v1/token?grant_type=password", _
            "{""email"":" & JsonString(email) & ",""password"":" & JsonString(password) & "}", _
            "application/json", "", httpStatus)
        password = vbNullString
        If httpStatus < 200 Or httpStatus >= 300 Then Err.Raise vbObjectError + 701, , "Accesso Supabase non riuscito (" & httpStatus & "). Verifica email e password."
        token = JsonGetString(response, "access_token")
        gRefreshToken = JsonGetString(response, "refresh_token")
        userId = JsonGetObjectString(response, "user", "id")
        If Len(token) = 0 Or Len(userId) = 0 Or Len(gRefreshToken) = 0 Then Err.Raise vbObjectError + 702, , "Risposta di accesso incompleta."
        gCachedAccessToken = token
        gCachedUserId = userId
        gCachedEmail = email
    Else
        token = gCachedAccessToken
        userId = gCachedUserId
    End If

    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    settingsSaved = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Token e ID utente già ottenuti o rinnovati nel blocco di autenticazione sopra.

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
        lastRow = Application.Max(ws.Cells(ws.Rows.Count, idCol).End(xlUp).Row, ws.Cells(ws.Rows.Count, hdr("concetto")).End(xlUp).Row)
    End If

    ' Prima scarica il database: non inviare righe Excel con ID esistente senza confronto con uno snapshot.
    url = SUPABASE_URL & "/rest/v1/" & TABLE_NAME & "?select=id,movement_date,concept,category,movement_type,amount,notes&user_id=eq." & userId & "&order=movement_date.asc"
    response = HttpRequest("GET", url, "", "text/csv", token, httpStatus)
    If httpStatus < 200 Or httpStatus >= 300 Then Err.Raise vbObjectError + 710, , "Errore scaricando i movimenti (HTTP " & httpStatus & "): " & Left$(response, 300)
    remoteCsv = response

    ' Pre-controllo dei dati Excel: nessuna scrittura al database in questa fase.
    For r = headerRow + 1 To lastRow
        If Len(Trim$(CStr(ws.Cells(r, hdr("concetto")).Value))) > 0 And IsDate(ws.Cells(r, hdr("data")).Value) And LCase$(Trim$(CStr(ws.Cells(r, hdr("tipo")).Value))) <> "saldo iniziale" Then
            If Not IsNumeric(ws.Cells(r, hdr("importo")).Value) Or CDbl(ws.Cells(r, hdr("importo")).Value) <= 0 Then
                Err.Raise vbObjectError + 705, , "Importo non valido alla riga " & r & ". Nessuna riga verrà sincronizzata."
            End If
            movementType = Trim$(CStr(ws.Cells(r, hdr("tipo")).Value))
            If movementType <> "Entrata" And movementType <> "Uscita" Then Err.Raise vbObjectError + 706, , "Tipo non valido alla riga " & r & ": usare Entrata o Uscita."
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
    Set apiHdr = CreateObject("Scripting.Dictionary"): apiHdr.CompareMode = vbTextCompare
    For ac = 1 To tmpWs.Cells(1, tmpWs.Columns.Count).End(xlToLeft).Column
        apiHdr(LCase$(Trim$(CStr(tmpWs.Cells(1, ac).Value)))) = ac
    Next ac
    If Not apiHdr.Exists("id") Then Err.Raise vbObjectError + 711, , "La risposta del database non contiene la colonna ID."
    apiIdCol = apiHdr("id"): apiLast = tmpWs.Cells(tmpWs.Rows.Count, apiIdCol).End(xlUp).Row

    ' Indicizza i record remoti prima di modificare il foglio Excel.
    For ar = 2 To apiLast
        id = Trim$(CStr(tmpWs.Cells(ar, apiIdCol).Value))
        If Len(id) > 0 Then
            If apiRows.Exists(id) Then Err.Raise vbObjectError + 715, , "ID duplicato nella risposta Supabase: " & id & ". Nessuna sincronizzazione prosegue."
            apiRows.Add id, ar
        End If
    Next ar

    Set rowById = CreateObject("Scripting.Dictionary"): rowById.CompareMode = vbTextCompare
    If Not targetTable Is Nothing Then
        If Not targetTable.DataBodyRange Is Nothing Then
            For r = targetTable.DataBodyRange.Row To targetTable.DataBodyRange.Row + targetTable.DataBodyRange.Rows.Count - 1
                id = Trim$(CStr(ws.Cells(r, idCol).Value))
                If Len(id) > 0 Then
                    If rowById.Exists(id) Then Err.Raise vbObjectError + 716, , "ID duplicato in Excel: " & id & " (righe " & rowById(id) & " e " & r & "). Nessuna sincronizzazione prosegue."
                    rowById.Add id, r
                End If
            Next r
        End If
    Else
        For r = headerRow + 1 To ws.Cells(ws.Rows.Count, idCol).End(xlUp).Row
            id = Trim$(CStr(ws.Cells(r, idCol).Value))
            If Len(id) > 0 Then
                If rowById.Exists(id) Then Err.Raise vbObjectError + 716, , "ID duplicato in Excel: " & id & " (righe " & rowById(id) & " e " & r & "). Nessuna sincronizzazione prosegue."
                rowById.Add id, r
            End If
        Next r
    End If

    ' Confronta le due copie con l'ultimo snapshot riuscito.
    For Each k In rowById.Keys
        id = CStr(k)
        If apiRows.Exists(id) Then
            If Not newIds.Exists(id) Then
                r = CLng(rowById(id)): ar = CLng(apiRows(id))
                If snapshot.Exists(id) Then
                    snapVals = snapshot(id)
                    excelChanged = Not ExcelRowMatchesSnapshot(ws, r, hdr, snapVals)
                    remoteChanged = Not RemoteRowMatchesSnapshot(tmpWs, ar, apiHdr, snapVals)
                    If excelChanged And remoteChanged Then
                        If Not ExcelRowMatchesRemote(ws, r, hdr, tmpWs, ar, apiHdr) Then
                            Err.Raise vbObjectError + 712, , "Conflitto sul movimento ID " & id & ". Excel e app sono stati modificati in modo diverso dopo l'ultima sincronizzazione. Nessuna delle due versioni è stata sovrascritta."
                        End If
                    ElseIf excelChanged And Not remoteChanged Then
                        ' Aggiornamento differito: prima si controllano tutti i conflitti.
                    End If
                ElseIf Not ExcelRowMatchesRemote(ws, r, hdr, tmpWs, ar, apiHdr) Then
                    Err.Raise vbObjectError + 712, , "Conflitto sul movimento ID " & id & ". Campi diversi: " & DescribeRowDifferences(ws, r, hdr, tmpWs, ar, apiHdr) & ". Nessuno dei due dati è stato sovrascritto."
                End If
            End If
        ElseIf Not newIds.Exists(id) Then
            Err.Raise vbObjectError + 713, , "L'ID " & id & " presente in Excel non risulta nel database. Sincronizzazione interrotta per evitare perdita di dati."
        End If
    Next k

    ' Seconda fase: applica le modifiche Excel solo dopo che tutti i conflitti sono stati esclusi.
    For Each k In rowById.Keys
        id = CStr(k)
        If apiRows.Exists(id) And snapshot.Exists(id) Then
            If Not newIds.Exists(id) Then
                r = CLng(rowById(id)): ar = CLng(apiRows(id))
                snapVals = snapshot(id)
                excelChanged = Not ExcelRowMatchesSnapshot(ws, r, hdr, snapVals)
                remoteChanged = Not RemoteRowMatchesSnapshot(tmpWs, ar, apiHdr, snapVals)
                If excelChanged And Not remoteChanged Then
                    If Not PatchMovementFromExcel(ws, r, hdr, id, userId, token, httpStatus, response) Then
                        Err.Raise vbObjectError + 714, , "Aggiornamento Excel -> app non riuscito per ID " & id & " (HTTP " & httpStatus & "): " & Left$(response, 250)
                    End If
                    tmpWs.Cells(ar, apiHdr("movement_date")).Value = Format$(CDate(ws.Cells(r, hdr("data")).Value), "yyyy-mm-dd")
                    tmpWs.Cells(ar, apiHdr("concept")).Value = ws.Cells(r, hdr("concetto")).Value
                    tmpWs.Cells(ar, apiHdr("category")).Value = ws.Cells(r, hdr("categoria")).Value
                    tmpWs.Cells(ar, apiHdr("movement_type")).Value = ws.Cells(r, hdr("tipo")).Value
                    tmpWs.Cells(ar, apiHdr("amount")).Value = CDbl(ws.Cells(r, hdr("importo")).Value)
                    tmpWs.Cells(ar, apiHdr("notes")).Value = OptionalCellText(ws, r, hdr, "note")
                End If
            End If
        End If
    Next k

    ' Solo dopo aver verificato tutti gli ID esistenti, invia le righe Excel senza ID.
    ' Le nuove righe vengono aggiunte anche alla copia remota in memoria per la riconciliazione finale.
    For r = headerRow + 1 To lastRow
        If Len(Trim$(CStr(ws.Cells(r, hdr("concetto")).Value))) > 0 And IsDate(ws.Cells(r, hdr("data")).Value) And LCase$(Trim$(CStr(ws.Cells(r, hdr("tipo")).Value))) <> "saldo iniziale" Then
            id = Trim$(CStr(ws.Cells(r, idCol).Value))
            If Len(id) = 0 Then
                body = "{""user_id"":" & JsonString(userId) & _
                    ",""movement_date"":" & JsonString(Format$(CDate(ws.Cells(r, hdr("data")).Value), "yyyy-mm-dd")) & _
                    ",""concept"":" & JsonString(CStr(ws.Cells(r, hdr("concetto")).Value)) & _
                    ",""category"":" & JsonString(CStr(ws.Cells(r, hdr("categoria")).Value)) & _
                    ",""movement_type"":" & JsonString(Trim$(CStr(ws.Cells(r, hdr("tipo")).Value))) & _
                    ",""amount"":" & Replace(Format$(CDbl(ws.Cells(r, hdr("importo")).Value), "0.00"), ",", ".") & _
                    ",""notes"":" & JsonString(OptionalCellText(ws, r, hdr, "note")) & "}"
                response = HttpRequest("POST", SUPABASE_URL & "/rest/v1/" & TABLE_NAME & "?select=id", body, "application/json", token, httpStatus, "return=representation")
                If httpStatus < 200 Or httpStatus >= 300 Then Err.Raise vbObjectError + 708, , "Errore inserendo la riga " & r & " (HTTP " & httpStatus & "): " & Left$(response, 300)
                id = JsonGetString(response, "id")
                If Len(id) = 0 Then Err.Raise vbObjectError + 709, , "Il server non ha restituito l'ID del nuovo movimento alla riga " & r
                ws.Cells(r, idCol).Value = id
                newIds(id) = True
                rowById.Add id, r
                apiLast = apiLast + 1
                ar = apiLast
                tmpWs.Cells(ar, apiHdr("id")).Value = id
                tmpWs.Cells(ar, apiHdr("movement_date")).Value = Format$(CDate(ws.Cells(r, hdr("data")).Value), "yyyy-mm-dd")
                tmpWs.Cells(ar, apiHdr("concept")).Value = ws.Cells(r, hdr("concetto")).Value
                tmpWs.Cells(ar, apiHdr("category")).Value = ws.Cells(r, hdr("categoria")).Value
                tmpWs.Cells(ar, apiHdr("movement_type")).Value = ws.Cells(r, hdr("tipo")).Value
                tmpWs.Cells(ar, apiHdr("amount")).Value = CDbl(ws.Cells(r, hdr("importo")).Value)
                tmpWs.Cells(ar, apiHdr("notes")).Value = OptionalCellText(ws, r, hdr, "note")
                apiRows(id) = ar
                sent = sent + 1
            End If
        End If
    Next r

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

    SaveSyncSnapshot snapshotWs, ws, hdr, rowById
    tmpWb.Close SaveChanges:=False
    If Len(Dir$(tempPath)) > 0 Then Kill tempPath
    Application.Calculation = oldCalc: Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen

    ' Salva automaticamente solo se il file era gia salvato prima del ciclo.
    ' Se l'utente aveva modifiche pendenti, non le rendere permanenti senza consenso.
    If workbookWasSaved Then ThisWorkbook.Save

    gSyncBusy = False
    If gAutoSyncEnabled Then
        PianificaProssimaSincronizzazione
    Else
        MsgBox "Sincronizzazione diretta completata." & vbCrLf & _
            "Nuovi movimenti inviati da Excel: " & sent & vbCrLf & "Movimenti aggiunti in Excel: " & added & vbCrLf & _
            "Movimenti ricevuti dal database: " & downloaded & vbCrLf & _
            "Nessun movimento è stato cancellato." & vbCrLf & _
            "Le modifiche univoche sono state riconciliate; in caso di conflitto la macro si ferma senza sovrascrivere le versioni diverse.", vbInformation
    End If
    Exit Sub

Cancelled:
    gSyncBusy = False
    If gAutoSyncEnabled Then FermaSincronizzazioneAutomatica True
    Exit Sub

Fail:
    wasAutoSync = gAutoSyncEnabled
    errText = "Errore " & Err.Number & ": " & Err.Description
    On Error Resume Next
    password = vbNullString
    If Not tmpWb Is Nothing Then tmpWb.Close SaveChanges:=False
    If Len(tempPath) > 0 Then
        If Len(Dir$(tempPath)) > 0 Then Kill tempPath
    End If
    If settingsSaved Then
        Application.Calculation = oldCalc
        Application.EnableEvents = oldEvents
        Application.ScreenUpdating = oldScreen
    End If
    gSyncBusy = False
    If gAutoSyncEnabled Then
        gAutoSyncEnabled = False
        On Error Resume Next
        If gNextSyncTime <> 0 Then Application.OnTime EarliestTime:=gNextSyncTime, Procedure:="SincronizzazioneBilancioCasaTimer", Schedule:=False
        On Error GoTo 0
    End If
    On Error GoTo 0
    MsgBox errText & IIf(wasAutoSync, vbCrLf & "Automazione arrestata.", ""), vbExclamation, "Sincronizzazione Bilancio Casa"
End Sub

' Avvia una sincronizzazione immediata e poi ripete ogni 15 minuti.
Public Sub AvviaSincronizzazioneAutomatica()
    gAutoSyncEnabled = True
    SincronizzaBilancioCasaDiretto
End Sub

' Arresta il timer; non modifica i dati già sincronizzati.
Public Sub FermaSincronizzazioneAutomatica(Optional ByVal silenzioso As Boolean = False)
    gAutoSyncEnabled = False
    On Error Resume Next
    If gNextSyncTime <> 0 Then Application.OnTime EarliestTime:=gNextSyncTime, Procedure:="SincronizzazioneBilancioCasaTimer", Schedule:=False
    On Error GoTo 0
    gNextSyncTime = 0
    If Not silenzioso Then MsgBox "Sincronizzazione automatica arrestata.", vbInformation
End Sub

Public Sub SincronizzazioneBilancioCasaTimer()
    If Not gAutoSyncEnabled Or gSyncBusy Then Exit Sub
    gTimerTriggered = True
    SincronizzaBilancioCasaDiretto
    gTimerTriggered = False
End Sub

Private Sub PianificaProssimaSincronizzazione()
    On Error Resume Next
    If gNextSyncTime <> 0 Then Application.OnTime EarliestTime:=gNextSyncTime, Procedure:="SincronizzazioneBilancioCasaTimer", Schedule:=False
    On Error GoTo 0
    gNextSyncTime = Now + TimeSerial(0, 15, 0)
    Application.OnTime EarliestTime:=gNextSyncTime, Procedure:="SincronizzazioneBilancioCasaTimer", Schedule:=True
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
    s = Replace(s, Chr$(92), Chr$(92) & Chr$(92))
    s = Replace(s, Chr$(34), Chr$(92) & Chr$(34))
    s = Replace(s, vbCr, Chr$(92) & "r"): s = Replace(s, vbLf, Chr$(92) & "n"): s = Replace(s, vbTab, Chr$(92) & "t")
    JsonString = Chr$(34) & s & Chr$(34)
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

Private Function OptionalCellText(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal hdr As Object, ByVal key As String) As String
    If hdr.Exists(key) Then
        If Not IsError(ws.Cells(rowNum, hdr(key)).Value) Then
            OptionalCellText = CStr(ws.Cells(rowNum, hdr(key)).Value)
        End If
    Else
        OptionalCellText = ""
    End If
End Function

Private Function RemoteDateISO(ByVal v As Variant) As String
    ' OpenText può convertire la data ISO del CSV in una data Excel locale.
    ' Normalizza entrambe le rappresentazioni nello stesso formato ISO.
    If IsDate(v) Then
        RemoteDateISO = Format$(CDate(v), "yyyy-mm-dd")
    Else
        RemoteDateISO = Left$(Trim$(CStr(v)), 10)
    End If
End Function


Private Function ExcelRowMatchesRemote(ByVal ws As Worksheet, ByVal excelRow As Long, ByVal hdr As Object, _
    ByVal remoteWs As Worksheet, ByVal remoteRow As Long, ByVal apiHdr As Object) As Boolean
    On Error GoTo NotMatch
    If Format$(CDate(ws.Cells(excelRow, hdr("data")).Value), "yyyy-mm-dd") <> RemoteDateISO(remoteWs.Cells(remoteRow, apiHdr("movement_date")).Value) Then GoTo NotMatch
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


Private Function DescribeRowDifferences(ByVal ws As Worksheet, ByVal excelRow As Long, ByVal hdr As Object, _
    ByVal remoteWs As Worksheet, ByVal remoteRow As Long, ByVal apiHdr As Object) As String
    Dim result As String
    On Error GoTo SafeExit
    If Format$(CDate(ws.Cells(excelRow, hdr("data")).Value), "yyyy-mm-dd") <> RemoteDateISO(remoteWs.Cells(remoteRow, apiHdr("movement_date")).Value) Then result = result & "Data; "
    If Trim$(CStr(ws.Cells(excelRow, hdr("concetto")).Value)) <> Trim$(CStr(remoteWs.Cells(remoteRow, apiHdr("concept")).Value)) Then result = result & "Concetto; "
    If Trim$(CStr(ws.Cells(excelRow, hdr("categoria")).Value)) <> Trim$(CStr(remoteWs.Cells(remoteRow, apiHdr("category")).Value)) Then result = result & "Categoria; "
    If LCase$(Trim$(CStr(ws.Cells(excelRow, hdr("tipo")).Value))) <> LCase$(Trim$(CStr(remoteWs.Cells(remoteRow, apiHdr("movement_type")).Value))) Then result = result & "Tipo; "
    If Abs(CDbl(ws.Cells(excelRow, hdr("importo")).Value) - Val(Replace(CStr(remoteWs.Cells(remoteRow, apiHdr("amount")).Value), ",", "."))) > 0.005 Then result = result & "Importo; "
    If hdr.Exists("note") Then
        If Trim$(CStr(ws.Cells(excelRow, hdr("note")).Value)) <> Trim$(CStr(remoteWs.Cells(remoteRow, apiHdr("notes")).Value)) Then result = result & "Note; "
    End If
SafeExit:
    If Len(result) = 0 Then result = "differenza non identificata"
    If Right$(result, 2) = "; " Then result = Left$(result, Len(result) - 2)
    DescribeRowDifferences = result
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


Private Function GetOrCreateSyncSheet(ByVal wb As Workbook) As Worksheet
    On Error Resume Next
    Set GetOrCreateSyncSheet = wb.Worksheets("_SyncBilancio")
    On Error GoTo 0
    If GetOrCreateSyncSheet Is Nothing Then
        Set GetOrCreateSyncSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        GetOrCreateSyncSheet.Name = "_SyncBilancio"
        With GetOrCreateSyncSheet
            .Cells(1, 1).Value = "id"
            .Cells(1, 2).Value = "data"
            .Cells(1, 3).Value = "concetto"
            .Cells(1, 4).Value = "categoria"
            .Cells(1, 5).Value = "tipo"
            .Cells(1, 6).Value = "importo"
            .Cells(1, 7).Value = "note"
            .Visible = xlSheetVeryHidden
        End With
    End If
End Function


Private Function LoadSyncSnapshot(ByVal snapWs As Worksheet) As Object
    Dim d As Object, last As Long, r As Long, id As String
    Dim vals(1 To 6) As Variant
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    last = snapWs.Cells(snapWs.Rows.Count, 1).End(xlUp).Row
    For r = 2 To last
        id = Trim$(CStr(snapWs.Cells(r, 1).Value))
        If Len(id) > 0 Then
            If d.Exists(id) Then Err.Raise vbObjectError + 717, , "ID duplicato nello snapshot _SyncBilancio: " & id & ". Non è sicuro continuare."
            vals(1) = snapWs.Cells(r, 2).Value
            vals(2) = snapWs.Cells(r, 3).Value
            vals(3) = snapWs.Cells(r, 4).Value
            vals(4) = snapWs.Cells(r, 5).Value
            vals(5) = snapWs.Cells(r, 6).Value
            vals(6) = snapWs.Cells(r, 7).Value
            d(id) = vals
        End If
    Next r
    Set LoadSyncSnapshot = d
End Function


Private Sub SaveSyncSnapshot(ByVal snapWs As Worksheet, ByVal ws As Worksheet, ByVal hdr As Object, ByVal rowById As Object)
    Dim k As Variant, r As Long, dest As Long, id As String
    snapWs.Cells.ClearContents
    snapWs.Cells(1, 1).Value = "id"
    snapWs.Cells(1, 2).Value = "data"
    snapWs.Cells(1, 3).Value = "concetto"
    snapWs.Cells(1, 4).Value = "categoria"
    snapWs.Cells(1, 5).Value = "tipo"
    snapWs.Cells(1, 6).Value = "importo"
    snapWs.Cells(1, 7).Value = "note"
    dest = 2
    For Each k In rowById.Keys
        id = CStr(k): r = CLng(rowById(id))
        snapWs.Cells(dest, 1).Value = id
        snapWs.Cells(dest, 2).Value = ws.Cells(r, hdr("data")).Value
        snapWs.Cells(dest, 3).Value = ws.Cells(r, hdr("concetto")).Value
        snapWs.Cells(dest, 4).Value = ws.Cells(r, hdr("categoria")).Value
        snapWs.Cells(dest, 5).Value = ws.Cells(r, hdr("tipo")).Value
        snapWs.Cells(dest, 6).Value = CDbl(ws.Cells(r, hdr("importo")).Value)
        snapWs.Cells(dest, 7).Value = OptionalCellText(ws, r, hdr, "note")
        dest = dest + 1
    Next k
    snapWs.Visible = xlSheetVeryHidden
End Sub


Private Function ExcelRowMatchesSnapshot(ByVal ws As Worksheet, ByVal excelRow As Long, ByVal hdr As Object, ByVal vals As Variant) As Boolean
    On Error GoTo NotMatch
    If Format$(CDate(ws.Cells(excelRow, hdr("data")).Value), "yyyy-mm-dd") <> Format$(CDate(vals(1)), "yyyy-mm-dd") Then GoTo NotMatch
    If CStr(ws.Cells(excelRow, hdr("concetto")).Value) <> CStr(vals(2)) Then GoTo NotMatch
    If CStr(ws.Cells(excelRow, hdr("categoria")).Value) <> CStr(vals(3)) Then GoTo NotMatch
    If CStr(ws.Cells(excelRow, hdr("tipo")).Value) <> CStr(vals(4)) Then GoTo NotMatch
    If Abs(CDbl(ws.Cells(excelRow, hdr("importo")).Value) - CDbl(vals(5))) > 0.005 Then GoTo NotMatch
    If hdr.Exists("note") Then
        If CStr(ws.Cells(excelRow, hdr("note")).Value) <> CStr(vals(6)) Then GoTo NotMatch
    End If
    ExcelRowMatchesSnapshot = True
    Exit Function
NotMatch:
    ExcelRowMatchesSnapshot = False
End Function

Private Function RemoteRowMatchesSnapshot(ByVal remoteWs As Worksheet, ByVal remoteRow As Long, ByVal apiHdr As Object, ByVal vals As Variant) As Boolean
    On Error GoTo NotMatch
    If RemoteDateISO(remoteWs.Cells(remoteRow, apiHdr("movement_date")).Value) <> Format$(CDate(vals(1)), "yyyy-mm-dd") Then GoTo NotMatch
    If CStr(remoteWs.Cells(remoteRow, apiHdr("concept")).Value) <> CStr(vals(2)) Then GoTo NotMatch
    If CStr(remoteWs.Cells(remoteRow, apiHdr("category")).Value) <> CStr(vals(3)) Then GoTo NotMatch
    If CStr(remoteWs.Cells(remoteRow, apiHdr("movement_type")).Value) <> CStr(vals(4)) Then GoTo NotMatch
    If Abs(Val(Replace(CStr(remoteWs.Cells(remoteRow, apiHdr("amount")).Value), ",", ".")) - CDbl(vals(5))) > 0.005 Then GoTo NotMatch
    If CStr(remoteWs.Cells(remoteRow, apiHdr("notes")).Value) <> CStr(vals(6)) Then GoTo NotMatch
    RemoteRowMatchesSnapshot = True
    Exit Function
NotMatch:
    RemoteRowMatchesSnapshot = False
End Function


Private Function PatchMovementFromExcel(ByVal ws As Worksheet, ByVal excelRow As Long, ByVal hdr As Object, _
    ByVal id As String, ByVal userId As String, ByVal token As String, ByRef status As Long, ByRef response As String) As Boolean
    Dim body As String, url As String
    body = "{""movement_date"":" & JsonString(Format$(CDate(ws.Cells(excelRow, hdr("data")).Value), "yyyy-mm-dd")) & _
        ",""concept"":" & JsonString(CStr(ws.Cells(excelRow, hdr("concetto")).Value)) & _
        ",""category"":" & JsonString(CStr(ws.Cells(excelRow, hdr("categoria")).Value)) & _
        ",""movement_type"":" & JsonString(CStr(ws.Cells(excelRow, hdr("tipo")).Value)) & _
        ",""amount"":" & Replace(Format$(CDbl(ws.Cells(excelRow, hdr("importo")).Value), "0.00"), ",", ".") & _
        ",""notes"":" & JsonString(OptionalCellText(ws, excelRow, hdr, "note")) & "}"
    url = SUPABASE_URL & "/rest/v1/" & TABLE_NAME & "?id=eq." & id & "&user_id=eq." & userId
    response = HttpRequest("PATCH", url, body, "application/json", token, status, "return=representation")
    If status < 200 Or status >= 300 Then
        PatchMovementFromExcel = False
    ElseIf InStr(1, response, id, vbTextCompare) = 0 Then
        status = 0
        PatchMovementFromExcel = False
    Else
        PatchMovementFromExcel = True
    End If
End Function

