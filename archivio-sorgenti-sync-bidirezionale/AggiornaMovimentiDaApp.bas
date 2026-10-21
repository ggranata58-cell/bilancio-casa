Attribute VB_Name = "AggiornaMovimentiDaApp"
Option Explicit

' Importa il CSV esportato da Bilancio Casa nel foglio Movimenti.
' Cerca le intestazioni nel foglio Excel (anche alla riga 4).
' Aggiorna per ID stabile, aggiunge nuovi ID e non elimina righe.
Public Sub AggiornaMovimentiDaApp()
    Dim target As Worksheet, srcWb As Workbook, src As Worksheet
    Dim filePath As Variant, lastSrc As Long, lastTgt As Long
    Dim srcHeaders As Object, tgtHeaders As Object, idRows As Object
    Dim c As Long, r As Long, destRow As Long, idColSrc As Long, idColTgt As Long
    Dim srcHeaderRow As Long, tgtHeaderRow As Long
    Dim header As String, movementId As String, updated As Long, added As Long
    Dim key As Variant, oldCalc As XlCalculation
    Dim lo As ListObject, targetTable As ListObject, newListRow As ListRow
    Dim numeroErrore As Long, descrizioneErrore As String
    Dim csvText As String, tempPath As String, stm As Object

    On Error GoTo Fail
    Set target = ThisWorkbook.Worksheets("Movimenti")
    filePath = Application.GetOpenFilename("File CSV (*.csv),*.csv", , "Seleziona il CSV esportato da Bilancio Casa")
    If VarType(filePath) = vbBoolean Then Exit Sub

    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Excel può scambiare un CSV che inizia con "ID" per un file SYLK.
    ' Crea una copia temporanea UTF-8 cambiando solo "ID" in "Id" nella prima intestazione.
    ' Il CSV originale esportato dall'app resta intatto.
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile CStr(filePath)
    csvText = stm.ReadText
    stm.Close

    If Len(csvText) > 0 Then
        If Left$(csvText, 1) = ChrW(65279) Then csvText = Mid$(csvText, 2)
        If UCase$(Left$(csvText, 2)) = "ID" Then csvText = "Id" & Mid$(csvText, 3)
    End If

    tempPath = Environ$("TEMP") & "\BilancioCasa_import_temp.csv"
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText csvText
    stm.SaveToFile tempPath, 2
    stm.Close
    Set stm = Nothing

    Workbooks.OpenText Filename:=tempPath, Origin:=65001, DataType:=xlDelimited, _
        TextQualifier:=xlTextQualifierDoubleQuote, Semicolon:=True, Comma:=False, Tab:=False, _
        Local:=True
    Set srcWb = ActiveWorkbook
    Set src = srcWb.Worksheets(1)
    srcHeaderRow = 1

    Set srcHeaders = CreateObject("Scripting.Dictionary")
    Set tgtHeaders = CreateObject("Scripting.Dictionary")
    Set idRows = CreateObject("Scripting.Dictionary")
    srcHeaders.CompareMode = vbTextCompare
    tgtHeaders.CompareMode = vbTextCompare
    idRows.CompareMode = vbTextCompare

    For c = 1 To src.Cells(srcHeaderRow, src.Columns.Count).End(xlToLeft).Column
        header = NormalizzaIntestazione(CStr(src.Cells(srcHeaderRow, c).Value))
        If Len(header) > 0 Then srcHeaders(header) = c
    Next c

    ' Prima cerca la tabella Excel: le intestazioni possono essere alla riga 4.
    tgtHeaderRow = 0
    For Each lo In target.ListObjects
        For c = 1 To lo.ListColumns.Count
            header = NormalizzaIntestazione(CStr(lo.HeaderRowRange.Cells(1, c).Value))
            If header = "idmovimento" Then
                Set targetTable = lo
                tgtHeaderRow = lo.HeaderRowRange.Row
                Exit For
            End If
        Next c
        If tgtHeaderRow > 0 Then Exit For
    Next lo

    ' Se il foglio non usa una tabella Excel, cerca l'intestazione nelle prime 20 righe.
    If tgtHeaderRow = 0 Then
        For r = 1 To 20
            For c = 1 To target.Cells(r, target.Columns.Count).End(xlToLeft).Column
                If NormalizzaIntestazione(CStr(target.Cells(r, c).Value)) = "idmovimento" Then
                    tgtHeaderRow = r
                    Exit For
                End If
            Next c
            If tgtHeaderRow > 0 Then Exit For
        Next r
    End If

    If tgtHeaderRow = 0 Then
        Err.Raise vbObjectError + 201, , "Non trovo l'intestazione ID movimento nel foglio Movimenti nelle prime 20 righe."
    End If

    For c = 1 To target.Cells(tgtHeaderRow, target.Columns.Count).End(xlToLeft).Column
        header = NormalizzaIntestazione(CStr(target.Cells(tgtHeaderRow, c).Value))
        If Len(header) > 0 Then tgtHeaders(header) = c
    Next c

    If Not srcHeaders.Exists("idmovimento") Then
        Err.Raise vbObjectError + 202, , "ID movimento non riconosciuto nel CSV. Controllare la prima riga delle intestazioni."
    End If
    If Not tgtHeaders.Exists("idmovimento") Then
        Err.Raise vbObjectError + 203, , "ID movimento non riconosciuto nel foglio Excel."
    End If
    If Not srcHeaders.Exists("data") Or Not srcHeaders.Exists("concetto") Or _
       Not srcHeaders.Exists("tipo") Or Not srcHeaders.Exists("importo") Then
        Err.Raise vbObjectError + 204, , "Il CSV non contiene tutte le colonne obbligatorie."
    End If

    idColSrc = srcHeaders("idmovimento")
    idColTgt = tgtHeaders("idmovimento")
    lastSrc = src.Cells(src.Rows.Count, idColSrc).End(xlUp).Row
    If Not targetTable Is Nothing Then
        If targetTable.DataBodyRange Is Nothing Then
            lastTgt = tgtHeaderRow
        Else
            lastTgt = targetTable.DataBodyRange.Row + targetTable.DataBodyRange.Rows.Count - 1
        End If
    Else
        lastTgt = target.Cells(target.Rows.Count, idColTgt).End(xlUp).Row
        If lastTgt < tgtHeaderRow Then lastTgt = tgtHeaderRow
    End If

    For r = tgtHeaderRow + 1 To lastTgt
        movementId = Trim$(CStr(target.Cells(r, idColTgt).Value))
        If Len(movementId) > 0 Then
            If Not idRows.Exists(movementId) Then idRows.Add movementId, r
        End If
    Next r

    For r = srcHeaderRow + 1 To lastSrc
        movementId = Trim$(CStr(src.Cells(r, idColSrc).Value))
        If Len(movementId) > 0 Then
            If idRows.Exists(movementId) Then
                destRow = CLng(idRows(movementId))
                updated = updated + 1
            Else
                If Not targetTable Is Nothing Then
                    Set newListRow = targetTable.ListRows.Add
                    destRow = newListRow.Range.Row
                Else
                    lastTgt = target.Cells(target.Rows.Count, idColTgt).End(xlUp).Row
                    If lastTgt < tgtHeaderRow Then lastTgt = tgtHeaderRow
                    destRow = lastTgt + 1
                End If
                idRows.Add movementId, destRow
                added = added + 1
            End If

            For Each key In srcHeaders.Keys
                If tgtHeaders.Exists(CStr(key)) Then
                    target.Cells(destRow, tgtHeaders(CStr(key))).Value = src.Cells(r, srcHeaders(CStr(key))).Value
                End If
            Next key
        End If
    Next r

    srcWb.Close SaveChanges:=False
    If Len(tempPath) > 0 Then Kill tempPath
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Sincronizzazione completata." & vbCrLf & _
           "Movimenti aggiornati: " & updated & vbCrLf & _
           "Nuovi movimenti aggiunti: " & added & vbCrLf & _
           "Le righe non presenti nel CSV non sono state eliminate.", vbInformation
    Exit Sub

Fail:
    numeroErrore = Err.Number
    descrizioneErrore = Err.Description
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    If Len(tempPath) > 0 Then
        If Len(Dir$(tempPath)) > 0 Then Kill tempPath
    End If
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    MsgBox "Errore " & numeroErrore & ": " & descrizioneErrore, vbExclamation
End Sub

Private Function NormalizzaIntestazione(ByVal s As String) As String
    s = Replace(s, ChrW(65279), "")
    s = Replace(s, ChrW(160), " ")
    s = LCase$(Trim$(s))
    s = Replace(s, " ", "")
    s = Replace(s, "+", "piu")
    s = Replace(s, "-", "")
    s = Replace(s, "/", "")
    NormalizzaIntestazione = s
End Function
