Attribute VB_Name = "AggiornaMovimentiDaApp"
Option Explicit

' Importa il CSV esportato da Bilancio Casa nel foglio Movimenti.
' Aggiorna le righe esistenti usando ID movimento e aggiunge solo gli ID nuovi.
' Non elimina righe Excel assenti dal CSV e ignora righe senza ID (es. saldo iniziale).
Public Sub AggiornaMovimentiDaApp()
    Dim target As Worksheet, srcWb As Workbook, src As Worksheet
    Dim filePath As Variant, lastSrc As Long, lastTgt As Long
    Dim srcHeaders As Object, tgtHeaders As Object, idRows As Object
    Dim c As Long, r As Long, destRow As Long, idColSrc As Long, idColTgt As Long
    Dim header As String, movementId As String, updated As Long, added As Long
    Dim key As Variant, targetRow As Long, oldCalc As XlCalculation

    On Error GoTo Fail
    Set target = ThisWorkbook.Worksheets("Movimenti")
    filePath = Application.GetOpenFilename("File CSV (*.csv),*.csv", , "Seleziona il CSV esportato da Bilancio Casa")
    If VarType(filePath) = vbBoolean Then Exit Sub

    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Workbooks.OpenText Filename:=CStr(filePath), Origin:=65001, DataType:=xlDelimited, _
        TextQualifier:=xlTextQualifierDoubleQuote, Semicolon:=True, Comma:=False, Tab:=False, _
        Local:=True
    Set srcWb = ActiveWorkbook
    Set src = srcWb.Worksheets(1)

    Set srcHeaders = CreateObject("Scripting.Dictionary")
    Set tgtHeaders = CreateObject("Scripting.Dictionary")
    Set idRows = CreateObject("Scripting.Dictionary")
    srcHeaders.CompareMode = vbTextCompare
    tgtHeaders.CompareMode = vbTextCompare
    idRows.CompareMode = vbTextCompare

    For c = 1 To src.Cells(1, src.Columns.Count).End(xlToLeft).Column
        header = NormalizzaIntestazione(CStr(src.Cells(1, c).Value))
        If Len(header) > 0 Then srcHeaders(header) = c
    Next c
    For c = 1 To target.Cells(1, target.Columns.Count).End(xlToLeft).Column
        header = NormalizzaIntestazione(CStr(target.Cells(1, c).Value))
        If Len(header) > 0 Then tgtHeaders(header) = c
    Next c

    If Not srcHeaders.Exists("idmovimento") Or Not tgtHeaders.Exists("idmovimento") Then
        Err.Raise vbObjectError + 201, , "Manca l'intestazione 'ID movimento' nel CSV o nel foglio Movimenti."
    End If
    If Not srcHeaders.Exists("data") Or Not srcHeaders.Exists("concetto") Or _
       Not srcHeaders.Exists("tipo") Or Not srcHeaders.Exists("importo") Then
        Err.Raise vbObjectError + 202, , "Il CSV non contiene tutte le colonne obbligatorie."
    End If

    idColSrc = srcHeaders("idmovimento")
    idColTgt = tgtHeaders("idmovimento")
    lastSrc = src.Cells(src.Rows.Count, idColSrc).End(xlUp).Row
    lastTgt = target.Cells(target.Rows.Count, idColTgt).End(xlUp).Row

    For r = 2 To lastTgt
        movementId = Trim$(CStr(target.Cells(r, idColTgt).Value))
        If Len(movementId) > 0 Then
            If Not idRows.Exists(movementId) Then idRows.Add movementId, r
        End If
    Next r

    For r = 2 To lastSrc
        movementId = Trim$(CStr(src.Cells(r, idColSrc).Value))
        If Len(movementId) > 0 Then
            If idRows.Exists(movementId) Then
                destRow = CLng(idRows(movementId))
                updated = updated + 1
            Else
                lastTgt = target.Cells(target.Rows.Count, idColTgt).End(xlUp).Row
                destRow = lastTgt + 1
                idRows.Add movementId, destRow
                added = added + 1
            End If

            ' Copia solo colonne presenti in entrambi i fogli.
            ' La colonna Mese del CSV viene ignorata se non esiste nel foglio Excel.
            For Each key In srcHeaders.Keys
                If tgtHeaders.Exists(CStr(key)) Then
                    target.Cells(destRow, tgtHeaders(CStr(key))).Value = src.Cells(r, srcHeaders(CStr(key))).Value
                End If
            Next key
        End If
    Next r

    srcWb.Close SaveChanges:=False
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Sincronizzazione completata." & vbCrLf & _
           "Movimenti aggiornati: " & updated & vbCrLf & _
           "Nuovi movimenti aggiunti: " & added & vbCrLf & _
           "Le righe non presenti nel CSV non sono state eliminate.", vbInformation
    Exit Sub

Fail:
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Sincronizzazione interrotta: " & Err.Description, vbExclamation
End Sub

Private Function NormalizzaIntestazione(ByVal s As String) As String
    s = LCase$(Trim$(s))
    s = Replace(s, ChrW(65279), "")
    s = Replace(s, " ", "")
    s = Replace(s, "+", "piu")
    s = Replace(s, "-", "")
    s = Replace(s, "/", "")
    NormalizzaIntestazione = s
End Function
