Attribute VB_Name = "CreaFoglioContanti"
Option Explicit

' Ricrea da zero il foglio Contanti senza modificare Movimenti o le macro di sincronizzazione.
' Richiede la tabella Excel MovimentiTbl con le colonne Data, Categoria, Tipo, Importo e Saldo.
' I prelievi sono identificati dalla categoria "Prelievo contanti" e dal tipo "Uscita".

Public Sub CreaFoglioContanti()
    Dim wb As Workbook
    Dim wsMov As Worksheet, ws As Worksheet
    Dim lo As ListObject
    Dim oldSheet As Worksheet
    Dim firstYear As Long, lastYear As Long, y As Long, m As Long
    Dim dateCol As Long, c As Range, hasDate As Boolean
    Dim answer As VbMsgBoxResult
    Dim formulaBank As String, formulaWithdrawals As String
    Dim monthlyRow As Long, annualRow As Long, i As Long

    On Error GoTo Fail
    Set wb = ThisWorkbook

    On Error Resume Next
    Set wsMov = Nothing
    Set lo = Nothing
    Set wsMov = wb.Worksheets("Movimenti")
    If Not wsMov Is Nothing Then Set lo = wsMov.ListObjects("MovimentiTbl")
    On Error GoTo Fail

    If wsMov Is Nothing Or lo Is Nothing Then
        MsgBox "Non trovo la tabella MovimentiTbl nel foglio Movimenti. Non ho modificato nulla.", vbExclamation
        Exit Sub
    End If

    If Not HasTableColumn(lo, "Data") Or Not HasTableColumn(lo, "Categoria") Or _
       Not HasTableColumn(lo, "Tipo") Or Not HasTableColumn(lo, "Importo") Or _
       Not HasTableColumn(lo, "Saldo") Then
        MsgBox "MovimentiTbl deve contenere le colonne Data, Categoria, Tipo, Importo e Saldo. Non ho modificato nulla.", vbExclamation
        Exit Sub
    End If

    answer = MsgBox("Il foglio Contanti verrà ricreato da zero. Se esiste già, il suo contenuto sarà sostituito. Continuare?", _
                    vbYesNo + vbQuestion + vbDefaultButton2, "Ricrea foglio Contanti")
    If answer <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets("Contanti").Delete
    wb.Worksheets("Contanti e pareggio").Delete
    On Error GoTo Fail
    Application.DisplayAlerts = True

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "Contanti"

    firstYear = Year(Date)
    lastYear = Year(Date)
    If Not lo.DataBodyRange Is Nothing Then
        dateCol = lo.ListColumns("Data").Index
        For Each c In lo.ListColumns(dateCol).DataBodyRange.Cells
            If IsDate(c.Value) Then
                y = Year(CDate(c.Value))
                If Not hasDate Then
                    firstYear = y
                    hasDate = True
                ElseIf y < firstYear Then
                    firstYear = y
                End If
                If y > lastYear Then lastYear = y
            End If
        Next c
    End If
    If firstYear > Year(Date) Then firstYear = Year(Date)
    If lastYear < Year(Date) Then lastYear = Year(Date)

    With ws
        .Range("A1").Value = "CONTANTI"
        .Range("A1:F1").Merge
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 16

        .Range("A3").Value = "Contanti disponibili (€)"
        .Range("B3").Value = 50
        .Range("B3").Interior.Color = RGB(255, 242, 204)
        .Range("B3").NumberFormat = "#,##0.00 €"
        .Range("A4").Value = "Totale soldi disponibili (€)"
        formulaBank = "=B3+SE.ERRORE(INDICE(MovimentiTbl[Saldo];RIGHE(MovimentiTbl[Saldo]));0)"
        .Range("B4").FormulaLocal = formulaBank
        .Range("B4").NumberFormat = "#,##0.00 €"
        .Range("A4:B4").Font.Bold = True
        .Range("A4:B4").Font.Size = 14

        .Range("A6").Value = "Controllo contanti"
        .Range("A6").Font.Bold = True
        .Range("A7").Value = "Saldo iniziale contanti (€)"
        .Range("B7").Value = 50
        .Range("A8").Value = "Prelievi complessivi registrati (€)"
        .Range("B8").FormulaLocal = "=SOMMA.PIÙ.SE(MovimentiTbl[Importo];MovimentiTbl[Categoria];""Prelievo contanti"";MovimentiTbl[Tipo];""Uscita"")"
        .Range("A9").Value = "Contanti teorici prima delle spese (€)"
        .Range("B9").FormulaLocal = "=B7+B8"
        .Range("A10").Value = "Spese in contanti stimate/non registrate (€)"
        .Range("B10").FormulaLocal = "=B9-B3"
        .Range("A11").Value = "Controllo"
        .Range("B11").FormulaLocal = "=SE(B10< -0,01;""VERIFICARE: contanti attuali superiori al teorico"";""Coerente: differenza attribuita a spese in contanti non registrate"")"
        .Range("B7:B10").NumberFormat = "#,##0.00 €"
        .Range("B11").WrapText = True

        .Range("A14").Value = "RIEPILOGO MENSILE PRELIEVI"
        .Range("A14:C14").Merge
        .Range("A14").Font.Bold = True
        .Range("A15").Value = "Anno"
        .Range("B15").Value = Year(Date)
        .Range("B15").Interior.Color = RGB(255, 242, 204)
        .Range("A16").Value = "Mese"
        .Range("B16").Value = "Dal"
        .Range("C16").Value = "Prelievi (€)"
        .Range("A16:C16").Font.Bold = True

        For m = 1 To 12
            monthlyRow = 16 + m
            .Cells(monthlyRow, 1).FormulaLocal = "=TESTO(DATA($B$15;" & m & ";1);""mmmm"")"
            .Cells(monthlyRow, 2).FormulaLocal = "=DATA($B$15;" & m & ";1)"
            .Cells(monthlyRow, 2).NumberFormat = "dd/mm/yyyy"
            .Cells(monthlyRow, 3).FormulaLocal = "=SOMMA.PIÙ.SE(MovimentiTbl[Importo];MovimentiTbl[Categoria];""Prelievo contanti"";MovimentiTbl[Tipo];""Uscita"";MovimentiTbl[Data];"">=""&B" & monthlyRow & ";MovimentiTbl[Data];""<""&DATA($B$15;" & (m + 1) & ";1))"
            .Cells(monthlyRow, 3).NumberFormat = "#,##0.00 €"
        Next m
        .Range("A29").Value = "Totale anno selezionato"
        .Range("C29").FormulaLocal = "=SOMMA(C17:C28)"
        .Range("A29:C29").Font.Bold = True
        .Range("C29").NumberFormat = "#,##0.00 €"

        .Range("E14").Value = "RIEPILOGO ANNUALE PRELIEVI"
        .Range("E14:F14").Merge
        .Range("E14").Font.Bold = True
        .Range("E16").Value = "Anno"
        .Range("F16").Value = "Prelievi (€)"
        .Range("E16:F16").Font.Bold = True
        annualRow = 17
        For i = firstYear To lastYear
            .Cells(annualRow, 5).Value = i
            .Cells(annualRow, 6).FormulaLocal = "=SOMMA.PIÙ.SE(MovimentiTbl[Importo];MovimentiTbl[Categoria];""Prelievo contanti"";MovimentiTbl[Tipo];""Uscita"";MovimentiTbl[Data];"">=""&DATA(E" & annualRow & ";1;1);MovimentiTbl[Data];""<""&DATA(E" & annualRow + 1 & ";1;1))"
            .Cells(annualRow, 6).NumberFormat = "#,##0.00 €"
            annualRow = annualRow + 1
        Next i

        .Columns("A").ColumnWidth = 39
        .Columns("B").ColumnWidth = 19
        .Columns("C").ColumnWidth = 18
        .Columns("D").ColumnWidth = 3
        .Columns("E").ColumnWidth = 13
        .Columns("F").ColumnWidth = 18
        .Rows(11).RowHeight = 34
        .Range("A1:F1").Interior.Color = RGB(217, 225, 242)
        .Range("A14:C14").Interior.Color = RGB(217, 225, 242)
        .Range("E14:F14").Interior.Color = RGB(217, 225, 242)
        .Range("A16:C16").Interior.Color = RGB(242, 242, 242)
        .Range("E16:F16").Interior.Color = RGB(242, 242, 242)
        .Range("B3:B4").Borders.LineStyle = xlContinuous
        .Range("A7:B11").Borders.LineStyle = xlContinuous
        .Range("A16:C29").Borders.LineStyle = xlContinuous
        .Range("E16:F" & Application.Max(16, annualRow - 1)).Borders.LineStyle = xlContinuous
        .Range("A1:F29").VerticalAlignment = xlVAlignCenter
        .Activate
        .Range("B3").Select
    End With

    Application.ScreenUpdating = True
    MsgBox "Foglio Contanti creato. Aggiorna manualmente solo la cella gialla B3 con il contante effettivo; i prelievi e i riepiloghi si aggiornano dalla tabella MovimentiTbl.", vbInformation
    Exit Sub

Fail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Errore durante la creazione del foglio Contanti: " & Err.Description, vbExclamation
End Sub

Private Function HasTableColumn(ByVal lo As ListObject, ByVal columnName As String) As Boolean
    Dim lc As ListColumn
    On Error Resume Next
    Set lc = lo.ListColumns(columnName)
    HasTableColumn = Not lc Is Nothing
    On Error GoTo 0
End Function
