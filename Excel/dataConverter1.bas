Attribute VB_Name = "Module3"
Sub ConvertDatesInActiveColumn()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim oldDate As String
    Dim newDate As Date
    Dim activeColumn As Long
    
    Set ws = ActiveSheet
    activeColumn = ActiveCell.Column
    
    Set rng = ws.Range(ws.Cells(1, activeColumn), ws.Cells(ws.Rows.Count, activeColumn).End(xlUp))
    
    For Each cell In rng
        If IsNumeric(cell.Value) And Len(cell.Value) = 8 Then
            oldDate = cell.Value
            newDate = DateSerial(Left(oldDate, 4), Mid(oldDate, 5, 2), Right(oldDate, 2))
            cell.NumberFormat = "dd.mm.yyyy"
            cell.Value = newDate
        End If
    Next cell
End Sub
