Attribute VB_Name = "Module2"
' Seçili kolondaki boş hücreleri siler, sildiği hücrenin yerini alttaki hücreleri yukarı çekerek doldurur.
Sub DeleteBlankCellsInSelectedColumn()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim activeColumn As Long

    Set ws = ActiveSheet
    activeColumn = ActiveCell.Column

    lastRow = ws.Cells(ws.Rows.Count, activeColumn).End(xlUp).Row

    For i = lastRow To 1 Step -1
        If ws.Cells(i, activeColumn).Value = "" Then
            ws.Cells(i, activeColumn).Delete Shift:=xlUp
        End If
    Next i
End Sub

