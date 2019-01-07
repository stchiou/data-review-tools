Sub Data_Review()
    Dim LastRow As Integer
    Dim curRow As Integer
    Dim i As Integer
    Sheets.Add after:=Sheets("QA Data")
    Sheets(Sheets.Count).Select
    Sheets(Sheets.Count).Name = "Data"
    On Error Resume Next
        Range(1, 1).Select
        Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    Worksheets("QA Data").Select
    LastRow = Cells(1, 1).End(xlDown).Row
    Dim nb(LastRow) As String
    Dim pg(LastRow) As String
    Worksheets("QA Data").Range(Cells(1, 5), Cells(LastRow, 5)).Copy _
    Destination:=Worksheets("Data").Range("A1")
    Worksheets("QA Data").Range(Cells(1, 12), Cells(LastRow, 12)).Copy _
    Destination:=Worksheets("Data").Range("B1")
    Worksheets("Data").Cells(1, 3).Value = "Note Book"
    Worksheets("Data").Cells(1, 4).Value = "Page"
    For i = 2 To LastRow
        Cells(i, 7).Select
        nb(i) = Mid(Cells(i, 7).Value, Find("Book ", Cells(i, 7).Value) + 5, 5)
        pg(i) = Mid(Cells(i, 7).Value, Find("page ", Cells(i, 7).Value) + 5, 2)
        Worksheets("Data").Cells(i, 3).Value = nb(i)
        Worksheets("Data").Cells(i, 4).Value = pg(i)
    Next i







End Sub
