Attribute VB_Name = "Module1"
Sub Data_Review()
    Dim LastRow As Integer
    Dim curRow As Integer
    Dim i As Integer
    Dim j As Integer
    Dim temp(100)as string
    Dim nb(100) As integer
    Dim pg(100) As integer
    Dim p1 as Integer
    Dim p2 as Integer
    Dim reviewer(30, 3) as string
    'Create a new sheet for consolidated data'
    Sheets.Add after:=Sheets("QA Data")
    Sheets(Sheets.Count).Select
    Sheets(Sheets.Count).Name = "Data"
    'Remove extra blank lines on the spreadsheet'
    On Error Resume Next
        Range(1, 1).Select
        Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    'Determine number of record in the sheet'
    Worksheets("QA Data").Select
    LastRow = Cells(1, 1).End(xlDown).Row
    For i=2 to 30
      For j=4 to 6
        reviewer(i,j)=Worksheets("supplement").Range(Cells(i,j)).Value
      Next j
    Next i
    'Make a copy of Date and Method from the source sheet to the target sheet'
    Worksheets("QA Data").Range(Cells(1, 5), Cells(LastRow, 5)).Copy _
    Destination:=Worksheets("Data").Range("A1")
    Worksheets("QA Data").Range(Cells(1, 12), Cells(LastRow, 12)).Copy _
    Destination:=Worksheets("Data").Range("B1")
    Worksheets("QA Data").Range(Cells(1,6), Cells(LastRow,6)).Copy _
    Destination:=Worksheets("Data").Range("E1")
    Worksheets("QA Data").Range(Cells(1,8), Cells(LastRow,6)).Copy _
    Destination:=Worksheets("Data").Range("F1")
    'Create Headers for Notebook and Page Number columns'
    Worksheets("Data").Cells(1, 3).Value = "Note Book"
    Worksheets("Data").Cells(1, 4).Value = "Page"
    'Parse Notebook and Page Number from the source sheet to the target sheet'
    For i = 2 To LastRow
        Cells(i, 7).Select
        temp(i)=Cells(i,7).Value
        p1=Instr(temp(i),"Book ")
        p2=Instr(temp(i),"page ")
        nb(i) = Mid(Cells(i, 7).Value, p1 + 5, 5)
        pg(i) = Mid(Cells(i, 7).Value, p2 + 5, 2)
        Worksheets("Data").Cells(i, 3).Value = nb(i)
        Worksheets("Data").Cells(i, 4).Value = pg(i)
    Next i







End Sub
