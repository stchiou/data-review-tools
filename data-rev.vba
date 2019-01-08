Attribute VB_Name = "Module1"
Sub Data_Review()
    Dim LastRow As Integer
    Dim curRow As Integer
    Dim i As Integer
    Dim j As Integer
    Dim col_g(100) As String
    Dim nb(100) As Integer
    Dim pg(100) As Integer
    Dim col_j(100) As String
    Dim col_m(100) As String
    Dim p1 As Integer
    Dim p2 As Integer
    Dim p_rev as String
    Dim rev as String
    Dim rel as String
    Dim comment as String
    Dim drpos as Integer
    Dim drpunc as Integer
    Dim dr as String
    Dim rlpos as Integer
    Dim tempstr as string
    'Create a new sheet for consolidated data'
    Sheets.Add after:=Sheets("supplement")
    Sheets(Sheets.Count).Select
    Sheets(Sheets.Count).Name = "Data"
    'Remove extra blank lines on the spreadsheet'
    On Error Resume Next
        Range(1, 1).Select
        Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    'Determine number of record in the sheet'
    Worksheets("QA Data").Select
    LastRow = Cells(1, 1).End(xlDown).Row
    'Make a copy of Date and Method from the source sheet to the target sheet'
    Worksheets("QA Data").Range(Cells(1, 5), Cells(LastRow, 5)).Copy _
    Destination:=Worksheets("Data").Range("A1")
    Worksheets("QA Data").Range(Cells(1, 12), Cells(LastRow, 12)).Copy _
    Destination:=Worksheets("Data").Range("B1")
    Worksheets("QA Data").Range(Cells(1, 6), Cells(LastRow, 6)).Copy _
    Destination:=Worksheets("Data").Range("E1")
    Worksheets("QA Data").Range(Cells(1, 8), Cells(LastRow, 8)).Copy _
    Destination:=Worksheets("Data").Range("F1")
    'Create Headers for Notebook and Page Number columns'
    Worksheets("Data").Cells(1, 3).Value = "Note Book"
    Worksheets("Data").Cells(1, 4).Value = "Page"
    Worksheets("Data").Cells(1, 5).Value = "Data Reviewer"
    Worksheets("Data").Cells(1, 6).Value = "Released by"
    'Parse Notebook and Page Number from the source sheet to the target sheet'
     For i = 2 To LastRow
        Cells(i, 7).Select
        col_g(i) = Cells(i, 7).Value
        p1 = InStr(col_g(i), "Book ")
        p2 = InStr(col_g(i), "page ")
        nb(i) = Mid(Cells(i, 7).Value, p1 + 5, 5)
        pg(i) = Mid(Cells(i, 7).Value, p2 + 5, 2)
        Worksheets("Data").Cells(i, 3).Value = nb(i)
        Worksheets("Data").Cells(i, 4).Value = pg(i)
        col_j(i) = cells(i, 10).Value
        cells(i,4).value=comment
        drpos=Instr(comment,"Data reviewer ")+14
        tempstr=mid(comment,drpos,len(comment))
        drpunc=Instr(tempstr,"     ")+drpos
        dr=mid(comment,drpos,(drpunc-drpos))
        Worksheets("Data").Cells(i,5).Value=dr
        rlpos=Instr(comment,"Released by ")+12
        rl=mid(comment,rlpos,len(comment))
        Worksheets("Data").Cells(i,6).Value=rl
    Next i
End Sub
