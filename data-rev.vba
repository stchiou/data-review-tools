Attribute VB_Name = "DataReviewer"
    Public LastRow As Integer
    Public curRow As Integer
    Public col_g(1000) As String
    Public nb(1000) As Integer
    Public pg(1000) As Integer
    Public col_j(1000) As String
    Public col_m(1000) As String
    Public p1 As Integer
    Public p2 As Integer
    Public p_rev(1000) As String
    Public drpos(1000) As Integer
    Public drpunc(1000) As Integer
    Public dr(1000) As String
    Public rlpos(1000) As Integer
    Public rl(1000) As String
    Public tempstr As String
   
    Public cur_row As Integer
Sub DR_GenData()
    Dim i As Integer
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

    'Make a copy of Date to the target sheet'
    Worksheets("QA Data").Range(Cells(1, 5), Cells(LastRow, 5)).Copy _
    Destination:=Worksheets("Data").Range("A1")   'Column A; Date'
    'Make a copy of Method to the target sheet'
    Worksheets("QA Data").Range(Cells(1, 12), Cells(LastRow, 12)).Copy _
    Destination:=Worksheets("Data").Range("B1")   'Column B; Method'
    'Make a copy of Lot Numbers to the target sheet'
    Worksheets("QA Data").Range(Cells(1, 3), Cells(LastRow, 3)).Copy _
    Destination:=Worksheets("Data").Range("C1")   'Column C; Lot Number'
    Worksheets("QA Data").Range(Cells(1, 4), Cells(LastRow, 4)).Copy _
    Destination:=Worksheets("Data").Range("D1")   'Column D; List Number'
    'Make a copy of Error type to the target sheet'
    Worksheets("QA Data").Range(Cells(1, 6), Cells(LastRow, 6)).Copy _
    Destination:=Worksheets("Data").Range("E1")   'Column E; Error Type'
    'Make a copy of Error Class to the target sheet'
    Worksheets("QA Data").Range(Cells(1, 8), Cells(LastRow, 8)).Copy _
    Destination:=Worksheets("Data").Range("F1")   'Column F; Error class'
    'Create Headers for Notebook and Page Number columns'
    Worksheets("Data").Cells(1, 7).value = "Data Reviewer" 'Column G; Data Reviewer'
    Worksheets("Data").Cells(1, 8).value = "Released by"   'Column H; Released by'
    Worksheets("Data").Cells(1, 9).value = "Note Book"     'Column I; Notebook'
    Worksheets("Data").Cells(1, 10).value = "Page"          'Column J; Page'
    'Parse Notebook and Page Number from the source sheet to the target sheet'
     For i = 2 To LastRow
        Cells(i, 7).Select
        col_g(i) = Cells(i, 7).value
        p1 = InStr(col_g(i), "Book ")
        p2 = InStr(col_g(i), "page ")
        nb(i) = Mid(Cells(i, 7).value, p1 + 5, 5)
        pg(i) = Mid(Cells(i, 7).value, p2 + 5, 2)
        Worksheets("Data").Cells(i, 9).value = nb(i) 'Column I; Notebook'
        Worksheets("Data").Cells(i, 10).value = pg(i) 'Column J; page'
        col_j(i) = Cells(i, 10).value
        col_m(i) = Cells(i, 13).value
        drpos(i) = InStr(col_m(i), "Data review") + 14
        If drpos(i) = 14 Then
            dr(i) = col_j(i)
        Else
            tempstr = Mid(col_m(i), drpos(i), Len(col_m(i)))
            drpunc(i) = InStr(tempstr, "     ") + drpos(i)
            dr(i) = Mid(col_m(i), drpos(i), (drpunc(i) - drpos(i)))
        End If
        If InStr(dr(i), "N/A") > 0 Then
          dr(i) = ""
        Else
        End If
        If InStr(dr(i), "?") > 0 Then
            dr(i) = ""
        Else
        End If
        Worksheets("Data").Cells(i, 7).value = dr(i)  'Column G; Data Reviewer'
        rlpos(i) = InStr(col_m(i), "Released by ") + 12
        rl(i) = Mid(col_m(i), rlpos(i), Len(col_m(i)))
        Worksheets("Data").Cells(i, 8).value = rl(i)  'Column H; Released by'
    Next i

    'copy reviewer's name, error class, and error type to result sheet'
    Worksheets("Data").Select
    Range(Cells(1, 7), Cells(LastRow, 7)).Copy _
    Destination:=Worksheets("Results").Range("A1")
    tempstr = "A" & LastRow + 1
    Range(Cells(2, 8), Cells(LastRow, 8)).Copy _
    Destination:=Worksheets("Results").Range(tempstr)
    Range(Cells(1, 6), Cells(LastRow, 6)).Copy _
    Destination:=Worksheets("Results").Range("B1")
    tempstr = "B" & LastRow + 1
    Range(Cells(2, 6), Cells(LastRow, 6)).Copy _
    Destination:=Worksheets("Results").Range(tempstr)
    Range(Cells(1, 5), Cells(LastRow, 5)).Copy _
    Destination:=Worksheets("Results").Range("C1")
    tempstr = "C" & LastRow + 1
    Range(Cells(2, 5), Cells(LastRow, 5)).Copy _
    Destination:=Worksheets("Results").Range(tempstr)
    'Worksheets("Results").Select
    'Range(Cells(1, 1), Cells(LastRow * 2, 1)).Copy _
    'Destination:=Worksheets("Results").Range("E1")
    'Remove duplicates
    'Range(Cells(1, 5), Cells(LastRow * 2, 5)).RemoveDuplicates Columns:=1, Header:=xlYes
    'Remove blank cell on data review name column'
    Worksheets("Results").Activate
    On Error Resume Next
        Range(1, 1).Select
        Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete 'Shift:=xlShiftUp
    Dim name_count As Integer
    Dim res_name_count As Integer
    Dim res_pos As Integer
    Dim nam_pos As Integer
    Dim dr_name_rs As String
    Dim restr As String
    'Detect row number'
    res_name_count = Worksheets("Results").Cells(1, 2).End(xlDown).Row
    name_count = Worksheets("Name").Cells(1, 1).End(xlDown).Row
    Cells(2, 1).Select
    For res_pos = 2 To res_name_count
        
    Next res_pos
End Sub
