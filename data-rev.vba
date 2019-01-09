Attribute VB_Name = "Module1"
    Public LastRow As Integer
    Public curRow As Integer
    Public col_g(1000) As String
    Dim nb(1000) As Integer
    Dim pg(1000) As Integer
    Dim col_j(1000) As String
    Dim col_m(1000) As String
    Dim p1 As Integer
    Dim p2 As Integer
    Dim p_rev(1000) As String
    Dim drpos(1000) As Integer
    Dim drpunc(1000) As Integer
    Dim dr(1000) As String
    Dim rlpos(1000) As Integer
    Dim rl(1000) As String
    Dim tempstr As String
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
        nb(i) = mid(Cells(i, 7).value, p1 + 5, 5)
        pg(i) = mid(Cells(i, 7).value, p2 + 5, 2)
        Worksheets("Data").Cells(i, 9).value = nb(i) 'Column I; Notebook'
        Worksheets("Data").Cells(i, 10).value = pg(i) 'Column J; page'
        col_j(i) = Cells(i, 10).value
        col_m(i) = Cells(i, 13).value
        drpos(i) = InStr(col_m(i), "Data review") + 14
        If drpos(i) = 14 Then
            dr(i) = col_j(i)
        Else
            tempstr = mid(col_m(i), drpos(i), Len(col_m(i)))
            drpunc(i) = InStr(tempstr, "     ") + drpos(i)
            dr(i) = mid(col_m(i), drpos(i), (drpunc(i) - drpos(i)))
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
        rl(i) = mid(col_m(i), rlpos(i), Len(col_m(i)))
        Worksheets("Data").Cells(i, 8).value = rl(i)  'Column H; Released by'
    Next i
End Sub
Sub DR_Calculate()
    Dim dr_record(1000) As String
    Dim cur_row As Integer
    Worksheets("Data").Select
    Range(Cells(1, 7), Cells(LastRow, 7)).Copy _
    Destination:=Worksheets("Results").Range("A1")
    tempstr = "A" & LastRow + 1
    Range(Cells(2, 8), Cells(LastRow, 8)).Copy _
    Destination:=Worksheets("Results").Range(tempstr)
    Worksheets("Results").Select
    On Error Resume Next
      Range(1, 1).Select
      Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    Worksheets("Results").Cells(1, 1).Select
    cur_row = 1
    While activecell.value <> ""
      cur_row = cur_row + 1
      dr_record(i) = activecell.offset(1, 0).value
    Wend
End Sub
