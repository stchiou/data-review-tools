Attribute VB_Name = "DataReviewer"
Sub DR_GenData()
'--------------------------------------------------------
'Script for processing and summarize Data Reviewer Error
'v1.0
'Sean Chiou, Jan 7, 2019
'--------------------------------------------------------
    Dim LastRow As Integer               'Last row on the spreadsheet
    Dim curRow As Integer                'Current row of the spreadsheet
    Dim col_g(1000) As String            'Value of Column G of the Sheet "QA Data" (Error Description)
    Dim nb(1000) As Integer              'Value of the Notebook number parsed from col_g()
    Dim pg(1000) As Integer              'Value of the Page number parsed from col_g()
    Dim col_j(1000) As String            'Value of Column J of the Sheet "QA Data" (Previous Reviewer)
    Dim col_m(1000) As String            'Value of Column M of the Sheet "QA Data" (Comments)
    Dim p1 As Integer                    'Temporary variable to store the position of the pattern "Book" while processing Column G of Sheet "QA Data"
    Dim p2 As Integer                    'Temporary variable to store the position of the pattern "Page" while processing Column G of Sheet "QA Data"
    Dim drpos(1000) As Integer           'Position of "Data Review" in Column M of the Sheet "QA Data"
    Dim drpunc(1000) As Integer          'Position of "  " (two consecutive spaces) in Column M of the Sheet "QA Data"
    Dim dr(1000) As String               'Value of names after "Data Reviewer", parsed from Column M of the Sheet "QA Data"
    Dim rlpos(1000) As Integer           'Position of "Released by" in Column M of the Sheet "QA Data"
    Dim rl(1000) As String               'Value of names after "Released by", parsed from Column M of the Sheet "QA Data"
    Dim tempstr As String                'Temporary string holder while processing data
    Dim i As Integer
'-----------------------------------------------------------------------------------------------------------------------------
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
'Make a copy of Column E(Date) from Sheet "QA Data" to Column A of Sheet "Data"
    Worksheets("QA Data").Range(Cells(1, 5), Cells(LastRow, 5)).Copy _
    Destination:=Worksheets("Data").Range("A1")
'Make a copy of Column L (Method) from Sheet "QA Data" to Column B of Sheet "Data"
    Worksheets("QA Data").Range(Cells(1, 12), Cells(LastRow, 12)).Copy _
    Destination:=Worksheets("Data").Range("B1")   'Column B; Method'
'Make a copy of Lot Numbers from Sheet "QA Data" to Sheet "Data"
    Worksheets("QA Data").Range(Cells(1, 3), Cells(LastRow, 3)).Copy _
    Destination:=Worksheets("Data").Range("C1")   'Column C; Lot Number'
    Worksheets("QA Data").Range(Cells(1, 4), Cells(LastRow, 4)).Copy _
    Destination:=Worksheets("Data").Range("D1")   'Column D; List Number'
'Make a copy of Error type to the target sheet'
    Worksheets("QA Data").Range(Cells(1, 6), Cells(LastRow, 6)).Copy _
    Destination:=Worksheets("Data").Range("E1")   'Column E; Error Type'
'Make a copy of Error Class to the target sheet'
    Worksheets("QA Data").Range(Cells(1, 8), Cells(LastRow, 8)).Copy _
    Destination:=Worksheets("Data").Range("F1")                 'Column F: Error class'
    'Create Headers for Notebook and Page Number columns'
    Worksheets("Data").Cells(1, 7).value = "Previous Reviewer"  'Column G: Previous Reviewer
    Worksheets("Data").Cells(1, 8).value = "Data Reviewer"      'Column H: Data Reviewer
    Worksheets("Data").Cells(1, 9).value = "Released by"        'Column I: Released by
    Worksheets("Data").Cells(1, 10).value = "Note Book"         'Column J: Notebook
    Worksheets("Data").Cells(1, 11).value = "Page"              'Column K; Page
    'processing data row by row
    For i = 2 To LastRow
        'Parse Notebook and Page Number from the source sheet to the target sheet'
        Cells(i, 7).Select
        col_g(i) = Cells(i, 7).value
        p1 = InStr(col_g(i), "Book ")
        p2 = InStr(col_g(i), "page ")
        nb(i) = Mid(Cells(i, 7).value, p1 + 5, 5)               'fill array nb() with notebook number
        pg(i) = Mid(Cells(i, 7).value, p2 + 5, 2)               'fill array pg() with page number
        Worksheets("Data").Cells(i, 10).value = nb(i)           'fill notebook number into Column J of worksheet "Data"
        Worksheets("Data").Cells(i, 11).value = pg(i)           'fill page number into Column K of worksheet "Data"
        col_j(i) = Cells(i, 10).value                           'fill value of Previous Reviewer from column J of worksheet "QA Data" into array col_j()
        col_m(i) = Cells(i, 13).value                           'fill value of Comment  from column M of worksheet "QA Data" into array col_m()
        'Set the value of Column J of worksheet "QA Data" that contains "N/A" and "?" to blank
        If InStr(col_j(i), "N/A") > 0 Then
            col_j(i) = ""
            Else
            If InStr(col_j(i), "?") > 0 Then
                col_j(i) = ""
                Else
            End If
            col_j(i) = col_j(i)
        End If
        'Matches pattern "Data review" to "Comment" to see if the pattern exists
        drpos(i) = InStr(col_m(i), "Data review")
        If drpos(i) <> 0 Then
            dr_rev = col_j(i)
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
