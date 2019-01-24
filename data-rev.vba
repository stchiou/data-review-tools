Attribute VB_Name = "DataReviewer"
'-------------------------------------------------------------------------------------------------------------------------------
'VBA Script for processing and summarize Data Reviewer Error, v1.0
'by Sean Chiou
'Jan 7, 2019
'-------------------------------------------------------------------------------------------------------------------------------
Public LastRow As Integer               'Last row on the spreadsheet
Public dr_name() As String
Public dr_num As Integer
Public res_name_count As Integer
Sub DR_GenData()
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
    Dim temppos As Integer               'Temporary position holder for "  " found in Column M
    Dim i As Integer
    Dim j As Integer
    Dim word_count As Integer
'-------------------------------------------------------------------------------------------------------------------------------
'Create a new sheet for consolidated data'
    sheets.Add after:=sheets("QA Data")
    sheets(sheets.count).Select
    sheets(sheets.count).Name = "Data"
    sheets.Add after:=sheets("Data")
    sheets(sheets.count).Select
    sheets(sheets.count).Name = "Results"
'Remove extra blank lines on the spreadsheet'
    'Worksheets("Data").Range("A:A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
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
    Destination:=Worksheets("Data").Range("B1")
'Make a copy of Column C (Lot Numbers) from Sheet "QA Data" to Column C of Sheet "Data"
    Worksheets("QA Data").Range(Cells(1, 3), Cells(LastRow, 3)).Copy _
    Destination:=Worksheets("Data").Range("C1")
'Make a copy of Column D (List Number) from Sheet "QA Data" to Column D of Sheet "Data"
    Worksheets("QA Data").Range(Cells(1, 4), Cells(LastRow, 4)).Copy _
    Destination:=Worksheets("Data").Range("D1")
'Make a copy of Column F (Error Type) from Sheet "QA Data" to Column E of Sheet "Data"
    Worksheets("QA Data").Range(Cells(1, 6), Cells(LastRow, 6)).Copy _
    Destination:=Worksheets("Data").Range("E1")
'Make a copy of column H (Error Class) from Sheet "QA Data" to Column F of Sheet "Data"
    Worksheets("QA Data").Range(Cells(1, 8), Cells(LastRow, 8)).Copy _
    Destination:=Worksheets("Data").Range("F1")
'Create Headers for Notebook and Page Number columns'
    Worksheets("Data").Cells(1, 7).Value = "Previous Reviewer"      'Column G: Previous Reviewer
    Worksheets("Data").Cells(1, 8).Value = "Data Reviewer"          'Column H: Data Reviewer
    Worksheets("Data").Cells(1, 9).Value = "Released by"            'Column I: Released by
    Worksheets("Data").Cells(1, 10).Value = "Note Book"             'Column J: Notebook
    Worksheets("Data").Cells(1, 11).Value = "Page"                  'Column K; Page
'processing data row by row
    For i = 2 To LastRow
    'Parse Notebook and Page Number from the source sheet to the target sheet'
        Cells(i, 7).Select
        col_g(i) = Cells(i, 7).Value
        p1 = InStr(col_g(i), "Book ")
        p2 = InStr(col_g(i), "page ")
        nb(i) = Mid(Cells(i, 7).Value, p1 + 5, 5)               'fill array nb() with notebook number
        pg(i) = Mid(Cells(i, 7).Value, p2 + 5, 2)               'fill array pg() with page number
        Worksheets("Data").Cells(i, 10).Value = nb(i)           'fill notebook number into Column J of worksheet "Data"
        Worksheets("Data").Cells(i, 11).Value = pg(i)           'fill page number into Column K of worksheet "Data"
        col_j(i) = Cells(i, 10).Value                           'fill value of Previous Reviewer from column J of worksheet "QA Data" into array col_j()
        col_m(i) = Cells(i, 13).Value                           'fill value of Comment  from column M of worksheet "QA Data" into array col_m()
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
        word_count = UBound(Split(col_m(i), "  "), 1) + 1
        Dim splited_col_m() As String
        splited_col_m = Split(col_m(i), "  ")
        For j = 0 To word_count
            If InStr(Trim(splited_col_m(j)), "Data review") > 0 Then
                If InStr(Trim(splited_col_m(j)), "Data reviewer") > 0 Then
                    dr(i) = Mid(Trim(splited_col_m(j)), 14, Len(splited_col_m(j)))
                Else
                    dr(i) = Mid(Trim(splited_col_m(j)), 12, Len(splited_col_m(j)))
                End If
                Exit For
            Else
                dr(i) = ""
                If InStr(Trim(splited_col_m(j)), "Released by") > 0 Then
                    rl(i) = Mid(Trim(splited_col_m(j)), 12, Len(splited_col_m(j)))
                    Exit For
                Else
                    rl(i) = ""
                End If
            End If
        Next j
       
    'Matches pattern "Data review" to "Comment" to see if the pattern exists
        Worksheets("Data").Cells(i, 7).Value = col_j(i)  'Column G: previous Reviewer
        Worksheets("Data").Cells(i, 8).Value = Trim(dr(i))     'Column H: Data Reviewer
        Worksheets("Data").Cells(i, 9).Value = Trim(rl(i))     'Column I: Released by
    Next i
    Worksheets("Data").Activate
    ActiveSheet.Buttons.Add Range("L1").Left, Range("L1").Top, Range("L1").Width, Range("L1").Height
    ActiveSheet.Shapes.Range(Array("Button 1")).Select
    Selection.Characters.Text = "Summary"
    Selection.OnAction = "tabulate"
    MsgBox "Click on Summary Button (L1) to summarize data after finish editing."
    Cells(5, 12).Activate
End Sub
Sub tabulate()
    Dim i As Integer
    Dim j As Integer
    Dim to_be_matched As String
    'copy reviewer's name, error class, and error type to result sheet'
    Worksheets("Data").Select
    Range(Cells(1, 7), Cells(LastRow, 7)).Copy _
    Destination:=Worksheets("Results").Range("A1")
    tempstr = "A" & LastRow + 1
    Range(Cells(2, 8), Cells(LastRow, 8)).Copy _
    Destination:=Worksheets("Results").Range(tempstr)
    tempstr = "A" & LastRow * 2 + 1
    Range(Cells(2, 9), Cells(LastRow, 9)).Copy _
    Destination:=Worksheets("Results").Range(tempstr)
    Range(Cells(1, 6), Cells(LastRow, 6)).Copy _
    Destination:=Worksheets("Results").Range("B1")
    tempstr = "B" & LastRow + 1
    Range(Cells(2, 6), Cells(LastRow, 6)).Copy _
    Destination:=Worksheets("Results").Range(tempstr)
    tempstr = "B" & LastRow * 2 + 1
    Range(Cells(2, 6), Cells(LastRow, 6)).Copy _
    Destination:=Worksheets("Results").Range(tempstr)
    Range(Cells(1, 5), Cells(LastRow, 5)).Copy _
    Destination:=Worksheets("Results").Range("C1")
    tempstr = "C" & LastRow + 1
    Range(Cells(2, 5), Cells(LastRow, 5)).Copy _
    Destination:=Worksheets("Results").Range(tempstr)
    tempstr = "C" & LastRow * 2 + 1
    Range(Cells(2, 5), Cells(LastRow, 5)).Copy _
    Destination:=Worksheets("Results").Range(tempstr)
    Worksheets("Results").Range("A:A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    res_name_count = Worksheets("Results").Cells(1, 3).End(xlDown).Row
    tempstr = "C" & res_name_count
    ActiveWorkbook.Worksheets("Results").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Results").Sort.SortFields.Add Key:=Range("A2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Results").Sort
        .SetRange Range("A2:" & tempstr)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Worksheets("Results").Range("A1").Value = "Name"
    dr_num = Worksheets("names").Cells(1, 1).End(xlDown).Row
    ReDim dr_name(dr_num) As String
    For i = 1 To dr_num
        dr_name(i) = Worksheets("names").Cells(i, 1).Value
    Next i
    For j = 2 To res_name_count
        to_be_matched = Worksheets("Results").Cells(j, 1).Value
        For i = 1 To dr_num
            If InStr(dr_name(i), to_be_matched) > 0 Then
                Worksheets("Results").Cells(j, 1).Value = Worksheets("names").Cells(i, 4).Value
            Else
            End If
        Next i
    Next j
    Worksheets("Results").Activate
    ActiveSheet.Buttons.Add Range("D1").Left, Range("D1").Top, Range("D1").Width, Range("D1").Height
    ActiveSheet.Shapes.Range(Array("Button 1")).Select
    Selection.Characters.Text = "Continue"
    Selection.OnAction = "Summarize"
    MsgBox "Click on Continue Button (D1) to continue summarize data after finish editing."
    Cells(5, 12).Activate
End Sub
Sub summarize()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim unique_name_num As Integer
    Dim unique_type_num As Integer
    Dim unique_name() As String
    Dim unique_type() As String
    Dim temp() As Integer
    Worksheets("Results").Range("A2:A" & res_name_count).AdvancedFilter Action:=xlFilterCopy, copytorange:=Range("F1"), unique:=True
    Worksheets("Results").Range("C2:C" & res_name_count).AdvancedFilter Action:=xlFilterCopy, copytorange:=Range("J1"), unique:=True
    Worksheets("Results").Cells(1, 6).Value = "Name"
    Worksheets("Results").Cells(1, 7).Value = "Critical"
    Worksheets("Results").Cells(1, 8).Value = "Major"
    Worksheets("Results").Cells(1, 9).Value = "Minor"
    unique_type_num = Worksheets("Results").Cells(1, 10).End(xlDown).Row
    unique_name_num = Worksheets("Results").Cells(1, 6).End(xlDown).Row
    Worksheets("Results").Range(Cells(2, 10), Cells(unique_type_num, 10)).Copy
    Worksheets("Results").Range("J1").PasteSpecial Transpose:=True
    Worksheets("Results").Range(Cells(2, 10), Cells(unique_type_num, 10)).Value = ""
    Dim col_count As Integer
    col_count = sheets("Results").Cells(1, 6).End(xlToLeft).Column
    ReDim temp(unique_name_num, unique_type_num + 3) As Integer
    For i = 2 To unique_name_num
        unique_name(i) = Cells(i, 6).Value
    Next i
    For i = 1 To unique_type_num - 1
        unique_type(i) = Cells(1, 9 + i)
    Next i
    Worksheets("Results").Activate
    For i = 2 To unique_name_num
        For j = 7 To col_count
            For k = 2 To res_name_count
                
            Next k
        Next j
    Next i
End Sub

