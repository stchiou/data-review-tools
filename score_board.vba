Attribute VB_Name = "Data_Reviewer_ScoreBoard"
Public week_num As Integer
Public yr As Integer
Public wn As String
Public ReviewerName() As String
Public ReviewerNum As Integer
Sub reviewer_score()
'-----------------------------------------------------------------------------
'Prepare for data entry
'-----------------------------------------------------------------------------
    Dim date1 As Date
    Dim date2 As Date
    Dim btn1 As Button
    Dim btn2 As Button
    Dim btn3 As Button
    Dim typelist As String
    Dim lastrow As Long

typelist = "Impurity/Potency, Impurity, Potency, Assay, ID"
yr = InputBox("Please enter the year of the records.")
week_num = InputBox("Please enter week number (1-52).")
date1 = DateSerial(yr, 1, (week_num - 1) * 7 + 1)
date2 = date1 + 6
If week_num > 9 Then
    wn = week_num
Else
    wn = "0" & week_num
End If
Worksheets.Add(after:=Worksheets(Worksheets.Count)).name = "Week_" & wn & "_" & yr
    Cells(1, 1).Value = "Review Date"
    Cells(1, 2).Value = "Name"
    Cells(1, 3).Value = "Assigment Type"
    Cells(1, 4).Value = "Lot Assigned"
    Cells(1, 5).Value = "Lot with Error"
    Cells(1, 6).Value = "Number of Error"
    Cells(1, 7).Value = "Additional Error"
    Cells(1, 8).Value = "Penalty"
    Cells(1, 9).Value = "Score"
Range("A2:A1048576").Select
With Selection.Validation
        .Delete
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, formula1:=date1, formula2:=date2
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "Wrong Date"
        .InputMessage = "Enter date between " & date1 & " and " & date2 & "."
        .ErrorMessage = "Week " & week_num & " is between " & date1 & " and " & date2 & "."
        .ShowInput = True
        .ShowError = True
End With
    ReviewerNum = Worksheets("Names").Cells(1, 1).End(xlDown).Row
    ReDim ReviewerName(ReviewerNum) As String
    For i = 1 To ReviewerNum
        ReviewerName(i) = Worksheets("Names").Cells(i, 1).Value
    Next i
    
    Range("B2").Select
    With Selection.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, formula1:="=Names!$A$1:$A$" & ReviewerNum
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Data Reviewer Name"
        .ErrorTitle = ""
        .InputMessage = "Select name from the drop-down list."
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Selection.AutoFill Destination:=Range("B2:B1048576"), Type:=xlFillDefault
    Range("B2").End(xlDown).Select
    Range("C2").Select
    With Selection.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, formula1:=typelist
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Assignment Type"
        .ErrorTitle = "Assigment type not supported"
        .InputMessage = "Select assignment type from the list"
        .ErrorMessage = "Valid entries are Impurity/Potency, Impurity, Potency, Assay, or ID."
        .ShowInput = True
        .ShowError = True
    End With
    Selection.AutoFill Destination:=Range("C2:C1048576"), Type:=xlFillDefault
    Range("C2:C2").End(xlDown).Select
entry_prompt:
Set btn1 = ActiveSheet.Buttons.Add(Range("L1").Left, 0, 120, 25)
btn1.Select
With Selection
    .OnAction = "Compute"
    .Caption = "Compute Scores"
    .Font.Bold = True
End With
Worksheets("Week_" & wn & "_" & yr).Activate
Cells(1, 1).Select
MsgBox ("Enter Data in columns A-F. Click the 'Calculate' button to compute penalty and final score.")
Set btn2 = ActiveSheet.Buttons.Add(Range("L5").Left, 30, 120, 25)
btn2.Select
With Selection
    .OnAction = "Gen_report"
    .Caption = "Generate Report"
    .Font.Bold = True
End With
MsgBox ("Click Report to generate monthly report.")
Set btn3 = ActiveSheet.Buttons.Add(Range("L9").Left, 60, 120, 25)
btn3.Select
With Selection
    .OnAction = "reviewer_score"
    .Caption = "Add New Sheet"
    .Font.Bold = True
End With
Cells(2, 1).Activate
End Sub
Sub Compute()
'--------------------------------------------------------------------------------
'variables for store input
'--------------------------------------------------------------------------------
    Dim entry_date() As Date
    Dim reviewer() As String
    Dim assigment() As String
    Dim LotAssigned() As Integer
    Dim LotError() As Integer
    Dim ErrorNum() As Integer
    Dim ExtraError() As Integer
    Dim penalty() As Double
    Dim score() As Double
    Dim record_num As Long
    Dim i As Integer
    Dim j As Integer
    Dim ShName As String
'------------------------------------------------------------------------------
'variables for specifying report
'------------------------------------------------------------------------------
    Dim report_type As Integer
    Dim report_year As Integer
    Dim report_month As Integer
    Dim report_quarter As Integer
    Dim report_week As Integer
ShName = ActiveWorkbook.ActiveSheet.name
If week_num <> 0 Then
    week_num = week_num
Else
    week_num = Mid(ShName, 6, 2)
End If
Cells(1, 1).Activate
record_num = Cells(1, 1).End(xlDown).Row
    ReDim entry_date(record_num) As Date
    ReDim reviewer(record_num) As String
    ReDim assignment(record_num) As String
    ReDim LotAssigned(record_num) As Integer
    ReDim LotError(record_num) As Integer
    ReDim ErrorNum(record_num) As Integer
    ReDim ExtraError(record_num) As Integer
    ReDim penalty(record_num) As Double
    ReDim score(record_num) As Double
    Cells(2, 1).Activate
    For i = 2 To record_num
      entry_date(i) = ActiveCell.Value
      reviewer(i) = ActiveCell.Offset(0, 1).Value
      assignment(i) = ActiveCell.Offset(0, 2).Value
      Select Case assignment(i)
        Case Is = "Impurity/Potency"
            assignment(i) = 5
        Case Is = "Impurity"
            assignment(i) = 4
        Case Is = "Potency"
            assignment(i) = 3
        Case Is = "Assay"
            assignment(i) = 2
        Case Is = "ID"
            assignment(i) = 1
      End Select
      LotAssigned(i) = ActiveCell.Offset(0, 3).Value
      LotError(i) = ActiveCell.Offset(0, 4).Value
      ErrorNum(i) = ActiveCell.Offset(0, 5).Value
      ExtraError(i) = ActiveCell.Offset(0, 6).Value
      penalty(i) = (LotError(i) * ErrorNum(i)) / (assignment(i)) + (ExtraError(i) / assignment(i))
      score(i) = (assignment(i) * LotAssigned(i) - penalty(i)) / (assignment(i) * LotAssigned(i)) * 100
      If score(i) < 0 Then
        score(i) = 0
      Else
        score(i) = score(i)
      End If
      ActiveCell.Offset(0, 7).Value = penalty(i)
      ActiveCell.Offset(0, 8).Value = score(i)
      ActiveCell.Offset(1, 0).Activate
    Next i
End Sub
'--------------------------------------------------------------------------------------
'Generate Report
'--------------------------------------------------------------------------------------
Sub Gen_report()
    Dim start_week As Integer
    Dim end_week As Integer
    Dim week_num As Integer
    Dim year As Integer
    Dim month As Integer
    Dim report_start_week As Integer
    Dim report_end_week As Integer
    Dim SheetNum As Integer
    Dim SheetName() As String
    Dim rowNum() As Integer
    Dim ReportRecNum As Integer
    Dim summary() As Double
    Dim month_start As Date
    Dim month_end As Date
    Dim i As Integer
    Dim j As Integer
    Dim temp() As Double
    Dim month_name As String
    Dim Curr_Rec As Integer
    Dim ReportSheet() As Variant
    Dim NextRow As Long
    Dim RecNum As Integer
'---------------------------------------------------------
'Array dimension
'----------------

'---------------------------------------------------------
    ReviewerNum = ActiveWorkbook.Worksheets("Names").Cells(1, 1).End(xlDown).Row
    ReDim ReviewerName(ReviewerNum) As String
    For i = 1 To ReviewerNum
        ReviewerName(i) = Worksheets("Names").Cells(i, 1).Value
    Next i
    year = InputBox("Enter the year of the report")
    month = InputBox("Enter the month for report: " _
        & vbCr & "1. January" _
        & vbCr & "2. February" _
        & vbCr & "3. March" _
        & vbCr & "4. April" _
        & vbCr & "5. May" _
        & vbCr & "6. June" _
        & vbCr & "7. July" _
        & vbCr & "8. August" _
        & vbCr & "9. September" _
        & vbCr & "10. October" _
        & vbCr & "11. November" _
        & vbCr & "12. December")
    month_start = year & "/" & month & "/1"
    month_end = WorksheetFunction.EoMonth(month_start, 0)
    report_start_week = WorksheetFunction.WeekNum(month_start)
    report_end_week = WorksheetFunction.WeekNum(month_end)
    SheetNum = report_end_week - report_start_week + 1
    ReDim SheetName(report_end_week) As String
    ReDim rowNum(report_end_week) As Integer
    Select Case month
        Case Is = 1
            month_name = "January"
        Case Is = 2
            month_name = "February"
        Case Is = 3
            month_name = "March"
        Case Is = 4
            month_name = "April"
        Case Is = 5
            month_name = "May"
        Case Is = 6
            month_name = "June"
        Case Is = 7
            month_name = "July"
        Case Is = 8
            month_name = "August"
        Case Is = 9
            month_name = "September"
        Case Is = 10
            month_name = "October"
        Case Is = 11
            month_name = "November"
        Case Is = 12
            month_name = "December"
    End Select
    ReportRecNum = 0
    For i = report_start_week To report_end_week
        If i < 10 Then
            wn = "0" & i
        Else
            wn = i
        End If
        SheetName(i) = "Week_" & wn & "_" & year
        rowNum(i) = Worksheets(SheetName(i)).Cells(1, 1).End(xlDown).Row
        ReportRecNum = ReportRecNum + rowNum(i) - 1
    Next i
MsgBox ("Processing monthly report of " & month_name & " " & year & " with " & ReportRecNum & " records.")
ReDim temp(5, ReportRecNum) As Double
Curr_Rec = 0
For i = report_start_week To report_end_week
    Worksheets(SheetName(i)).Activate
    For j = 2 To rowNum(i)
        Curr_Rec = Curr_Rec + 1
        Cells(j, 1).Activate
            temp(1, Curr_Rec) = ActiveCell.Value
            ActiveCell.Offset(0, 1).Activate
        Select Case ActiveCell.Value
            Case Is = ReviewerName(1)
                temp(2, Curr_Rec) = 1
            Case Is = ReviewerName(2)
                temp(2, Curr_Rec) = 2
            Case Is = ReviewerName(3)
                temp(2, Curr_Rec) = 3
            Case Is = ReviewerName(4)
                temp(2, Curr_Rec) = 4
            Case Is = ReviewerName(5)
                temp(2, Curr_Rec) = 5
            Case Is = ReviewerName(6)
                temp(2, Curr_Rec) = 6
            Case Is = ReviewerName(7)
                temp(2, Curr_Rec) = 7
            Case Is = ReviewerName(8)
                temp(2, Curr_Rec) = 8
            Case Is = ReviewerName(9)
                temp(2, Curr_Rec) = 9
            Case Is = ReviewerName(10)
                temp(2, Curr_Rec) = 10
            Case Is = ReviewerName(11)
                temp(2, Curr_Rec) = 11
            Case Is = ReviewerName(12)
                temp(2, Curr_Rec) = 12
            Case Is = ReviewerName(13)
                temp(2, Curr_Rec) = 13
            Case Is = ReviewerName(14)
                temp(2, Curr_Rec) = 14
            Case Is = ReviewerName(15)
                temp(2, Curr_Rec) = 15
            Case Is = ReviewerName(16)
                temp(2, Curr_Rec) = 16
            Case Is = ReviewerName(17)
                temp(2, Curr_Rec) = 17
            Case Is = ReviewerName(18)
                temp(2, Curr_Rec) = 18
            Case Is = ReviewerName(19)
                temp(2, Curr_Rec) = 19
            Case Is = ReviewerName(20)
                temp(2, Curr_Rec) = 20
            Case Is = ReviewerName(21)
                temp(2, Curr_Rec) = 21
            Case Is = ReviewerName(22)
                temp(2, Curr_Rec) = 22
            Case Is = ReviewerName(23)
                temp(2, Curr_Rec) = 23
            Case Is = ReviewerName(24)
                temp(2, Curr_Rec) = 24
            Case Is = ReviewerName(25)
                temp(2, Curr_Rec) = 25
            Case Is = ReviewerName(26)
                temp(2, Curr_Rec) = 26
            Case Is = ReviewerName(27)
                temp(2, Curr_Rec) = 27
            Case Is = ReviewerName(28)
                temp(2, Curr_Rec) = 28
            Case Is = ReviewerName(29)
                temp(2, Curr_Rec) = 29
        End Select
        Select Case ActiveCell.Offset(0, 1).Value
            Case Is = "Impurity/Potency"
                temp(3, Curr_Rec) = 5
            Case Is = "Impurity"
                temp(3, Curr_Rec) = 4
            Case Is = "Potency"
                temp(3, Curr_Rec) = 3
            Case Is = "Assay"
                temp(3, Curr_Rec) = 2
            Case Is = "ID"
                temp(3, Curr_Rec) = 1
        End Select
        temp(4, Curr_Rec) = ActiveCell.Offset(0, 2).Value
        temp(5, Curr_Rec) = ActiveCell.Offset(0, 7).Value
       
    Next j
Next i
MsgBox ("Data loaded")
With Application
    .SheetsInNewWorkbook = ReviewerNum
    .Workbooks.Add
    .SheetsInNewWorkbook = ReviewerNum
End With
For i = 1 To ReviewerNum
   Sheets("Sheet" & i).name = ReviewerName(i)
   Worksheets(ReviewerName(i)).Cells(1, 1).Value = "Date"
   Worksheets(ReviewerName(i)).Cells(1, 2).Value = "Name"
   Worksheets(ReviewerName(i)).Cells(1, 3).Value = "Type"
   Worksheets(ReviewerName(i)).Cells(1, 4).Value = "Lot"
   Worksheets(ReviewerName(i)).Cells(1, 5).Value = "Score"
Next i
ActiveWorkbook.SaveAs Filename:="\\hpdrmf01\f_DRIVE\CQ Lab\Data Review Tracking\Test\" & month_name & year & ".xlsx", Password:="Seniors"
For i = 1 To ReportRecNum
    Select Case temp(2, i)
        Case Is = 1
            Worksheets(ReviewerName(1)).Activate
            NextRow = Worksheets(ReviewerName(1)).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
            Worksheets(ReviewerName(1)).Cells(NextRow, 1).Activate
            ActiveCell.Value = temp(1, i)
            ActiveCell.NumberFormat = "mm/dd/yyyy"
            ActiveCell.Offset(0, 1).Value = ReviewerName(1)
            Select Case temp(3, i)
                Case Is = 1
                    ActiveCell.Offset(0, 2).Value = "ID"
                Case Is = 2
                    ActiveCell.Offset(0, 2).Value = "Assay"
                Case Is = 3
                    ActiveCell.Offset(0, 2).Value = "Potency"
                Case Is = 4
                    ActiveCell.Offset(0, 2).Value = "Impurity"
                Case Is = 5
                    ActiveCell.Offset(0, 2).Value = "Impurity/Potency"
            End Select
            ActiveCell.Offset(0, 3).Value = temp(4, i)
            ActiveCell.Offset(0, 4).Value = temp(5, i)
        Case Is = 2
            Worksheets(ReviewerName(2)).Activate
            NextRow = Worksheets(ReviewerName(2)).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
            Worksheets(ReviewerName(2)).Cells(NextRow, 1).Activate
            ActiveCell.Value = temp(1, i)
            ActiveCell.NumberFormat = "mm/dd/yyyy"
            ActiveCell.Offset(0, 1).Value = ReviewerName(2)
            Select Case temp(3, i)
                Case Is = 1
                    ActiveCell.Offset(0, 2).Value = "ID"
                Case Is = 2
                    ActiveCell.Offset(0, 2).Value = "Assay"
                Case Is = 3
                    ActiveCell.Offset(0, 2).Value = "Potency"
                Case Is = 4
                    ActiveCell.Offset(0, 2).Value = "Impurity"
                Case Is = 5
                    ActiveCell.Offset(0, 2).Value = "Impurity/Potency"
            End Select
            ActiveCell.Offset(0, 3).Value = temp(4, i)
            ActiveCell.Offset(0, 4).Value = temp(5, i)
        Case Is = 3
            Worksheets(ReviewerName(3)).Activate
            NextRow = Worksheets(ReviewerName(3)).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
            Worksheets(ReviewerName(3)).Cells(NextRow, 1).Activate
            ActiveCell.Value = temp(1, i)
            ActiveCell.NumberFormat = "mm/dd/yyyy"
            ActiveCell.Offset(0, 1).Value = ReviewerName(3)
            Select Case temp(3, i)
                Case Is = 1
                    ActiveCell.Offset(0, 2).Value = "ID"
                Case Is = 2
                    ActiveCell.Offset(0, 2).Value = "Assay"
                Case Is = 3
                    ActiveCell.Offset(0, 2).Value = "Potency"
                Case Is = 4
                    ActiveCell.Offset(0, 2).Value = "Impurity"
                Case Is = 5
                    ActiveCell.Offset(0, 2).Value = "Impurity/Potency"
            End Select
            ActiveCell.Offset(0, 3).Value = temp(4, i)
            ActiveCell.Offset(0, 4).Value = temp(5, i)
        Case Is = 4
            Worksheets(ReviewerName(4)).Activate
            NextRow = Worksheets(ReviewerName(4)).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
            Worksheets(ReviewerName(4)).Cells(NextRow, 1).Activate
            ActiveCell.Value = temp(1, i)
            ActiveCell.NumberFormat = "mm/dd/yyyy"
            ActiveCell.Offset(0, 1).Value = ReviewerName(4)
            Select Case temp(3, i)
                Case Is = 1
                    ActiveCell.Offset(0, 2).Value = "ID"
                Case Is = 2
                    ActiveCell.Offset(0, 2).Value = "Assay"
                Case Is = 3
                    ActiveCell.Offset(0, 2).Value = "Potency"
                Case Is = 4
                    ActiveCell.Offset(0, 2).Value = "Impurity"
                Case Is = 5
                    ActiveCell.Offset(0, 2).Value = "Impurity/Potency"
            End Select
            ActiveCell.Offset(0, 3).Value = temp(4, i)
            ActiveCell.Offset(0, 4).Value = temp(5, i)
        Case Is = 5
            Worksheets(ReviewerName(5)).Activate
            NextRow = Worksheets(ReviewerName(5)).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
            Worksheets(ReviewerName(5)).Cells(NextRow, 1).Activate
            ActiveCell.Value = temp(1, i)
            ActiveCell.NumberFormat = "mm/dd/yyyy"
            ActiveCell.Offset(0, 1).Value = ReviewerName(5)
            Select Case temp(3, i)
                Case Is = 1
                    ActiveCell.Offset(0, 2).Value = "ID"
                Case Is = 2
                    ActiveCell.Offset(0, 2).Value = "Assay"
                Case Is = 3
                    ActiveCell.Offset(0, 2).Value = "Potency"
                Case Is = 4
                    ActiveCell.Offset(0, 2).Value = "Impurity"
                Case Is = 5
                    ActiveCell.Offset(0, 2).Value = "Impurity/Potency"
            End Select
            ActiveCell.Offset(0, 3).Value = temp(4, i)
            ActiveCell.Offset(0, 4).Value = temp(5, i)
        Case Is = 6
            Worksheets(ReviewerName(6)).Activate
            NextRow = Worksheets(ReviewerName(6)).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
            Worksheets(ReviewerName(6)).Cells(NextRow, 1).Activate
            ActiveCell.Value = temp(1, i)
            ActiveCell.NumberFormat = "mm/dd/yyyy"
            ActiveCell.Offset(0, 1).Value = ReviewerName(6)
            Select Case temp(3, i)
                Case Is = 1
                    ActiveCell.Offset(0, 2).Value = "ID"
                Case Is = 2
                    ActiveCell.Offset(0, 2).Value = "Assay"
                Case Is = 3
                    ActiveCell.Offset(0, 2).Value = "Potency"
                Case Is = 4
                    ActiveCell.Offset(0, 2).Value = "Impurity"
                Case Is = 5
                    ActiveCell.Offset(0, 2).Value = "Impurity/Potency"
            End Select
            ActiveCell.Offset(0, 3).Value = temp(4, i)
            ActiveCell.Offset(0, 4).Value = temp(5, i)
        Case Is = 7
            Worksheets(ReviewerName(7)).Activate
            NextRow = Worksheets(ReviewerName(7)).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
            Worksheets(ReviewerName(7)).Cells(NextRow, 1).Activate
            ActiveCell.Value = temp(1, i)
            ActiveCell.NumberFormat = "mm/dd/yyyy"
            ActiveCell.Offset(0, 1).Value = ReviewerName(7)
            Select Case temp(3, i)
                Case Is = 1
                    ActiveCell.Offset(0, 2).Value = "ID"
                Case Is = 2
                    ActiveCell.Offset(0, 2).Value = "Assay"
                Case Is = 3
                    ActiveCell.Offset(0, 2).Value = "Potency"
                Case Is = 4
                    ActiveCell.Offset(0, 2).Value = "Impurity"
                Case Is = 5
                    ActiveCell.Offset(0, 2).Value = "Impurity/Potency"
            End Select
            ActiveCell.Offset(0, 3).Value = temp(4, i)
            ActiveCell.Offset(0, 4).Value = temp(5, i)
        Case Is = 8
            Worksheets(ReviewerName(8)).Activate
            NextRow = Worksheets(ReviewerName(8)).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
            Worksheets(ReviewerName(8)).Cells(NextRow, 1).Activate
            ActiveCell.Value = temp(1, i)
            ActiveCell.NumberFormat = "mm/dd/yyyy"
            ActiveCell.Offset(0, 1).Value = ReviewerName(8)
            Select Case temp(3, i)
                Case Is = 1
                    ActiveCell.Offset(0, 2).Value = "ID"
                Case Is = 2
                    ActiveCell.Offset(0, 2).Value = "Assay"
                Case Is = 3
                    ActiveCell.Offset(0, 2).Value = "Potency"
                Case Is = 4
                    ActiveCell.Offset(0, 2).Value = "Impurity"
                Case Is = 5
                    ActiveCell.Offset(0, 2).Value = "Impurity/Potency"
            End Select
            ActiveCell.Offset(0, 3).Value = temp(4, i)
            ActiveCell.Offset(0, 4).Value = temp(5, i)
        Case Is = 9
            Worksheets(ReviewerName(9)).Activate
            NextRow = Worksheets(ReviewerName(9)).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
            Worksheets(ReviewerName(9)).Cells(NextRow, 1).Activate
            ActiveCell.Value = temp(1, i)
            ActiveCell.NumberFormat = "mm/dd/yyyy"
            ActiveCell.Offset(0, 1).Value = ReviewerName(9)
            Select Case temp(3, i)
                Case Is = 1
                    ActiveCell.Offset(0, 2).Value = "ID"
                Case Is = 2
                    ActiveCell.Offset(0, 2).Value = "Assay"
                Case Is = 3
                    ActiveCell.Offset(0, 2).Value = "Potency"
                Case Is = 4
                    ActiveCell.Offset(0, 2).Value = "Impurity"
                Case Is = 5
                    ActiveCell.Offset(0, 2).Value = "Impurity/Potency"
            End Select
            ActiveCell.Offset(0, 3).Value = temp(4, i)
            ActiveCell.Offset(0, 4).Value = temp(5, i)
        Case Is = 10
            Worksheets(ReviewerName(10)).Activate
            NextRow = Worksheets(ReviewerName(10)).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
            Worksheets(ReviewerName(10)).Cells(NextRow, 1).Activate
            ActiveCell.Value = temp(1, i)
            ActiveCell.NumberFormat = "mm/dd/yyyy"
            ActiveCell.Offset(0, 1).Value = ReviewerName(10)
            Select Case temp(3, i)
                Case Is = 1
                    ActiveCell.Offset(0, 2).Value = "ID"
                Case Is = 2
                    ActiveCell.Offset(0, 2).Value = "Assay"
                Case Is = 3
                    ActiveCell.Offset(0, 2).Value = "Potency"
                Case Is = 4
                    ActiveCell.Offset(0, 2).Value = "Impurity"
                Case Is = 5
                    ActiveCell.Offset(0, 2).Value = "Impurity/Potency"
            End Select
            ActiveCell.Offset(0, 3).Value = temp(4, i)
            ActiveCell.Offset(0, 4).Value = temp(5, i)
        Case Is = 11
            Worksheets(ReviewerName(11)).Activate
            NextRow = Worksheets(ReviewerName(11)).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
            Worksheets(ReviewerName(11)).Cells(NextRow, 1).Activate
            ActiveCell.Value = temp(1, i)
            ActiveCell.NumberFormat = "mm/dd/yyyy"
            ActiveCell.Offset(0, 1).Value = ReviewerName(11)
            Select Case temp(3, i)
                Case Is = 1
                    ActiveCell.Offset(0, 2).Value = "ID"
                Case Is = 2
                    ActiveCell.Offset(0, 2).Value = "Assay"
                Case Is = 3
                    ActiveCell.Offset(0, 2).Value = "Potency"
                Case Is = 4
                    ActiveCell.Offset(0, 2).Value = "Impurity"
                Case Is = 5
                    ActiveCell.Offset(0, 2).Value = "Impurity/Potency"
            End Select
            ActiveCell.Offset(0, 3).Value = temp(4, i)
            ActiveCell.Offset(0, 4).Value = temp(5, i)
        Case Is = 12
            Worksheets(ReviewerName(12)).Activate
            NextRow = Worksheets(ReviewerName(12)).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
            Worksheets(ReviewerName(12)).Cells(NextRow, 1).Activate
            ActiveCell.Value = temp(1, i)
            ActiveCell.NumberFormat = "mm/dd/yyyy"
            ActiveCell.Offset(0, 1).Value = ReviewerName(12)
            Select Case temp(3, i)
                Case Is = 1
                    ActiveCell.Offset(0, 2).Value = "ID"
                Case Is = 2
                    ActiveCell.Offset(0, 2).Value = "Assay"
                Case Is = 3
                    ActiveCell.Offset(0, 2).Value = "Potency"
                Case Is = 4
                    ActiveCell.Offset(0, 2).Value = "Impurity"
                Case Is = 5
                    ActiveCell.Offset(0, 2).Value = "Impurity/Potency"
            End Select
            ActiveCell.Offset(0, 3).Value = temp(4, i)
            ActiveCell.Offset(0, 4).Value = temp(5, i)
        Case Is = 13
            Worksheets(ReviewerName(13)).Activate
            NextRow = Worksheets(ReviewerName(13)).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
            Worksheets(ReviewerName(13)).Cells(NextRow, 1).Activate
            ActiveCell.Value = temp(1, i)
            ActiveCell.NumberFormat = "mm/dd/yyyy"
            ActiveCell.Offset(0, 1).Value = ReviewerName(13)
            Select Case temp(3, i)
                Case Is = 1
                    ActiveCell.Offset(0, 2).Value = "ID"
                Case Is = 2
                    ActiveCell.Offset(0, 2).Value = "Assay"
                Case Is = 3
                    ActiveCell.Offset(0, 2).Value = "Potency"
                Case Is = 4
                    ActiveCell.Offset(0, 2).Value = "Impurity"
                Case Is = 5
                    ActiveCell.Offset(0, 2).Value = "Impurity/Potency"
            End Select
            ActiveCell.Offset(0, 3).Value = temp(4, i)
            ActiveCell.Offset(0, 4).Value = temp(5, i)
        Case Is = 14
            Worksheets(ReviewerName(14)).Activate
            NextRow = Worksheets(ReviewerName(14)).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
            Worksheets(ReviewerName(14)).Cells(NextRow, 1).Activate
            ActiveCell.Value = temp(1, i)
            ActiveCell.NumberFormat = "mm/dd/yyyy"
            ActiveCell.Offset(0, 1).Value = ReviewerName(14)
            Select Case temp(3, i)
                Case Is = 1
                    ActiveCell.Offset(0, 2).Value = "ID"
                Case Is = 2
                    ActiveCell.Offset(0, 2).Value = "Assay"
                Case Is = 3
                    ActiveCell.Offset(0, 2).Value = "Potency"
                Case Is = 4
                    ActiveCell.Offset(0, 2).Value = "Impurity"
                Case Is = 5
                    ActiveCell.Offset(0, 2).Value = "Impurity/Potency"
            End Select
            ActiveCell.Offset(0, 3).Value = temp(4, i)
            ActiveCell.Offset(0, 4).Value = temp(5, i)
        Case Is = 15
            Worksheets(ReviewerName(15)).Activate
            NextRow = Worksheets(ReviewerName(15)).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
            Worksheets(ReviewerName(15)).Cells(NextRow, 1).Activate
            ActiveCell.Value = temp(1, i)
            ActiveCell.NumberFormat = "mm/dd/yyyy"
            ActiveCell.Offset(0, 1).Value = ReviewerName(15)
            Select Case temp(3, i)
                Case Is = 1
                    ActiveCell.Offset(0, 2).Value = "ID"
                Case Is = 2
                    ActiveCell.Offset(0, 2).Value = "Assay"
                Case Is = 3
                    ActiveCell.Offset(0, 2).Value = "Potency"
                Case Is = 4
                    ActiveCell.Offset(0, 2).Value = "Impurity"
                Case Is = 5
                    ActiveCell.Offset(0, 2).Value = "Impurity/Potency"
            End Select
            ActiveCell.Offset(0, 3).Value = temp(4, i)
            ActiveCell.Offset(0, 4).Value = temp(5, i)
        Case Is = 16
            Worksheets(ReviewerName(16)).Activate
            NextRow = Worksheets(ReviewerName(16)).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
            Worksheets(ReviewerName(16)).Cells(NextRow, 1).Activate
            ActiveCell.Value = temp(1, i)
            ActiveCell.NumberFormat = "mm/dd/yyyy"
            ActiveCell.Offset(0, 1).Value = ReviewerName(16)
            Select Case temp(3, i)
                Case Is = 1
                    ActiveCell.Offset(0, 2).Value = "ID"
                Case Is = 2
                    ActiveCell.Offset(0, 2).Value = "Assay"
                Case Is = 3
                    ActiveCell.Offset(0, 2).Value = "Potency"
                Case Is = 4
                    ActiveCell.Offset(0, 2).Value = "Impurity"
                Case Is = 5
                    ActiveCell.Offset(0, 2).Value = "Impurity/Potency"
            End Select
            ActiveCell.Offset(0, 3).Value = temp(4, i)
            ActiveCell.Offset(0, 4).Value = temp(5, i)
        Case Is = 17
            Worksheets(ReviewerName(17)).Activate
            NextRow = Worksheets(ReviewerName(17)).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
            Worksheets(ReviewerName(17)).Cells(NextRow, 1).Activate
            ActiveCell.Value = temp(1, i)
            ActiveCell.NumberFormat = "mm/dd/yyyy"
            ActiveCell.Offset(0, 1).Value = ReviewerName(17)
            Select Case temp(3, i)
                Case Is = 1
                    ActiveCell.Offset(0, 2).Value = "ID"
                Case Is = 2
                    ActiveCell.Offset(0, 2).Value = "Assay"
                Case Is = 3
                    ActiveCell.Offset(0, 2).Value = "Potency"
                Case Is = 4
                    ActiveCell.Offset(0, 2).Value = "Impurity"
                Case Is = 5
                    ActiveCell.Offset(0, 2).Value = "Impurity/Potency"
            End Select
            ActiveCell.Offset(0, 3).Value = temp(4, i)
            ActiveCell.Offset(0, 4).Value = temp(5, i)
        Case Is = 18
            Worksheets(ReviewerName(18)).Activate
            NextRow = Worksheets(ReviewerName(18)).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
            Worksheets(ReviewerName(18)).Cells(NextRow, 1).Activate
            ActiveCell.Value = temp(1, i)
            ActiveCell.NumberFormat = "mm/dd/yyyy"
            ActiveCell.Offset(0, 1).Value = ReviewerName(18)
            Select Case temp(3, i)
                Case Is = 1
                    ActiveCell.Offset(0, 2).Value = "ID"
                Case Is = 2
                    ActiveCell.Offset(0, 2).Value = "Assay"
                Case Is = 3
                    ActiveCell.Offset(0, 2).Value = "Potency"
                Case Is = 4
                    ActiveCell.Offset(0, 2).Value = "Impurity"
                Case Is = 5
                    ActiveCell.Offset(0, 2).Value = "Impurity/Potency"
            End Select
            ActiveCell.Offset(0, 3).Value = temp(4, i)
            ActiveCell.Offset(0, 4).Value = temp(5, i)
        Case Is = 19
            Worksheets(ReviewerName(19)).Activate
            NextRow = Worksheets(ReviewerName(19)).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
            Worksheets(ReviewerName(19)).Cells(NextRow, 1).Activate
            ActiveCell.Value = temp(1, i)
            ActiveCell.NumberFormat = "mm/dd/yyyy"
            ActiveCell.Offset(0, 1).Value = ReviewerName(19)
            Select Case temp(3, i)
                Case Is = 1
                    ActiveCell.Offset(0, 2).Value = "ID"
                Case Is = 2
                    ActiveCell.Offset(0, 2).Value = "Assay"
                Case Is = 3
                    ActiveCell.Offset(0, 2).Value = "Potency"
                Case Is = 4
                    ActiveCell.Offset(0, 2).Value = "Impurity"
                Case Is = 5
                    ActiveCell.Offset(0, 2).Value = "Impurity/Potency"
            End Select
            ActiveCell.Offset(0, 3).Value = temp(4, i)
            ActiveCell.Offset(0, 4).Value = temp(5, i)
        Case Is = 20
            Worksheets(ReviewerName(20)).Activate
            NextRow = Worksheets(ReviewerName(20)).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
            Worksheets(ReviewerName(20)).Cells(NextRow, 1).Activate
            ActiveCell.Value = temp(1, i)
            ActiveCell.NumberFormat = "mm/dd/yyyy"
            ActiveCell.Offset(0, 1).Value = ReviewerName(20)
            Select Case temp(3, i)
                Case Is = 1
                    ActiveCell.Offset(0, 2).Value = "ID"
                Case Is = 2
                    ActiveCell.Offset(0, 2).Value = "Assay"
                Case Is = 3
                    ActiveCell.Offset(0, 2).Value = "Potency"
                Case Is = 4
                    ActiveCell.Offset(0, 2).Value = "Impurity"
                Case Is = 5
                    ActiveCell.Offset(0, 2).Value = "Impurity/Potency"
            End Select
            ActiveCell.Offset(0, 3).Value = temp(4, i)
            ActiveCell.Offset(0, 4).Value = temp(5, i)
        Case Is = 21
            Worksheets(ReviewerName(21)).Activate
            NextRow = Worksheets(ReviewerName(21)).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
            Worksheets(ReviewerName(21)).Cells(NextRow, 1).Activate
            ActiveCell.Value = temp(1, i)
            ActiveCell.NumberFormat = "mm/dd/yyyy"
            ActiveCell.Offset(0, 1).Value = ReviewerName(21)
            Select Case temp(3, i)
                Case Is = 1
                    ActiveCell.Offset(0, 2).Value = "ID"
                Case Is = 2
                    ActiveCell.Offset(0, 2).Value = "Assay"
                Case Is = 3
                    ActiveCell.Offset(0, 2).Value = "Potency"
                Case Is = 4
                    ActiveCell.Offset(0, 2).Value = "Impurity"
                Case Is = 5
                    ActiveCell.Offset(0, 2).Value = "Impurity/Potency"
            End Select
            ActiveCell.Offset(0, 3).Value = temp(4, i)
            ActiveCell.Offset(0, 4).Value = temp(5, i)
        Case Is = 22
            Worksheets(ReviewerName(22)).Activate
            NextRow = Worksheets(ReviewerName(22)).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
            Worksheets(ReviewerName(22)).Cells(NextRow, 1).Activate
            ActiveCell.Value = temp(1, i)
            ActiveCell.NumberFormat = "mm/dd/yyyy"
            ActiveCell.Offset(0, 1).Value = ReviewerName(22)
            Select Case temp(3, i)
                Case Is = 1
                    ActiveCell.Offset(0, 2).Value = "ID"
                Case Is = 2
                    ActiveCell.Offset(0, 2).Value = "Assay"
                Case Is = 3
                    ActiveCell.Offset(0, 2).Value = "Potency"
                Case Is = 4
                    ActiveCell.Offset(0, 2).Value = "Impurity"
                Case Is = 5
                    ActiveCell.Offset(0, 2).Value = "Impurity/Potency"
            End Select
            ActiveCell.Offset(0, 3).Value = temp(4, i)
            ActiveCell.Offset(0, 4).Value = temp(5, i)
        Case Is = 23
            Worksheets(ReviewerName(23)).Activate
            NextRow = Worksheets(ReviewerName(23)).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
            Worksheets(ReviewerName(23)).Cells(NextRow, 1).Activate
            ActiveCell.Value = temp(1, i)
            ActiveCell.NumberFormat = "mm/dd/yyyy"
            ActiveCell.Offset(0, 1).Value = ReviewerName(23)
            Select Case temp(3, i)
                Case Is = 1
                    ActiveCell.Offset(0, 2).Value = "ID"
                Case Is = 2
                    ActiveCell.Offset(0, 2).Value = "Assay"
                Case Is = 3
                    ActiveCell.Offset(0, 2).Value = "Potency"
                Case Is = 4
                    ActiveCell.Offset(0, 2).Value = "Impurity"
                Case Is = 5
                    ActiveCell.Offset(0, 2).Value = "Impurity/Potency"
            End Select
            ActiveCell.Offset(0, 3).Value = temp(4, i)
            ActiveCell.Offset(0, 4).Value = temp(5, i)
        Case Is = 24
            Worksheets(ReviewerName(24)).Activate
            NextRow = Worksheets(ReviewerName(24)).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
            Worksheets(ReviewerName(24)).Cells(NextRow, 1).Activate
            ActiveCell.Value = temp(1, i)
            ActiveCell.NumberFormat = "mm/dd/yyyy"
            ActiveCell.Offset(0, 1).Value = ReviewerName(24)
            Select Case temp(3, i)
                Case Is = 1
                    ActiveCell.Offset(0, 2).Value = "ID"
                Case Is = 2
                    ActiveCell.Offset(0, 2).Value = "Assay"
                Case Is = 3
                    ActiveCell.Offset(0, 2).Value = "Potency"
                Case Is = 4
                    ActiveCell.Offset(0, 2).Value = "Impurity"
                Case Is = 5
                    ActiveCell.Offset(0, 2).Value = "Impurity/Potency"
            End Select
            ActiveCell.Offset(0, 3).Value = temp(4, i)
            ActiveCell.Offset(0, 4).Value = temp(5, i)
        Case Is = 25
            Worksheets(ReviewerName(25)).Activate
            NextRow = Worksheets(ReviewerName(25)).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
            Worksheets(ReviewerName(25)).Cells(NextRow, 1).Activate
            ActiveCell.Value = temp(1, i)
            ActiveCell.NumberFormat = "mm/dd/yyyy"
            ActiveCell.Offset(0, 1).Value = ReviewerName(25)
            Select Case temp(3, i)
                Case Is = 1
                    ActiveCell.Offset(0, 2).Value = "ID"
                Case Is = 2
                    ActiveCell.Offset(0, 2).Value = "Assay"
                Case Is = 3
                    ActiveCell.Offset(0, 2).Value = "Potency"
                Case Is = 4
                    ActiveCell.Offset(0, 2).Value = "Impurity"
                Case Is = 5
                    ActiveCell.Offset(0, 2).Value = "Impurity/Potency"
            End Select
            ActiveCell.Offset(0, 3).Value = temp(4, i)
            ActiveCell.Offset(0, 4).Value = temp(5, i)
        Case Is = 26
            Worksheets(ReviewerName(26)).Activate
            NextRow = Worksheets(ReviewerName(26)).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
            Worksheets(ReviewerName(26)).Cells(NextRow, 1).Activate
            ActiveCell.Value = temp(1, i)
            ActiveCell.NumberFormat = "mm/dd/yyyy"
            ActiveCell.Offset(0, 1).Value = ReviewerName(26)
            Select Case temp(3, i)
                Case Is = 1
                    ActiveCell.Offset(0, 2).Value = "ID"
                Case Is = 2
                    ActiveCell.Offset(0, 2).Value = "Assay"
                Case Is = 3
                    ActiveCell.Offset(0, 2).Value = "Potency"
                Case Is = 4
                    ActiveCell.Offset(0, 2).Value = "Impurity"
                Case Is = 5
                    ActiveCell.Offset(0, 2).Value = "Impurity/Potency"
            End Select
            ActiveCell.Offset(0, 3).Value = temp(4, i)
            ActiveCell.Offset(0, 4).Value = temp(5, i)
        Case Is = 27
            Worksheets(ReviewerName(27)).Activate
            NextRow = Worksheets(ReviewerName(27)).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
            Worksheets(ReviewerName(27)).Cells(NextRow, 1).Activate
            ActiveCell.Value = temp(1, i)
            ActiveCell.NumberFormat = "mm/dd/yyyy"
            ActiveCell.Offset(0, 1).Value = ReviewerName(27)
            Select Case temp(3, i)
                Case Is = 1
                    ActiveCell.Offset(0, 2).Value = "ID"
                Case Is = 2
                    ActiveCell.Offset(0, 2).Value = "Assay"
                Case Is = 3
                    ActiveCell.Offset(0, 2).Value = "Potency"
                Case Is = 4
                    ActiveCell.Offset(0, 2).Value = "Impurity"
                Case Is = 5
                    ActiveCell.Offset(0, 2).Value = "Impurity/Potency"
            End Select
            ActiveCell.Offset(0, 3).Value = temp(4, i)
            ActiveCell.Offset(0, 4).Value = temp(5, i)
        Case Is = 28
            Worksheets(ReviewerName(28)).Activate
            NextRow = Worksheets(ReviewerName(28)).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
            Worksheets(ReviewerName(28)).Cells(NextRow, 1).Activate
            ActiveCell.Value = temp(1, i)
            ActiveCell.NumberFormat = "mm/dd/yyyy"
            ActiveCell.Offset(0, 1).Value = ReviewerName(28)
            Select Case temp(3, i)
                Case Is = 1
                    ActiveCell.Offset(0, 2).Value = "ID"
                Case Is = 2
                    ActiveCell.Offset(0, 2).Value = "Assay"
                Case Is = 3
                    ActiveCell.Offset(0, 2).Value = "Potency"
                Case Is = 4
                    ActiveCell.Offset(0, 2).Value = "Impurity"
                Case Is = 5
                    ActiveCell.Offset(0, 2).Value = "Impurity/Potency"
            End Select
            ActiveCell.Offset(0, 3).Value = temp(4, i)
            ActiveCell.Offset(0, 4).Value = temp(5, i)
        Case Is = 29
            Worksheets(ReviewerName(29)).Activate
            NextRow = Worksheets(ReviewerName(29)).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
            Worksheets(ReviewerName(29)).Cells(NextRow, 1).Activate
            ActiveCell.Value = temp(1, i)
            ActiveCell.NumberFormat = "mm/dd/yyyy"
            ActiveCell.Offset(0, 1).Value = ReviewerName(29)
            Select Case temp(3, i)
                Case Is = 1
                    ActiveCell.Offset(0, 2).Value = "ID"
                Case Is = 2
                    ActiveCell.Offset(0, 2).Value = "Assay"
                Case Is = 3
                    ActiveCell.Offset(0, 2).Value = "Potency"
                Case Is = 4
                    ActiveCell.Offset(0, 2).Value = "Impurity"
                Case Is = 5
                    ActiveCell.Offset(0, 2).Value = "Impurity/Potency"
            End Select
            ActiveCell.Offset(0, 3).Value = temp(4, i)
            ActiveCell.Offset(0, 4).Value = temp(5, i)
    End Select
Next i
For i = 1 To ReviewerNum
    Worksheets(ReviewerName(i)).Activate
    NextRow = Worksheets(ReviewerName(i)).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
    Range("A2:A" & NextRow - 1).Sort key1:=Range("A1"), order1:=xlAscending
    Cells(NextRow, 1).Value = "Total"
    Cells(NextRow + 1, 1).Value = "Average"
    Cells(NextRow, 4).Value = WorksheetFunction.Sum(Range("D2:D" & NextRow - 1).Value)
    Cells(NextRow, 5).Value = WorksheetFunction.Sum(Range("E2:E" & NextRow - 1).Value)
    Cells(NextRow + 1, 4).Value = WorksheetFunction.Average(Range("D2:D" & NextRow - 1).Value)
    Cells(NextRow + 1, 5).Value = WorksheetFunction.Average(Range("E2:E" & NextRow - 1).Value)
Next i
End Sub
'----------------------------------------------------------------------
'  Period_End = DateSerial(Year_Num, Month_Num, Day_Num)
'    FirstWeekDay = Weekday(Period_End) + 10
'    week_num = WorksheetFunction.WeekNum(Period_End, FirstWeekDay)
'    Period_Begin = Period_End - 6
'--------------------------------------------------------------------
'Sub Macro1()
''
'' Macro1 Macro
''
'
''
'    With Selection.Validation
'        .Delete
'        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:= _
'        xlBetween, formula1:="1/1/2019", formula2:="1/31/2019"
'        .IgnoreBlank = True
'        .InCellDropdown = True
'        .InputTitle = ""
'        .ErrorTitle = ""
'        .InputMessage = ""
'        .ErrorMessage = ""
'        .ShowInput = True
'        .ShowError = True
'    End With
'End Sub


