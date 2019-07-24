Attribute VB_Name = "Data_Reviewer_ScoreBoard"
Public week_num As Integer
Sub reviewer_score()
'-----------------------------------------------------------------------------
'Prepare for data entry
'-----------------------------------------------------------------------------
    Dim date1 As Date
    Dim date2 As Date
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim btn As Button
    If week_num <> 0 Then
        GoTo CreateNewSheet
    Else
        week_num = InputBox("Please enter week number (1-52).")
    End If
CreateNewSheet:
    Set wb = ActiveWorkbook
    On Error Resume Next
    Set ws = wb.Sheets("Week_" & week_num)
    On Error GoTo 0
    If Not ws Is Nothing Then
        MsgBox "The Sheet called " & "Week_" & week_num & " already existed in the workbook.", vbExclamation, "Sheet Already Exists!"
        GoTo entry_prompt
    Else
        Set ws = wb.Sheets.Add(after:=wb.Sheets(wb.Sheets.Count))
        ws.name = "Week_" & week_num
    End If
    Cells(1, 1).Value = "Review Date"
    Cells(1, 2).Value = "Name"
    Cells(1, 3).Value = "Pot/Imp Assigned"
    Cells(1, 4).Value = "Pot/Imp with Error"
    Cells(1, 5).Value = "Pot/Imp Error"
    Cells(1, 6).Value = "Imp Assigned"
    Cells(1, 7).Value = "Imp with Error"
    Cells(1, 8).Value = "Imp Error"
    Cells(1, 9).Value = "Pot Assigned"
    Cells(1, 10).Value = "Pot with Error"
    Cells(1, 11).Value = "Pot Error"
    Cells(1, 12).Value = "Assay Assigned"
    Cells(1, 13).Value = "Assay with Error"
    Cells(1, 14).Value = "Assay Error"
    Cells(1, 15).Value = "ID Assigned"
    Cells(1, 16).Value = "ID with Error"
    Cells(1, 17).Value = "ID Error"
    Cells(1, 18).Value = "Penalty"
    Cells(1, 19).Value = "Final Score"
    Range("B2").Select
    With Selection.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, formula1:="=Names!$A$1:$A$27"
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
    Range("C2:L2").Select
    With Selection.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, formula1:="=Names!$D$1:$D$10"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Selection.AutoFill Destination:=Range("C2:L1048576"), Type:=xlFillDefault
    Range("C2:L2").End(xlDown).Select
entry_prompt:
    Set btn = ActiveSheet.Buttons.Add(Range("U1").Left, 0, 120, 25)
    btn.Select
    With Selection
    .OnAction = "Compute"
    .Caption = "Calculate"
    .Font.Bold = True
    End With
    Worksheets("Week_" & week_num).Activate
    Cells(1, 1).Select
MsgBox ("Enter Data in columns A-L. Click the 'Calculate' button to compute penalty and final score.")
    Set btn2 = ActiveSheet.Buttons.Add(Range("U5").Left, 30, 120, 25)
    btn2.Select
    With Selection
    .OnAction = "Gen_report"
    .Caption = "Report"
    .Font.Bold = True
    End With
MsgBox ("Click Report to generate monthly report.")
date1 = DateSerial(year(Date), 1, (week_num - 1) * 7 + 1)
date2 = date1 + 6

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
Cells(2, 1).Activate
End Sub
Sub Compute()
'--------------------------------------------------------------------------------
'variables for store input
'--------------------------------------------------------------------------------
    Dim entry_date() As Date
    Dim reviewer() As String
    Dim pot_imp_lot() As Integer
    Dim pot_imp_err_lot() As Integer
    Dim pot_imp_err() As Integer
    Dim pot_imp_pen() As Double
    Dim imp_lot() As Integer
    Dim imp_err_lot() As Integer
    Dim imp_err() As Integer
    Dim imp_pen() As Double
    Dim pot_lot() As Integer
    Dim pot_err_lot() As Integer
    Dim pot_err() As Integer
    Dim pot_pen() As Double
    Dim assay_lot() As Integer
    Dim assay_err_lot() As Integer
    Dim assay_err() As Integer
    Dim assay_pen() As Double
    Dim id_lot() As Integer
    Dim id_err_lot() As Integer
    Dim id_err() As Integer
    Dim id_pen() As Double
    Dim penalty() As Double
    Dim score() As Double
    Dim record_num As Long
    Dim i As Integer
    Dim j As Integer
'------------------------------------------------------------------------------
'variables for specifying report
'------------------------------------------------------------------------------
    Dim report_type As Integer
    Dim report_year As Integer
    Dim report_month As Integer
    Dim report_quarter As Integer
    Dim report_week As Integer
'------------------------------------------------------------------------------
'variables for calculate and store scores for each reviewer
'------------------------------------------------------------------------------
    Dim reviewer_num As Integer
    Dim reivew_count() As Integer
    Dim review_date() As Date
    Dim num_review_lot() As Integer
    Dim num_review_assay() As Integer
    Dim num_review_pot() As Integer
    Dim num_review_imp() As Integer
    Dim num_review_id() As Integer
    Dim review_score() As Long
    Dim review_penal() As Long
    Dim sht As Worksheet

If week_num = 0 Then
    week_num = InputBox("Which week do you want to calculate?", "Enter Week Number")
    On Error Resume Next
        Set sht = Worksheets("Week_" & week_num)
    On Error GoTo ErrHandler
ErrHandler:
    MsgBox ("Specified week does not exist, creating one.")
    reviewer_score
Else
    Worksheets("Week_" & week_num).Activate
End If
Cells(1, 1).Activate
record_num = ActiveSheet.UsedRange.Rows.Count
    ReDim entry_date(record_num) As Date
    ReDim reviewer(record_num) As String
    ReDim pot_imp_lot(record_num) As Integer
    ReDim pot_imp_err_lot(record_num) As Integer
    ReDim pot_imp_err(record_num) As Integer
    ReDim pot_imp_pen(record_num) As Double
    ReDim imp_lot(record_num) As Integer
    ReDim imp_err_lot(record_num) As Integer
    ReDim imp_err(record_num) As Integer
    ReDim imp_pen(record_num) As Double
    ReDim pot_lot(record_num) As Integer
    ReDim pot_err_lot(record_num) As Integer
    ReDim pot_err(record_num) As Integer
    ReDim pot_pen(record_num) As Double
    ReDim assay_lot(record_num) As Integer
    ReDim assay_err_lot(record_num) As Integer
    ReDim assay_err(record_num) As Integer
    ReDim assay_pen(record_num) As Double
    ReDim id_lot(record_num) As Integer
    ReDim id_err_lot(record_num) As Integer
    ReDim id_err(record_num) As Integer
    ReDim id_pen(record_num) As Double
    ReDim penalty(record_num) As Double
    ReDim score(record_num) As Double
    Cells(2, 1).Activate
    For i = 2 To record_num
      entry_date(i) = ActiveCell.Value
      reviewer(i) = ActiveCell.Offset(0, 1).Value
      pot_imp_lot(i) = ActiveCell.Offset(0, 2).Value
      pot_imp_err_lot(i) = ActiveCell.Offset(0, 3).Value
      pot_imp_err(i) = ActiveCell.Offset(0, 4).Value
      imp_lot(i) = ActiveCell.Offset(0, 5).Value
      imp_err_lot(i) = ActiveCell.Offset(0, 6).Value
      imp_err(i) = ActiveCell.Offset(0, 7).Value
      pot_lot(i) = ActiveCell.Offset(0, 8).Value
      pot_err_lot(i) = ActiveCell.Offset(0, 9).Value
      pot_err(i) = ActiveCell.Offset(0, 10).Value
      assay_lot(i) = ActiveCell.Offset(0, 11).Value
      assay_err_lot(i) = ActiveCell.Offset(0, 12).Value
      assay_err(i) = ActiveCell.Offset(0, 13).Value
      id_lot(i) = ActiveCell.Offset(0, 14).Value
      id_err_lot(i) = ActiveCell.Offset(0, 15).Value
      id_err(i) = ActiveCell.Offset(0, 16).Value
      If pot_imp_lot(i) = 0 Then
        pot_imp_pen(i) = 0
      Else
        pot_imp_pen(i) = pot_imp_err_lot(i) * pot_imp_err(i) / pot_imp_lot(i) * 5
      End If
      If imp_lot(i) = 0 Then
        imp_pen(i) = 0
      Else
        imp_pen(i) = imp_err_lot(i) * imp_err(i) / imp_lot(i) * 4
      End If
      If pot_lot(i) = 0 Then
        pot_pen(i) = 0
      Else
        pot_pen(i) = pot_err_lot(i) * pot_err(i) / pot_lot(i) * 3
      End If
      If assay_lot(i) = 0 Then
        assay_pen(i) = 0
      Else
        assay_pen(i) = assay_err_lot(i) * assay_err(i) / assay_lot(i) * 2
      End If
      If id_lot(i) = 0 Then
        id_pen(i) = 0
      Else
        id_pen(i) = id_err_lot(i) * id_err(i) / id_lot(i) * 1
      End If
      penalty(i) = pot_imp_pen(i) + imp_pen(i) + pot_pen(i) + assay_pen(i) + id_pen(i)
      score(i) = 100 - penalty(i)
      ActiveCell.Offset(0, 17).Value = penalty(i)
      ActiveCell.Offset(0, 18).Value = score(i)
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
    Dim report_start_date As Date
    Dim report_end_date As Date
    Dim month_start As Date
    Dim month_end As Date
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
    report_start_date = month_start - (Weekday(month_start) - 1)
    report_end_date = month_end - Weekday(month_end) + 7
    
    MsgBox ("Processing monthly report.")
End Sub
'----------------------------------------------------------------------
'  Period_End = DateSerial(Year_Num, Month_Num, Day_Num)
'    FirstWeekDay = Weekday(Period_End) + 10
'    week_num = WorksheetFunction.WeekNum(Period_End, FirstWeekDay)
'    Period_Begin = Period_End - 6
'--------------------------------------------------------------------
