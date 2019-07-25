Attribute VB_Name = "Data_Reviewer_ScoreBoard"
Public week_num As Integer
Public yr As Integer
Public wn As String
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
    
'    If week_num <> 0 Then
'        GoTo CreateNewSheet
'    Else
'        yr = InputBox("Please enter the year of the records.")
'        week_num = InputBox("Please enter week number (1-52).")
'    End If
'CreateNewSheet:
'    Set wb = ActiveWorkbook
'    On Error Resume Next
'    Set ws = wb.Sheets("Week_" & week_num & "_" & yr)
'    On Error GoTo 0
'    If Not ws Is Nothing Then
'        MsgBox "The Sheet called " & "Week_" & week_num & "_" & yr & " already existed in this workbook.", vbExclamation, "Sheet Already Exists!"
'        GoTo entry_prompt
'    Else
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
    Cells(1, 7).Value = "Penalty"
    Cells(1, 8).Value = "Score"
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
MsgBox ("Enter Data in columns A-L. Click the 'Calculate' button to compute penalty and final score.")
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
    .OnAction = "UpDate_Record"
    .Caption = "Update Scores"
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
record_num = ActiveSheet.UsedRange.Rows.Count
    ReDim entry_date(record_num) As Date
    ReDim reviewer(record_num) As String
    ReDim assignment(record_num) As String
    ReDim LotAssigned(record_num) As Integer
    ReDim LotError(record_num) As Integer
    ReDim penalty(record_num) As Double
    ReDim score(record_num) As Double
    Cells(2, 1).Activate
    For i = 2 To record_num
      entry_date(i) = ActiveCell.Value
      reviewer(i) = ActiveCell.Offset(0, 1).Value
      assignment(i) = ActiveCell.Offset(0, 2).Value
      Select Case assignment(i)
        Case Is = "Impurity/Potency"
        Case Is = "Impurity"
        Case Is = "Potency"
        Case Is = "Assay"
        Case Is = "ID"
        
      End Select
      LotAssigned(i) = ActiveCell.Offset(0, 3).Value
      LotError(i) = ActiveCell.Offset(0, 4).Value
      penalty (i)
'      pot_imp_lot(i) = ActiveCell.Offset(0, 2).Value
'      pot_imp_err_lot(i) = ActiveCell.Offset(0, 3).Value
'      pot_imp_err(i) = ActiveCell.Offset(0, 4).Value
'      imp_lot(i) = ActiveCell.Offset(0, 5).Value
'      imp_err_lot(i) = ActiveCell.Offset(0, 6).Value
'      imp_err(i) = ActiveCell.Offset(0, 7).Value
'      pot_lot(i) = ActiveCell.Offset(0, 8).Value
'      pot_err_lot(i) = ActiveCell.Offset(0, 9).Value
'      pot_err(i) = ActiveCell.Offset(0, 10).Value
'      assay_lot(i) = ActiveCell.Offset(0, 11).Value
'      assay_err_lot(i) = ActiveCell.Offset(0, 12).Value
'      assay_err(i) = ActiveCell.Offset(0, 13).Value
'      id_lot(i) = ActiveCell.Offset(0, 14).Value
'      id_err_lot(i) = ActiveCell.Offset(0, 15).Value
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
    Dim report_start_week As Integer
    Dim report_end_week As Integer
    Dim SheetNum As Integer
    Dim rowNum() As Integer
    Dim ReportRecNum As Integer
    Dim summary() As Double
    Dim month_start As Date
    Dim month_end As Date
    Dim i As Integer
    Dim j As Integer
    Dim temp() As Double
    
'---------------------------------------------------------
'Array dimension
'----------------

'---------------------------------------------------------
    ReDim summary(26, 2) As Double
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
    ReDim rowNum(report_end_week) As Integer
    ReportRecNum = 0
    For i = report_start_week To report_end_week
        rowNum(i) = Worksheets("Week_" & i).UsedRange.Rows.Count
        ReportRecNum = ReportRecNum + rowNum(i) - 1
    Next i
 
    MsgBox ("Processing monthly report.")
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


