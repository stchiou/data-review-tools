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
        xlBetween, formula1:="=Names!$A$1:$A$30"
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
      penalty(i) = (LotError(i) * ErrorNum(i)) / (assignment(i) * LotAssigned(i))
      score(i) = 100 - penalty(i)
      ActiveCell.Offset(0, 6).Value = penalty(i)
      ActiveCell.Offset(0, 7).Value = score(i)
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
ReDim temp(4, ReportRecNum) As Double
Curr_Rec = 0
For i = report_start_week To report_end_week
    Worksheets(SheetName(i)).Activate
    For j = 2 To rowNum(i)
        Curr_Rec = Curr_Rec + j - 1
        Cells(j, 2).Activate
        Select Case Cells(j, 2).Value
            Case Is = "Alam, Nuzhat P"
                temp(1, Curr_Rec) = 1
            Case Is = "Barnes, Michelle"
                temp(1, Curr_Rec) = 2
            Case Is = "Batts, George III"
                temp(1, Curr_Rec) = 3
            Case Is = "Beckwith, Catherine"
                temp(1, Curr_Rec) = 4
            Case Is = "Blair, Kenneth John"
                temp(1, Curr_Rec) = 5
            Case Is = "Bomboy, Dustin Shaun"
                temp(1, Curr_Rec) = 6
            Case Is = "Borrero López, Francheska"
                temp(1, Curr_Rec) = 7
            Case Is = "Cintron Barreto, Derickniel"
                temp(1, Curr_Rec) = 8
            Case Is = "Clark, Antonio"
                temp(1, Curr_Rec) = 9
            Case Is = "Clark, Janneth Lucia"
                temp(1, Curr_Rec) = 10
            Case Is = "Dudley, Jocelyn Imi"
                temp(1, Curr_Rec) = 11
            Case Is = "Ghahra, Parvaneh"
                temp(1, Curr_Rec) = 12
            Case Is = "Gray, Jason L."
                temp(1, Curr_Rec) = 13
            Case Is = "HARRISON, MARY"
                temp(1, Curr_Rec) = 14
            Case Is = "Lash, Tanya"
                temp(1, Curr_Rec) = 15
            Case Is = "Lee, Trecia"
                temp(1, Curr_Rec) = 16
            Case Is = "McRae, Tangelo"
                temp(1, Curr_Rec) = 17
            Case Is = "McBean, Coray"
                temp(1, Curr_Rec) = 18
            Case Is = "Nash, Shalena"
                temp(1, Curr_Rec) = 19
            Case Is = "Obdens, Aaron Benjamin"
                temp(1, Curr_Rec) = 20
            Case Is = "Polashuk, Michael"
                temp(1, Curr_Rec) = 21
            Case Is = "Riley, Lakesha"
                temp(1, Curr_Rec) = 22
            Case Is = "Silver, Carla Marie"
                temp(1, Curr_Rec) = 23
            Case Is = "Smith, Carlton E"
                temp(1, Curr_Rec) = 24
            Case Is = "Springer-Dickson, Sherlene"
                temp(1, Curr_Rec) = 25
            Case Is = "Tummala, Lok"
                temp(1, Curr_Rec) = 26
            Case Is = "Vines, Vernon"
                temp(1, Curr_Rec) = 27
            Case Is = "Wynn, Jason L."
                temp(1, Curr_Rec) = 28
            Case Is = "Zimmerman-Ford, Gisela Z"
                temp(1, Curr_Rec) = 29
        End Select
        Select Case ActiveCell.Offset(0, 1).Value
            Case Is = "Impurity/Potency"
                temp(2, Curr_Rec) = 5
            Case Is = "Impurity"
                temp(2, Curr_Rec) = 4
            Case Is = "Potency"
                temp(2, Curr_Rec) = 3
            Case Is = "Assay"
                temp(2, Curr_Rec) = 2
            Case Is = "ID"
                temp(2, Curr_Rec) = 1
        End Select
        temp(3, Curr_Rec) = ActiveCell.Offset(0, 2).Value
        temp(4, Curr_Rec) = ActiveCell.Offset(0, 6).Value
    Next j
Next i
MsgBox ("Data loaded")
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


