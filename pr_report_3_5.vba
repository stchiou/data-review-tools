Attribute VB_Name = "PR_Status_Report_v3"
Sub PR_Report()
'-----------------------------------------------------------------
'Macro for computing weekly PR Status
'Sean Chiou, version 3.5, 04/03/2019
'-----------------------------------------------------------------
'Items required:
'1. total opein-categorized by type of records
'2. closed last week
'3. aged > 30 days (bar chart, including data from previous 5 weeks, categorized by types:ER, QAR, LIR, RACAC, INC)
'4. aging up (age > 23 days)
'5. committed to close this week
'6. aged that will close
'7. PRs Opened by week LIR, RAAC, QAR, ER)
'8. PRs opened by month (LIR, RAAC, QAR, ER)
'9. PRs Opened by week and by month (LIR, RAAC, QAR, ER)
'10. PRs by writer
'11. PRs opened (CQ vs IM)

'-------------------------------------------------------------------------------------------------------------------
'Features:
'1. Combine output records with corresponding short description
'2. Computes age of the records
'3. Computes stage of the records based on age
'4. Generate reports
'------------------------------------------------------------------------------------------------------------------
Dim File_1 As String
Dim Report_Type As Integer
Dim Week_Num As Long
Dim Month_Num As Integer
Dim Quarter_Num As Integer
Dim Year_Num As Integer
Dim Day_Num As Integer
Dim r_y As Integer
Dim r_m As Integer
Dim r_d As Integer
Dim Period_End As Date
Dim Period_Begin As Date
Dim Sub_Per_Start() As Date
Dim Sub_Per_End() As Date
Dim UnitInPeriod As Integer
Dim Record_Num As Long
Dim FirstWeekDay As Integer
'-------------------------------------------------------
'Arrays for fields in raw data
'-------------------------------------------------------
Dim pr_id() As String
Dim title_short_description() As String
Dim responsible_person() As String
Dim record_type() As String
Dim investigation_type() As String
Dim related_records() As String
Dim event_code() As String
Dim qar_required() As String
Dim special_or_common_cuase() As String
Dim capa_effectiveness_bsc_metric() As String
Dim date_open() As Date
Dim discovery_date() As Date
Dim date_closed() As Date
Dim due_date() As Date
Dim original_due_date() As Date
Dim number_of_approved_extensions() As Integer
Dim qa_final_app_on() As Date
Dim site_qa_approval_on() As Date
Dim material_involved() As String
Dim bu_area() As String
Dim operation() As String
Dim test_description() As String
Dim other_test_description() As String
Dim procedure_method() As String
Dim product_families() As String
Dim product_names() As String
Dim initial_inv_analyst() As String
Dim hp_root_cause_categ_1() As String
Dim hp_root_cause_categ_2() As String
Dim hp_root_cause_categ_3() As String
Dim root_cause_cat_1() As String
Dim root_cause_cat_2() As String
Dim root_cause_cat_3() As String
Dim recom_diposition_comments() As String
Dim final_comments() As String
Dim assignable_cause_class() As String
Dim assignable_cause() As String
Dim supplier_name_lot_no() As String
Dim area_discovered() As String
Dim areas_affected() As String
Dim analyst_personnel_sub_category() As String
Dim pr_state() As String
Dim reason_for_investigating() As String
Dim idc_level_1() As String
Dim idc_level_2() As String
Dim idc_level_3() As String
Dim root_cause_lev_1() As String
Dim root_cause_lev_2() As String
Dim root_cause_lev_3() As String
'-----------------------------------------------------------------
'Array/variables for processing open records
'-----------------------------------------------------------------
Dim OpenRecNum As Integer
Dim Open_Index() As Integer
Dim OpenList() As Integer
Dim OpenList_Pos As Integer
Dim OpenAge() As Integer
Dim OpenStage() As Integer
Dim OpenRecType() As Integer
Dim OpenRecCount() As Integer
Dim OpenArea() As Integer
Dim OpenRec() As String
'---------------------------------------------------------------
'Arrays/variables for processing closed records
'----------------------------------------------------------------
Dim ClosedRecNum As Integer
Dim Closed_Index() As Integer
Dim ClosedList() As Integer
Dim ClosedList_Pos As Integer
Dim ClosedStage() As Integer
Dim ClosedRecType() As Integer
Dim ClosedRecCount() As Integer
Dim ClosedArea() As Integer
Dim ClosedRec() As String
'----------------------------------------------------------------
'Arrays/variables for processing new records
'----------------------------------------------------------------
Dim NewRecNum As Integer
Dim NewCount() As Integer
Dim New_Index() As Integer
Dim NewList() As Integer
Dim NewRecType() As Integer
Dim NewList_Pos As Integer
Dim NewRec() As String
'----------------------------------------------------------------
'Arrays/variables for processing cancelled records
'----------------------------------------------------------------
Dim CancelRecNum As Integer
Dim CancelCount() As Integer
Dim Cancel_Index() As Integer
Dim CancelList() As Integer
Dim CancelList_Pos As Integer
Dim CancelRecType() As Integer
Dim CancelRec() As String
'----------------------------------------------------------------
'Arrays/variables for processing report
'----------------------------------------------------------------
Dim Rep_Headers(32) As String
Dim ReplCol As Long
Dim ReplRow As Long
Dim week_range As Long
Dim month_range As Long
Dim quarter_range As Long
Dim year_range As Long
Dim Range_Weekly_Rec() As Long
Dim Range_Monthly_Rec() As Long
Dim Range_Quarterly_Rec() As Long
Dim Range_Annual_Rec() As Long
Dim New_record() As String
Dim DivCount() As Integer
Dim committed() As String
'-----------------------------------------------------------------
Dim ReportSheet_Name As String
Dim DataSheet_Name As String
Dim ChartSheet_Name As String
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim age As Integer
Dim stage As Integer
Dim address_1 As String
Dim address_2 As String
'---------------------------------------------------------------------------------
'Capture File Names and Path of Data files
'---------------------------------------------------------------------------------
Input_report_type:
    Report_Type = InputBox("Which type of report you want to generate?" _
        & vbCr & "1. Weekly" _
        & vbCr & "2. Monthly" _
        & vbCr & "3. Quarterly" _
        & vbCr & "4. Annually" _
        & vbCr & "5. Arbitary Range")
    If Report_Type = 1 Then
        GoTo Input_week_parameters:
    Else    'Report_type=1
        If Report_Type = 2 Then
            GoTo Input_month_parameters:
        Else    'Report_type=2
            If Report_Type = 3 Then
                GoTo Input_quarter_parameters:
            Else 'Report_type=3
                If Report_Type = 4 Then
                    GoTo Input_year_parameters:
                Else 'report_type=4
                    If Report_Type = 5 Then
                        GoTo Input_range_parameters:
                    Else 'Report_Type=5
                        GoTo Input_report_type:
                    End If 'Report_Type=5
                End If 'report_type=4
            End If 'Report_type=3
        End If  'Report_type=2
    End If  'Report_type=1
Input_week_parameters:
    Year_Num = InputBox("Input numeric value of the Year for the Report", "YEAR NUMBER")
    Month_Num = InputBox("Input numeric value the Month for the Report" _
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
        & vbCr & "12. December", "MONTH NUMBER")
    Day_Num = InputBox("Input numeric value of the day of the month for the Report", "DAY NUMBER")
    Period_End = DateSerial(Year_Num, Month_Num, Day_Num)
    FirstWeekDay = Weekday(Period_End) + 10
    Week_Num = WorksheetFunction.WeekNum(Period_End, FirstWeekDay)
    Period_Begin = Period_End - 6
    GoTo Input_data_file:
Input_month_parameters:
    Year_Num = InputBox("Input numeric value of the Year for the Report", "YEAR NUMBER")
    Month_Num = InputBox("Input numeric value the Month for the Report" _
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
        & vbCr & "12. December", "MONTH NUMBER")
    Period_Begin = DateSerial(Year_Num, Month_Num, 1)
    Period_End = DateSerial(Year_Num, Month_Num + 1, 0)
    
    GoTo Input_data_file:
Input_quarter_parameters:
    Year_Num = InputBox("Input numeric value of the Year for the Report", "YEAR NUMBER")
    Quarter_Num = InputBox("Input numeric value of the quarter for the report", "QUARTER NUMBER")
    Period_Begin = DateSerial(Year_Num, (Quarter_Num - 1) * 3 + 1, 1)
    Period_End = DateSerial(Year_Num, Quarter_Num * 3 + 1, 0)
    GoTo Input_data_file:
Input_year_parameters:
    Year_Num = InputBox("Input numeric value of Year of the Report", "YEAR NUMBER")
    Period_Begin = DateSerial(Year_Num, 1, 1)
    Period_End = DateSerial(Year_Num, 12, 31)
    GoTo Input_data_file:
Input_range_parameters:
    Year_Num = InputBox("Input numeric value of Year that report starts", "START YEAR")
    Month_Num = InputBox("Input numeric value of month that report starts" _
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
        & vbCr & "12. December", "START MONTH")
    Day_Num = InputBox("Input numeric value of day of the month that report starts", "START DAY")
    r_y = InputBox("Input numeric value of the Year that report ends", "END YEAR")
    r_m = InputBox("Input numeric value of the month that report ends" _
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
        & vbCr & "12. December", "END MONTH")
    r_d = InputBox("Input numeric value of day of the month that report ends", "END DAY")
    Period_Begin = DateSerial(Year_Num, Month_Num, Day_Num)
    Period_End = DateSerial(r_y, r_m, r_d)
    GoTo Input_data_file:
Input_data_file:
    File_1 = Application.GetOpenFilename _
        (Title:="Data File", _
        filefilter:="CSV (Comma delimited) (*.csv),*.csv")
    If MsgBox("File contains records to be processed is " & File_1 & ". Is this correct?", vbYesNo) = vbNo Then
        GoTo Input_data_file:
    Else
    End If
Verification:
    If MsgBox("The range of report is from " & Period_Begin & " to " & Period_End & ". Is this correct?", vbYesNo) = vbNo Then
        GoTo Input_report_type:
    Else
    End If
DataSheet_Name = Mid(File_1, InStrRev(File_1, "\") + 1, (Len(File_1) - InStrRev(File_1, "\") - 4))
Window_1 = DataSheet_Name & ".csv"
'-------------------------------------------------------------------------------
'Calculate Record Number and redeclare array for raw data
'-------------------------------------------------------------------------------
Workbooks.OpenText Filename:=File_1, local:=True
'Workbooks.Open Filename:=File_2, local:=True
Windows(Window_1).Activate
Record_Num = Cells(1, 1).End(xlDown).Row
ReDim pr_id(Record_Num)
ReDim title_short_description(Record_Num)
ReDim responsible_person(Record_Num)
ReDim record_type(Record_Num)
ReDim investigation_type(Record_Num)
ReDim related_records(Record_Num)
ReDim qar_required(Record_Num)
ReDim event_code(Record_Num)
ReDim special_or_common_cuase(Record_Num)
ReDim capa_effectiveness_bsc_metric(Record_Num)
ReDim date_open(Record_Num)
ReDim discovery_date(Record_Num)
ReDim date_closed(Record_Num)
ReDim due_date(Record_Num)
ReDim original_due_date(Record_Num)
ReDim number_of_approved_extensions(Record_Num)
ReDim qa_final_app_on(Record_Num)
ReDim site_qa_approval_on(Record_Num)
ReDim material_involved(Record_Num)
ReDim bu_area(Record_Num)
ReDim operation(Record_Num)
ReDim test_description(Record_Num)
ReDim other_test_description(Record_Num)
ReDim procedure_method(Record_Num)
ReDim product_families(Record_Num)
ReDim product_names(Record_Num)
ReDim initial_inv_analyst(Record_Num)
ReDim hp_root_cause_categ_1(Record_Num)
ReDim hp_root_cause_categ_2(Record_Num)
ReDim hp_root_cause_categ_3(Record_Num)
ReDim root_cause_cat_1(Record_Num)
ReDim root_cause_cat_2(Record_Num)
ReDim root_cause_cat_3(Record_Num)
ReDim recom_diposition_comments(Record_Num)
ReDim final_comments(Record_Num)
ReDim assignable_cause_class(Record_Num)
ReDim assignable_cause(Record_Num)
ReDim supplier_name_lot_no(Record_Num)
ReDim area_discovered(Record_Num)
ReDim areas_affected(Record_Num)
ReDim analyst_personnel_sub_category(Record_Num)
ReDim pr_state(Record_Num)
ReDim reason_for_investigation(Record_Num)
ReDim idc_level_1(Record_Num)
ReDim idc_level_2(Record_Num)
ReDim idc_level_3(Record_Num)
ReDim root_cause_lev_1(Record_Num)
ReDim root_cause_lev_2(Record_Num)
ReDim root_cause_lev_3(Record_Num)
For i = 2 To Record_Num
    Cells(i, 1).Activate
    pr_id(i) = ActiveCell.Value
    title_short_description(i) = ActiveCell.Offset(0, 1).Value
    responsible_person(i) = ActiveCell.Offset(0, 2).Value
    record_type(i) = ActiveCell.Offset(0, 3).Value
    investigation_type(i) = ActiveCell.Offset(0, 4).Value
    related_records(i) = ActiveCell.Offset(0, 5).Value
    qar_required(i) = ActiveCell.Offset(0, 6).Value
    special_or_common_cuase(i) = ActiveCell.Offset(0, 7).Value
    event_code(i) = ActiveCell.Offset(0, 8).Value
    capa_effectiveness_bsc_metric(i) = ActiveCell.Offset(0, 9).Value
    date_open(i) = ActiveCell.Offset(0, 10).Value
    discovery_date(i) = ActiveCell.Offset(0, 11).Value
    date_closed(i) = ActiveCell.Offset(0, 12).Value
    due_date(i) = ActiveCell.Offset(0, 13).Value
    original_due_date(i) = ActiveCell.Offset(0, 14).Value
    number_of_approved_extensions(i) = ActiveCell.Offset(0, 15).Value
    qa_final_app_on(i) = ActiveCell.Offset(0, 16).Value
    site_qa_approval_on(i) = ActiveCell.Offset(0, 17).Value
    material_involved(i) = ActiveCell.Offset(0, 18).Value
    bu_area(i) = ActiveCell.Offset(0, 19).Value
    operation(i) = ActiveCell.Offset(0, 20).Value
    test_description(i) = ActiveCell.Offset(0, 21).Value
    other_test_description(i) = ActiveCell.Offset(0, 22).Value
    procedure_method(i) = ActiveCell.Offset(0, 23).Value
    product_families(i) = ActiveCell.Offset(0, 24).Value
    product_names(i) = ActiveCell.Offset(0, 25).Value
    initial_inv_analyst(i) = ActiveCell.Offset(0, 26).Value
    hp_root_cause_categ_1(i) = ActiveCell.Offset(0, 27).Value
    hp_root_cause_categ_2(i) = ActiveCell.Offset(0, 28).Value
    hp_root_cause_categ_3(i) = ActiveCell.Offset(0, 29).Value
    root_cause_cat_1(i) = ActiveCell.Offset(0, 30).Value
    root_cause_cat_2(i) = ActiveCell.Offset(0, 31).Value
    root_cause_cat_3(i) = ActiveCell.Offset(0, 32).Value
    recom_diposition_comments(i) = ActiveCell.Offset(0, 33).Value
    final_comments(i) = ActiveCell.Offset(0, 34).Value
    assignable_cause_class(i) = ActiveCell.Offset(0, 35).Value
    assignable_cause(i) = ActiveCell.Offset(0, 36).Value
    supplier_name_lot_no(i) = ActiveCell.Offset(0, 37).Value
    area_discovered(i) = ActiveCell.Offset(0, 38).Value
    areas_affected(i) = ActiveCell.Offset(0, 39).Value
    analyst_personnel_sub_category(i) = ActiveCell.Offset(0, 40).Value
    pr_state(i) = ActiveCell.Offset(0, 41).Value
    reason_for_investigation(i) = ActiveCell.Offset(0, 42).Value
    idc_level_1(i) = ActiveCell.Offset(0, 43).Value
    idc_level_2(i) = ActiveCell.Offset(0, 44).Value
    idc_level_3(i) = ActiveCell.Offset(0, 45).Value
    root_cause_lev_1(i) = ActiveCell.Offset(0, 46).Value
    root_cause_lev_2(i) = ActiveCell.Offset(0, 47).Value
    root_cause_lev_3(i) = ActiveCell.Offset(0, 48).Value
Next i
'------------------------------------------------------------------------------
'Count Number of Opened Record
'------------------------------------------------------------------------------
OpenRecNum = 0
ReDim Open_Index(Record_Num)
For i = 2 To Record_Num
If pr_state(i) = "Cancelled" Then
        OpenRecNum = OpenRecNum
        Open_Index(i) = 0
Else
    If date_open(i) > Period_End + 1 Then
        OpenRecNum = OpenRecNum
        Open_Index(i) = 0
    Else
        If site_qa_approval_on(i) = 0 Then
            If qa_final_app_on(i) = 0 Then
                OpenRecNum = OpenRecNum + 1
                Open_Index(i) = i
            Else
                If qa_final_app_on(i) > Period_End + 1 Then
                    OpenRecNum = OpenRecNum + 1
                    Open_Index(i) = i
                Else
                    OpenRecNum = OpenRecNum
                    Open_Index(i) = 0
                End If
            End If
        Else
            If site_qa_approval_on(i) > Period_End + 1 Then
                OpenRecNum = OpenRecNum + 1
                Open_Index(i) = i
            Else
                OpenRecNum = OpenRecNum
                Open_Index(i) = 0
            End If
        End If
    End If
End If
Next i
'-------------------------------------------------------------------------------
'Fill the list of Opened Records with index numbers of the whole data set
'-------------------------------------------------------------------------------
ReDim OpenList(OpenRecNum)
OpenList_Pos = 1
For i = 1 To Record_Num
        If Open_Index(i) <> 0 Then
            OpenList(OpenList_Pos) = Open_Index(i)
            OpenList_Pos = OpenList_Pos + 1
        Else
        End If
Next i
'---------------------------------------------------------------------------------
'Calculate Age , Stage and Type of Opened Records
'---------------------------------------------------------------------------------
ReDim OpenAge(OpenRecNum)
ReDim OpenStage(OpenRecNum)
ReDim OpenRecType(OpenRecNum)
For i = 1 To OpenRecNum
    OpenAge(i) = Period_End - discovery_date(OpenList(i))
    If OpenAge(i) < 23 Then
        OpenStage(i) = 0
    Else
        If OpenAge(i) < 30 Then
            OpenStage(i) = 1
        Else
            If OpenAge(i) < 60 Then
                OpenStage(i) = 2
            Else
                If OpenAge(i) < 90 Then
                    OpenStage(i) = 3
                Else
                    If OpenAge(i) < 120 Then
                        OpenStage(i) = 4
                    Else
                        If OpenAge(i) < 150 Then
                            OpenStage(i) = 5
                        Else
                            If OpenAge(i) < 180 Then
                                OpenStage(i) = 6
                            Else
                                OpenStage(i) = 7
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    Select Case record_type(OpenList(i))
        Case "Laboratory Investigations / Laboratory Investigation Report (LIR)"
            OpenRecType(i) = 1
        Case "Laboratory Investigations / Readily Apparent Assignable Cause (RAAC)"
            OpenRecType(i) = 2
        Case "Manufacturing Investigations / Event Report"
            OpenRecType(i) = 3
        Case "Manufacturing Investigations / Quality Assurance Report (QAR)"
            OpenRecType(i) = 4
        Case "Manufacturing Investigations / Incident"
            OpenRecType(i) = 5
    End Select
Next i
'-------------------------------------------------------------
'Compute Area Affected/Area Originated
'-------------------------------------------------------------
ReDim OpenArea(OpenRecNum)
For i = 1 To OpenRecNum
    If Left(areas_affected(OpenList(i)), 8) = "RMT - CQ" Then
        OpenArea(i) = 1
    Else
        If Left(areas_affected(OpenList(i)), 8) = "RMT - SQ" Then
            OpenArea(i) = 2
        Else
            If Left(area_discovered(OpenList(i)), 8) = "RMT - CQ" Then
                OpenArea(i) = 1
            Else
                OpenArea(i) = 0
            End If
        End If
    End If
Next i
'--------------------------------------------------------------
'Compute Subtotal and Grand Total of the Opened Records Matrix
'--------------------------------------------------------------
ReDim OpenRecCount(6, 10, 3)
For i = 0 To 6
    For j = 0 To 10
        For k = 0 To 3
            OpenRecCount(i, j, k) = 0
        Next k
    Next j
Next i
'----------------------------------------------------------------------------
'Capturing record counts with OpenRecCount()
'---------------
'First dimension
'---------------
'0(n/a); 1(LIR); 2(RAAC); 3(ER); 4(QAR); 5(INC); 6(Total)
'----------------
'Second dimension
'----------------
'0(<23); 1(<30); 2(<60); 3(<90); 4(<120); 5(<150); 6(<180); 7(>=180); 8(on-time);
'9(aged); 10(total)
'----------------
'Third Dimension
'----------------
'0(others); 1(Chemistry); 2(Commodity); 3(Total)
'-----------------------------------------------------------------------------
'Count stage and type of records
'-------------------------------
For i = 1 To OpenRecNum
    Select Case OpenRecType(i)
        Case Is = 1
            Select Case OpenStage(i)
                Case Is = 0
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(1, 0, 0) = OpenRecCount(1, 0, 0) + 1
                        Case Is = 1
                            OpenRecCount(1, 0, 1) = OpenRecCount(1, 0, 1) + 1
                        Case Is = 2
                            OpenRecCount(1, 0, 2) = OpenRecCount(1, 0, 2) + 1
                    End Select
                Case Is = 1
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(1, 1, 0) = OpenRecCount(1, 1, 0) + 1
                        Case Is = 1
                            OpenRecCount(1, 1, 1) = OpenRecCount(1, 1, 1) + 1
                        Case Is = 2
                            OpenRecCount(1, 1, 2) = OpenRecCount(1, 1, 2) + 1
                    End Select
                Case Is = 2
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(1, 2, 0) = OpenRecCount(1, 2, 0) + 1
                        Case Is = 1
                            OpenRecCount(1, 2, 1) = OpenRecCount(1, 2, 1) + 1
                        Case Is = 2
                            OpenRecCount(1, 2, 2) = OpenRecCount(1, 2, 2) + 1
                    End Select
                Case Is = 3
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(1, 3, 0) = OpenRecCount(1, 3, 0) + 1
                        Case Is = 1
                            OpenRecCount(1, 3, 1) = OpenRecCount(1, 3, 1) + 1
                        Case Is = 2
                            OpenRecCount(1, 3, 2) = OpenRecCount(1, 3, 2) + 1
                    End Select
                Case Is = 4
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(1, 4, 0) = OpenRecCount(1, 4, 0) + 1
                        Case Is = 1
                            OpenRecCount(1, 4, 1) = OpenRecCount(1, 4, 1) + 1
                        Case Is = 2
                            OpenRecCount(1, 4, 2) = OpenRecCount(1, 4, 2) + 1
                    End Select
                Case Is = 5
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(1, 5, 0) = OpenRecCount(1, 5, 0) + 1
                        Case Is = 1
                            OpenRecCount(1, 5, 1) = OpenRecCount(1, 5, 1) + 1
                        Case Is = 2
                            OpenRecCount(1, 5, 2) = OpenRecCount(1, 5, 2) + 1
                    End Select
                Case Is = 6
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(1, 6, 0) = OpenRecCount(1, 6, 0) + 1
                        Case Is = 1
                            OpenRecCount(1, 6, 1) = OpenRecCount(1, 6, 1) + 1
                        Case Is = 2
                            OpenRecCount(1, 6, 2) = OpenRecCount(1, 6, 2) + 1
                    End Select
                Case Is = 7
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(1, 7, 0) = OpenRecCount(1, 7, 0) + 1
                        Case Is = 1
                            OpenRecCount(1, 7, 1) = OpenRecCount(1, 7, 1) + 1
                        Case Is = 2
                            OpenRecCount(1, 7, 2) = OpenRecCount(1, 7, 2) + 1
                    End Select
            End Select
        Case Is = 2
            Select Case OpenStage(i)
                Case Is = 0
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(2, 0, 0) = OpenRecCount(2, 0, 0) + 1
                        Case Is = 1
                            OpenRecCount(2, 0, 1) = OpenRecCount(2, 0, 1) + 1
                        Case Is = 2
                            OpenRecCount(2, 0, 2) = OpenRecCount(2, 0, 2) + 1
                    End Select
                Case Is = 1
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(2, 1, 0) = OpenRecCount(2, 1, 0) + 1
                        Case Is = 1
                            OpenRecCount(2, 1, 1) = OpenRecCount(2, 1, 1) + 1
                        Case Is = 2
                            OpenRecCount(2, 1, 2) = OpenRecCount(2, 1, 2) + 1
                    End Select
                Case Is = 2
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(2, 2, 0) = OpenRecCount(2, 2, 0) + 1
                        Case Is = 1
                            OpenRecCount(2, 2, 1) = OpenRecCount(2, 2, 1) + 1
                        Case Is = 2
                            OpenRecCount(2, 2, 2) = OpenRecCount(2, 2, 2) + 1
                    End Select
                Case Is = 3
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(2, 3, 0) = OpenRecCount(2, 3, 0) + 1
                        Case Is = 1
                            OpenRecCount(2, 3, 1) = OpenRecCount(2, 3, 1) + 1
                        Case Is = 2
                            OpenRecCount(2, 3, 2) = OpenRecCount(2, 3, 2) + 1
                    End Select
                Case Is = 4
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(2, 4, 0) = OpenRecCount(2, 4, 0) + 1
                        Case Is = 1
                            OpenRecCount(2, 4, 1) = OpenRecCount(2, 4, 1) + 1
                        Case Is = 2
                            OpenRecCount(2, 4, 2) = OpenRecCount(2, 4, 2) + 1
                    End Select
                Case Is = 5
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(2, 5, 0) = OpenRecCount(2, 5, 0) + 1
                        Case Is = 1
                            OpenRecCount(2, 5, 1) = OpenRecCount(2, 5, 1) + 1
                        Case Is = 2
                            OpenRecCount(2, 5, 2) = OpenRecCount(2, 5, 2) + 1
                    End Select
                Case Is = 6
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(2, 6, 0) = OpenRecCount(2, 6, 0) + 1
                        Case Is = 1
                            OpenRecCount(2, 6, 1) = OpenRecCount(2, 6, 1) + 1
                        Case Is = 2
                            OpenRecCount(2, 6, 2) = OpenRecCount(2, 6, 2) + 1
                    End Select
                Case Is = 7
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(2, 7, 0) = OpenRecCount(2, 7, 0) + 1
                        Case Is = 1
                            OpenRecCount(2, 7, 1) = OpenRecCount(2, 7, 1) + 1
                        Case Is = 2
                            OpenRecCount(2, 7, 2) = OpenRecCount(2, 7, 2) + 1
                    End Select
            End Select
        Case Is = 3
            Select Case OpenStage(i)
                Case Is = 0
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(3, 0, 0) = OpenRecCount(3, 0, 0) + 1
                        Case Is = 1
                            OpenRecCount(3, 0, 1) = OpenRecCount(3, 0, 1) + 1
                        Case Is = 2
                            OpenRecCount(3, 0, 2) = OpenRecCount(3, 0, 2) + 1
                    End Select
                Case Is = 1
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(3, 1, 0) = OpenRecCount(3, 1, 0) + 1
                        Case Is = 1
                            OpenRecCount(3, 1, 1) = OpenRecCount(3, 1, 1) + 1
                        Case Is = 2
                            OpenRecCount(3, 1, 2) = OpenRecCount(3, 1, 2) + 1
                    End Select
                Case Is = 2
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(3, 2, 0) = OpenRecCount(3, 2, 0) + 1
                        Case Is = 1
                            OpenRecCount(3, 2, 1) = OpenRecCount(3, 2, 1) + 1
                        Case Is = 2
                            OpenRecCount(3, 2, 2) = OpenRecCount(3, 2, 2) + 1
                    End Select
                Case Is = 3
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(3, 3, 0) = OpenRecCount(3, 3, 0) + 1
                        Case Is = 1
                            OpenRecCount(3, 3, 1) = OpenRecCount(3, 3, 1) + 1
                        Case Is = 2
                            OpenRecCount(3, 3, 2) = OpenRecCount(3, 3, 2) + 1
                    End Select
                Case Is = 4
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(3, 4, 0) = OpenRecCount(3, 4, 0) + 1
                        Case Is = 1
                            OpenRecCount(3, 4, 1) = OpenRecCount(3, 4, 1) + 1
                        Case Is = 2
                            OpenRecCount(3, 4, 2) = OpenRecCount(3, 4, 2) + 1
                    End Select
                Case Is = 5
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(3, 5, 0) = OpenRecCount(3, 5, 0) + 1
                        Case Is = 1
                            OpenRecCount(3, 5, 1) = OpenRecCount(3, 5, 1) + 1
                        Case Is = 2
                            OpenRecCount(3, 5, 2) = OpenRecCount(3, 5, 2) + 1
                    End Select
                Case Is = 6
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(3, 6, 0) = OpenRecCount(3, 6, 0) + 1
                        Case Is = 1
                            OpenRecCount(3, 6, 1) = OpenRecCount(3, 6, 1) + 1
                        Case Is = 2
                            OpenRecCount(3, 6, 2) = OpenRecCount(3, 6, 2) + 1
                    End Select
                Case Is = 7
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(3, 7, 0) = OpenRecCount(3, 7, 0) + 1
                        Case Is = 1
                            OpenRecCount(3, 7, 1) = OpenRecCount(3, 7, 1) + 1
                        Case Is = 2
                            OpenRecCount(3, 7, 2) = OpenRecCount(3, 7, 2) + 1
                    End Select
            End Select
        Case Is = 4
            Select Case OpenStage(i)
                Case Is = 0
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(4, 0, 0) = OpenRecCount(4, 0, 0) + 1
                        Case Is = 1
                            OpenRecCount(4, 0, 1) = OpenRecCount(4, 0, 1) + 1
                        Case Is = 2
                            OpenRecCount(4, 0, 2) = OpenRecCount(4, 0, 2) + 1
                    End Select
                Case Is = 1
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(4, 1, 0) = OpenRecCount(4, 1, 0) + 1
                        Case Is = 1
                            OpenRecCount(4, 1, 1) = OpenRecCount(4, 1, 1) + 1
                        Case Is = 2
                            OpenRecCount(4, 1, 2) = OpenRecCount(4, 1, 2) + 1
                    End Select
                Case Is = 2
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(4, 2, 0) = OpenRecCount(4, 2, 0) + 1
                        Case Is = 1
                            OpenRecCount(4, 2, 1) = OpenRecCount(4, 2, 1) + 1
                        Case Is = 2
                            OpenRecCount(4, 2, 2) = OpenRecCount(4, 2, 2) + 1
                    End Select
                Case Is = 3
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(4, 3, 0) = OpenRecCount(4, 3, 0) + 1
                        Case Is = 1
                            OpenRecCount(4, 3, 1) = OpenRecCount(4, 3, 1) + 1
                        Case Is = 2
                            OpenRecCount(4, 3, 2) = OpenRecCount(4, 3, 2) + 1
                    End Select
                Case Is = 4
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(4, 4, 0) = OpenRecCount(4, 4, 0) + 1
                        Case Is = 1
                            OpenRecCount(4, 4, 1) = OpenRecCount(4, 4, 1) + 1
                        Case Is = 2
                            OpenRecCount(4, 4, 2) = OpenRecCount(4, 4, 2) + 1
                    End Select
                Case Is = 5
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(4, 5, 0) = OpenRecCount(4, 5, 0) + 1
                        Case Is = 1
                            OpenRecCount(4, 5, 1) = OpenRecCount(4, 5, 1) + 1
                        Case Is = 2
                            OpenRecCount(4, 5, 2) = OpenRecCount(4, 5, 2) + 1
                    End Select
                Case Is = 6
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(4, 6, 0) = OpenRecCount(4, 6, 0) + 1
                        Case Is = 1
                            OpenRecCount(4, 6, 1) = OpenRecCount(4, 6, 1) + 1
                        Case Is = 2
                            OpenRecCount(4, 6, 2) = OpenRecCount(4, 6, 2) + 1
                    End Select
                Case Is = 7
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(4, 7, 0) = OpenRecCount(4, 7, 0) + 1
                        Case Is = 1
                            OpenRecCount(4, 7, 1) = OpenRecCount(4, 7, 1) + 1
                        Case Is = 2
                            OpenRecCount(4, 7, 2) = OpenRecCount(4, 7, 2) + 1
                    End Select
            End Select
        Case Is = 5
            Select Case OpenStage(i)
                Case Is = 0
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(5, 0, 0) = OpenRecCount(5, 0, 0) + 1
                        Case Is = 1
                            OpenRecCount(5, 0, 1) = OpenRecCount(5, 0, 1) + 1
                        Case Is = 2
                            OpenRecCount(5, 0, 2) = OpenRecCount(5, 0, 2) + 1
                    End Select
            
                Case Is = 1
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(5, 1, 0) = OpenRecCount(5, 1, 0) + 1
                        Case Is = 1
                            OpenRecCount(5, 1, 1) = OpenRecCount(5, 1, 1) + 1
                        Case Is = 2
                            OpenRecCount(5, 1, 2) = OpenRecCount(5, 1, 2) + 1
                    End Select
                Case Is = 2
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(5, 2, 0) = OpenRecCount(5, 2, 0) + 1
                        Case Is = 1
                            OpenRecCount(5, 2, 1) = OpenRecCount(5, 2, 1) + 1
                        Case Is = 2
                            OpenRecCount(5, 2, 2) = OpenRecCount(5, 2, 2) + 1
                    End Select
                Case Is = 3
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(5, 3, 0) = OpenRecCount(5, 3, 0) + 1
                        Case Is = 1
                            OpenRecCount(5, 3, 1) = OpenRecCount(5, 3, 1) + 1
                        Case Is = 2
                            OpenRecCount(5, 3, 2) = OpenRecCount(5, 3, 2) + 1
                    End Select
                Case Is = 4
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(5, 4, 0) = OpenRecCount(5, 4, 0) + 1
                        Case Is = 1
                            OpenRecCount(5, 4, 1) = OpenRecCount(5, 4, 1) + 1
                        Case Is = 2
                            OpenRecCount(5, 4, 2) = OpenRecCount(5, 4, 2) + 1
                    End Select
                Case Is = 5
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(5, 5, 0) = OpenRecCount(5, 5, 0) + 1
                        Case Is = 1
                            OpenRecCount(5, 5, 1) = OpenRecCount(5, 5, 1) + 1
                        Case Is = 2
                            OpenRecCount(5, 5, 2) = OpenRecCount(5, 5, 2) + 1
                    End Select
                Case Is = 6
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(5, 6, 0) = OpenRecCount(5, 6, 0) + 1
                        Case Is = 1
                            OpenRecCount(5, 6, 1) = OpenRecCount(5, 6, 1) + 1
                        Case Is = 2
                            OpenRecCount(5, 6, 2) = OpenRecCount(5, 6, 2) + 1
                    End Select
                Case Is = 7
                    Select Case OpenArea(i)
                        Case Is = 0
                            OpenRecCount(5, 7, 0) = OpenRecCount(5, 7, 0) + 1
                        Case Is = 1
                            OpenRecCount(5, 7, 1) = OpenRecCount(5, 7, 1) + 1
                        Case Is = 2
                            OpenRecCount(5, 7, 2) = OpenRecCount(5, 7, 2) + 1
                    End Select
                End Select
    End Select
Next i
'--------------------------------------------------------------
'Calculate Summary of the Opened Records
'--------------------------------------------------------------
For i = 1 To 6
    For j = 0 To 2
        OpenRecCount(i, 8, j) = OpenRecCount(i, 0, j) + OpenRecCount(i, 1, j)
        OpenRecCount(i, 9, j) = OpenRecCount(i, 2, j) + OpenRecCount(i, 3, j) _
        + OpenRecCount(i, 4, j) + OpenRecCount(i, 5, j) + OpenRecCount(i, 6, j)
        OpenRecCount(i, 10, j) = OpenRecCount(i, 8, j) + OpenRecCount(i, 9, j)
    Next j
Next i
For i = 0 To 10
    OpenRecCount(6, i, 0) = OpenRecCount(1, i, 0) + OpenRecCount(2, i, 0) _
    + OpenRecCount(3, i, 0) + OpenRecCount(4, i, 0) + OpenRecCount(5, i, 0)
    OpenRecCount(6, i, 1) = OpenRecCount(1, i, 1) + OpenRecCount(2, i, 1) _
    + OpenRecCount(3, i, 1) + OpenRecCount(4, i, 1) + OpenRecCount(5, i, 1)
    OpenRecCount(6, i, 2) = OpenRecCount(1, i, 2) + OpenRecCount(2, i, 2) _
    + OpenRecCount(3, i, 2) + OpenRecCount(4, i, 2) + OpenRecCount(5, i, 2)
Next i
For i = 1 To 6
    For j = 0 To 10
        OpenRecCount(i, j, 3) = OpenRecCount(i, j, 0) + OpenRecCount(i, j, 1) + OpenRecCount(i, j, 2)
    Next j
Next i




'----------------------------------------------------------------
'Write Open Record Description into Array
'----------------------------------------------------------------
ReDim OpenRec(OpenRecNum, 6)
ReDim OpenArea(OpenRecNum)
'--------------------------------------
'First Dimension
'---------------
'Open Record Number 1-OpenRecNum
'----------------
'Second Dimension
'---------------
'1(pr_id); 2(short_description); 3(responsible_person); 4(OpenStage); 5(OpenRecType); 6(OpenArea)
'--------------------------------------
For i = 1 To OpenRecNum
    OpenRec(i, 1) = pr_id(OpenList(i))
    OpenRec(i, 2) = title_short_description(OpenList(i))
    OpenRec(i, 3) = responsible_person(OpenList(i))
    OpenRec(i, 4) = OpenStage(i)
    OpenRec(i, 5) = OpenRecType(i)
    OpenRec(i, 6) = OpenArea(i)
Next i
'----------------------------------------------------------------
'Identify Closed Record within Specified Time Range
'----------------------------------------------------------------
ClosedRecNum = 0
ReDim Closed_Index(Record_Num)
For i = 2 To Record_Num
    If qa_final_app_on(i) >= Period_Begin Then
        If qa_final_app_on(i) < Period_End + 1 Then
            ClosedRecNum = ClosedRecNum + 1
            Closed_Index(i) = i
        Else
            ClosedRecNum = ClosedRecNum
            Closed_Index(i) = 0
        End If
    Else
        If site_qa_approval_on(i) >= Period_Begin Then
            If site_qa_approval_on(i) < Period_End + 1 Then
                ClosedRecNum = ClosedRecNum + 1
                Closed_Index(i) = i
            Else
                ClosedRecNum = ClosedRecNum
                Closed_Index(i) = 0
            End If
        Else
            ClosedRecNum = ClosedRecNum
            Closed_Index(i) = 0
        End If
    End If
Next i
'---------------------------------------------------------
'Writing closed record index into array
'---------------------------------------------------------
ReDim ClosedList(ClosedRecNum)
ClosedList_Pos = 1
For i = 2 To Record_Num
    If Closed_Index(i) <> 0 Then
        ClosedList(ClosedList_Pos) = Closed_Index(i)
        ClosedList_Pos = ClosedList_Pos + 1
    Else
    End If
Next i
'--------------------------------------------------------------------------
'Compute Age and Stage of closed record
'--------------------------------------------------------------------------
ReDim CloseAge(ClosedRecNum)
ReDim CloseStage(ClosedRecNum)
ReDim ClosedRecType(ClosedRecNum)
ReDim ClosedRecCount(ClosedRecNum)
For i = 1 To ClosedRecNum
    If qa_final_app_on(ClosedList(i)) <> 0 Then
        CloseAge(i) = qa_final_app_on(ClosedList(i)) - discovery_date(ClosedList(i))
    Else
        CloseAge(i) = site_qa_approval_on(ClosedList(i)) - discovery_date(ClosedList(i))
    End If
    If CloseAge(i) < 23 Then
        CloseStage(i) = 0
    Else
        If CloseAge(i) < 30 Then
            CloseStage(i) = 1
        Else
            If CloseAge(i) < 60 Then
                CloseStage(i) = 2
            Else
                If CloseAge(i) < 90 Then
                    CloseStage(i) = 3
                Else
                    If CloseAge(i) < 120 Then
                        CloseStage(i) = 4
                    Else
                        If CloseAge(i) < 150 Then
                            CloseStage(i) = 5
                        Else
                            If CloseAge(i) < 180 Then
                                CloseStage(i) = 6
                            Else
                                CloseStage(i) = 7
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
'----------------------------------------------------------------
'Closed Record Type
'----------------------------------------------------------------
    Select Case record_type(ClosedList(i))
        Case "Laboratory Investigations / Laboratory Investigation Report (LIR)"
            ClosedRecType(i) = 1
        Case "Laboratory Investigations / Readily Apparent Assignable Cause (RAAC)"
            ClosedRecType(i) = 2
        Case "Manufacturing Investigations / Event Report"
            ClosedRecType(i) = 3
        Case "Manufacturing Investigations / Quality Assurance Report (QAR)"
            ClosedRecType(i) = 4
        Case "Manufacturing Investigations / Incident"
            ClosedRecType(i) = 5
    End Select
Next i
'----------------------------------------------------------------
'Computing Summary of the Closed Records
'----------------------------------------------------------------
ReDim ClosedRecCount(6, 10)
For i = 0 To 6
    For j = 0 To 2
        ClosedRecCount(i, j) = 0
    Next j
Next i
'---------------------------------------------------------------
'First Dimension
'---------------
'0(n/a); 1(LIR); 2(RAAC); 3(ER); 4(QAR); 5(INC); 6(Total)
'---------------
'Second Dimension
'---------------
'0(<23); 1(<30); 2(<60); 3(<90); 4(<120); 5(<150); 6(<180); 7(>=180); 8(on-time)
'9(aged); 10(total)
'----------------------------------------------------------------
For i = 1 To ClosedRecNum
    Select Case ClosedRecType(i)
        Case Is = 1
            Select Case CloseStage(i)
                Case Is = 0
                    ClosedRecCount(1, 0) = ClosedRecCount(1, 0) + 1
                Case Is = 1
                    ClosedRecCount(1, 1) = ClosedRecCount(1, 1) + 1
                Case Is = 2
                    ClosedRecCount(1, 2) = ClosedRecCount(1, 2) + 1
                Case Is = 3
                    ClosedRecCount(1, 3) = ClosedRecCount(1, 3) + 1
                Case Is = 4
                    ClosedRecCount(1, 4) = ClosedRecCount(1, 4) + 1
                Case Is = 5
                    ClosedRecCount(1, 5) = ClosedRecCount(1, 5) + 1
                Case Is = 6
                    ClosedRecCount(1, 6) = ClosedRecCount(1, 6) + 1
                Case Is = 7
                    ClosedRecCount(1, 7) = ClosedRecCount(1, 7) + 1
            End Select
        Case Is = 2
            Select Case CloseStage(i)
                Case Is = 0
                    ClosedRecCount(2, 0) = ClosedRecCount(2, 0) + 1
                Case Is = 1
                    ClosedRecCount(2, 1) = ClosedRecCount(2, 1) + 1
                Case Is = 2
                    ClosedRecCount(2, 2) = ClosedRecCount(2, 2) + 1
                Case Is = 3
                    ClosedRecCount(2, 3) = ClosedRecCount(2, 3) + 1
                Case Is = 4
                    ClosedRecCount(2, 4) = ClosedRecCount(2, 4) + 1
                Case Is = 5
                    ClosedRecCount(2, 5) = ClosedRecCount(2, 5) + 1
                Case Is = 6
                    ClosedRecCount(2, 6) = ClosedRecCount(2, 6) + 1
                Case Is = 7
                    ClosedRecCount(2, 7) = ClosedRecCount(2, 7) + 1
            End Select
        Case Is = 3
            Select Case CloseStage(i)
                Case Is = 0
                    ClosedRecCount(3, 0) = ClosedRecCount(3, 0) + 1
                Case Is = 1
                    ClosedRecCount(3, 1) = ClosedRecCount(3, 1) + 1
                Case Is = 2
                    ClosedRecCount(3, 2) = ClosedRecCount(3, 2) + 1
                Case Is = 3
                    ClosedRecCount(3, 3) = ClosedRecCount(3, 3) + 1
                Case Is = 4
                    ClosedRecCount(3, 4) = ClosedRecCount(3, 4) + 1
                Case Is = 5
                    ClosedRecCount(3, 5) = ClosedRecCount(3, 5) + 1
                Case Is = 6
                    ClosedRecCount(3, 6) = ClosedRecCount(3, 6) + 1
                Case Is = 7
                    ClosedRecCount(3, 7) = ClosedRecCount(3, 7) + 1
            End Select
        Case Is = 4
            Select Case CloseStage(i)
                Case Is = 0
                    ClosedRecCount(4, 0) = ClosedRecCount(4, 0) + 1
                Case Is = 1
                    ClosedRecCount(4, 1) = ClosedRecCount(4, 1) + 1
                Case Is = 2
                    ClosedRecCount(4, 2) = ClosedRecCount(4, 2) + 1
                Case Is = 3
                    ClosedRecCount(4, 3) = ClosedRecCount(4, 3) + 1
                Case Is = 4
                    ClosedRecCount(4, 4) = ClosedRecCount(4, 4) + 1
                Case Is = 5
                    ClosedRecCount(4, 5) = ClosedRecCount(4, 5) + 1
                Case Is = 6
                    ClosedRecCount(4, 6) = ClosedRecCount(4, 6) + 1
                Case Is = 7
                    ClosedRecCount(4, 7) = ClosedRecCount(4, 7) + 1
            End Select
        Case Is = 5
            Select Case CloseStage(i)
                Case Is = 0
                    ClosedRecCount(5, 0) = ClosedRecCount(5, 0) + 1
                Case Is = 1
                    ClosedRecCount(5, 1) = ClosedRecCount(5, 1) + 1
                Case Is = 2
                    ClosedRecCount(5, 2) = ClosedRecCount(5, 2) + 1
                Case Is = 3
                    ClosedRecCount(5, 3) = ClosedRecCount(5, 3) + 1
                Case Is = 4
                    ClosedRecCount(5, 4) = ClosedRecCount(5, 4) + 1
                Case Is = 5
                    ClosedRecCount(5, 5) = ClosedRecCount(5, 5) + 1
                Case Is = 6
                    ClosedRecCount(5, 6) = ClosedRecCount(5, 6) + 1
                Case Is = 7
                    ClosedRecCount(5, 7) = ClosedRecCount(5, 7) + 1
            End Select
    End Select
Next i
'-------------------------------------------------------------------------------------
For i = 1 To 6
    ClosedRecCount(i, 8) = ClosedRecCount(i, 0) + ClosedRecCount(i, 1)
Next i
For i = 1 To 6
    ClosedRecCount(i, 9) = ClosedRecCount(i, 2) + ClosedRecCount(i, 3) _
    + ClosedRecCount(i, 4) + ClosedRecCount(i, 5) + ClosedRecCount(i, 6) _
    + ClosedRecCount(i, 7)
Next i
For i = 1 To 6
    ClosedRecCount(i, 10) = ClosedRecCount(i, 0) + ClosedRecCount(i, 1) _
    + ClosedRecCount(i, 2) + ClosedRecCount(i, 3) + ClosedRecCount(i, 4) _
    + ClosedRecCount(i, 5) + ClosedRecCount(i, 6) + ClosedRecCount(i, 7)
Next i
For i = 0 To 10
    ClosedRecCount(6, i) = ClosedRecCount(1, i) + ClosedRecCount(2, i) + ClosedRecCount(3, i) _
    + ClosedRecCount(4, i) + ClosedRecCount(5, i)
Next i
'----------------------------------------------------------------
'Write Closed Record Description into Array
'----------------------------------------------------------------
ReDim ClosedRec(ClosedRecNum, 6)
'--------------------------------------
'First Dimension
'---------------
'Closed Record Number 1-ClosedRecNum
'----------------
'Second Dimension
'---------------
'1(pr_id); 2(short_description); 3(responsible_person); 4(OpenStage); 5(OpenRecType); 6(OpenRecArea)
'--------------------------------------
For i = 1 To ClosedRecNum
    ClosedRec(i, 1) = pr_id(ClosedList(i))
    ClosedRec(i, 2) = title_short_description(ClosedList(i))
    ClosedRec(i, 3) = responsible_person(ClosedList(i))
    ClosedRec(i, 4) = CloseStage(i)
    ClosedRec(i, 5) = ClosedRecType(i)
    ClosedRec(i, 6) = areas_affected(ClosedList(i))
Next i
'---------------------------------------------------------------
'Collecting New Record
'---------------------------------------------------------------
ReDim New_Index(Record_Num)
NewRecNum = 0
For i = 1 To Record_Num
    If date_open(i) >= Period_Begin Then
        If date_open(i) < Period_End + 1 Then
            NewRecNum = NewRecNum + 1
            New_Index(i) = i
        Else
            NewRecNum = NewRecNum
            New_Index(i) = 0
        End If
    Else
        NewRecNum = NewRecNum
        New_Index(i) = 0
    End If
Next i
ReDim NewList(NewRecNum)
ReDim NewRec(NewRecNum, 5)
NewList_Pos = 1
For i = 2 To Record_Num
    If New_Index(i) <> 0 Then
        NewList(NewList_Pos) = New_Index(i)
        NewList_Pos = NewList_Pos + 1
    Else
    End If
Next i
ReDim NewCount(5)
For i = 0 To 5
    NewCount(i) = 0
Next i
For i = 1 To NewRecNum
    NewRec(i, 1) = pr_id(NewList(i))
    NewRec(i, 2) = title_short_description(NewList(i))
    NewRec(i, 3) = responsible_person(NewList(i))
    NewRec(i, 4) = areas_affected(NewList(i))
    Select Case record_type(NewList(i))
        Case "Laboratory Investigations / Laboratory Investigation Report (LIR)"
            NewRec(i, 5) = 1
            NewCount(1) = NewCount(1) + 1
        Case "Laboratory Investigations / Readily Apparent Assignable Cause (RAAC)"
            NewRec(i, 5) = 2
            NewCount(2) = NewCount(2) + 1
        Case "Manufacturing Investigations / Event Report"
            NewRec(i, 5) = 3
            NewCount(3) = NewCount(3) + 1
        Case "Manufacturing Investigations / Quality Assurance Report (QAR)"
            NewRec(i, 5) = 4
            NewCount(4) = NewCount(4) + 1
        Case "Manufacturing Investigations / Incident"
            NewRec(i, 5) = 5
            NewCount(5) = NewCount(5) + 1
    End Select
Next i
'----------------------------------------------------------------
'Collecting Cancelled Records
'----------------------------------------------------------------
ReDim Cancel_Index(Record_Num)
CancelRecNum = 0
For i = 1 To Record_Num
    If date_open(i) >= Period_Begin Then
        If date_open(i) <= Period_End Then
            If pr_state(i) = "Cancelled" Then
                CancelRecNum = CancelRecNum + 1
                Cancel_Index(i) = i
            Else
                CancelRecNum = CancelRecNum
                Cancel_Index(i) = 0
            End If
        Else
            CancelRecNum = CancelRecNum
            Cancel_Index(i) = 0
        End If
    Else
        CancelRecNum = CancelRecNum
        Cancel_Index(i) = 0
    End If
Next i
ReDim CancelList(CancelRecNum)
ReDim CancelRec(CancelRecNum, 5)
CancelList_Pos = 1
For i = 2 To Record_Num
    If Cancel_Index(i) <> 0 Then
        CancelList(CancelList_Pos) = Cancel_Index(i)
        CancelList_Pos = CancelList_Pos + 1
    Else
    End If
Next i
ReDim CancelCount(5)
For i = 0 To 5
    CancelCount(i) = 0
Next i
For i = 1 To CancelRecNum
    CancelRec(i, 1) = pr_id(CancelList(i))
    CancelRec(i, 2) = title_short_description(CancelList(i))
    CancelRec(i, 3) = responsible_person(CancelList(i))
    CancelRec(i, 4) = ""
    Select Case record_type(CancelList(i))
        Case "Laboratory Investigations / Laboratory Investigation Report (LIR)"
            CancelRec(i, 5) = 1
            CancelCount(1) = CancelCount(1) + 1
        Case "Laboratory Investigations / Readily Apparent Assignable Cause (RAAC)"
            CancelRec(i, 5) = 2
            CancelCount(2) = CancelCount(2) + 1
        Case "Manufacturing Investigations / Event Report"
            CancelRec(i, 5) = 3
            CancelCount(3) = CancelCount(3) + 1
        Case "Manufacturing Investigations / Quality Assurance Report (QAR)"
            CancelRec(i, 5) = 4
            CancelCount(4) = CancelCount(4) + 1
        Case "Manufacturing Investigations / Incident"
            CancelRec(i, 5) = 5
            CancelCount(5) = CancelCount(5) + 1
    End Select
Next i
'----------------------------------------------------------------
'Generate Summary Report
'----------------------------------------------------------------
Sheets.Add after:=Sheets(DataSheet_Name)
Sheets(Sheets.Count).Select
Select Case Report_Type
    Case Is = 1
        ReportSheet_Name = "Week_" & Week_Num & "_" & Year_Num
    Case Is = 2
        ReportSheet_Name = "Month_" & Month_Num & "_" & Year_Num
    Case Is = 3
        ReportSheet_Name = "Quarter_" & Quarter_Num & "_" & Year_Num
    Case Is = 4
        ReportSheet_Name = "Year_" & Year_Num
    Case Is = 5
        ReportSheet_Name = Year_Num & "_" & Month_Num & "_" & Day_Num & "_" & r_y & "_" & r_m & "_" & r_d
End Select
ReDim DivCount(6, 6) As Integer
'-----------------------------------------------------------------

Sheets(Sheets.Count).Name = ReportSheet_Name
'----------------------------------------------------------------
Summary_Headers:
'----------------------------------------------------------------
'Fill Header Values into Header Array
'----------------------------------------------------------------
Rep_Headers(1) = "Record Type"
Rep_Headers(2) = "<23 Days"
Rep_Headers(3) = "24-30 Days"
Rep_Headers(4) = "31-60 Days"
Rep_Headers(5) = "61-90 Days"
Rep_Headers(6) = "91-120 Days"
Rep_Headers(7) = "121-150 Days"
Rep_Headers(8) = "151-180 Days"
Rep_Headers(9) = ">180 Days"
Rep_Headers(10) = "On-Time"
Rep_Headers(11) = "Aged"
Rep_Headers(12) = "Total"
Rep_Headers(13) = "Record ID"
Rep_Headers(14) = "Short Description"
Rep_Headers(15) = "Responsible Person"
Rep_Headers(16) = "Record Stage"
Rep_Headers(17) = "Record Type"
Rep_Headers(18) = "Area"
Rep_Headers(19) = "LIR"
Rep_Headers(20) = "RAAC"
Rep_Headers(21) = "ER"
Rep_Headers(22) = "QAR"
Rep_Headers(23) = "INC"
Rep_Headers(24) = "Total"
Rep_Headers(25) = "Record Type"
Rep_Headers(26) = "Counts"
Rep_Headers(27) = "Opened, Chemistry"
Rep_Headers(28) = "Opened, Commodity"
Rep_Headers(29) = "Closed, Chemistry"
Rep_Headers(30) = "Closed, Commodity"
Rep_Headers(31) = "Opened, Others"
Rep_Headers(32) = "Closed, Others"
Worksheets(ReportSheet_Name).Cells(2, 1).Activate
For i = 0 To 1
    For j = 1 To 12
        Cells(2 + 8 * i, j).Value = Rep_Headers(j)
    Next j
Next i
For i = 0 To 4
    For j = 18 To 23
        Cells(3 + 8 * i + j - 18, 1).Value = Rep_Headers(j)
    Next j
Next i
For i = 0 To 1
    For j = 24 To 25
        Cells(18 + 8 * i, j - 23).Value = Rep_Headers(j)
    Next j
Next i
Cells(34, 1).Activate
ActiveCell.Value = Rep_Headers(17)
For i = 1 To 6
    ActiveCell.Offset(0, i).Value = Rep_Headers(25 + i)
Next i
'----------------------------------------------------------------
'Writing Record Summary Matrices
'----------------------------------------------------------------
Cells(1, 1).Value = "Records remain opened between " & Period_Begin & "-" & Period_End
For i = 1 To 6
  For j = 0 To 10
      Cells(i + 2, j + 2).Value = OpenRecCount(i, j)
  Next j
Next i
Cells(9, 1).Value = "Records Closed between " & Period_Begin & "-" & Period_End
For i = 1 To 6
  For j = 0 To 10
      Cells(i + 10, j + 2).Value = ClosedRecCount(i, j)
  Next j
Next i
Cells(17, 1).Value = "New Records opened between " & Period_Begin & "-" & Period_End
Cells(25, 1).Value = "Cancelled Records opened between " & Period_Begin & "-" & Period_End
Cells(33, 1).Value = "Records by Type and Area between " & Period_Begin & "-" & Period_End
'-------------------------------------------------------------------
'Writing New Record Summary
'-------------------------------------------------------------------
Cells(18, 2).Activate
For i = 1 To 5
    ActiveCell.Offset(1, 0).Value = NewCount(i)
    ActiveCell.Offset(1, 0).Activate
Next i
Cells(24, 2).Value = NewRecNum
'-----------------------------------------------------------------------
'Writing Cancelled Record Summary
'-----------------------------------------------------------------------
Cells(26, 2).Activate
For i = 1 To 5
    ActiveCell.Offset(1, 0).Value = CancelCount(i)
    ActiveCell.Offset(1, 0).Activate
Next i
Cells(32, 2).Value = CancelRecNum
'-----------------------------------------------------------------------
'Writing Division Summary
'-----------------------------------------------------------------------
Cells(35, 2).Activate
For j = 1 To 6
    For i = 1 To 6
        ActiveCell.Value = DivCount(i, j)
        ActiveCell.Offset(1, 0).Activate
    Next i
    ActiveCell.Offset(-6, 1).Activate
Next j
'----------------------------------------------------------------------------------
'Writing Detail Information of Open Records from Array into Spreadsheet while
'Updating Array that Captured Position of each Record in the Spreadsheet
'----------------------------------------------------------------------------------
ReplCol = Cells(2, 1).End(xlToRight).Column
Cells(1, ReplCol + 1).Activate
ActiveCell.Value = "Opened Records"
For i = 1 To 6
ActiveCell.Offset(1, i - 1).Value = Rep_Headers(12 + i)
Next i
ActiveCell.Offset(1, 0).Activate
For j = 1 To 5
    For i = 1 To OpenRecNum
        If OpenRec(i, 5) = j Then
            ActiveCell.Offset(1, 0).Value = OpenRec(i, 1)
            ActiveCell.Offset(1, 1).Value = OpenRec(i, 2)
            ActiveCell.Offset(1, 2).Value = OpenRec(i, 3)
            Select Case OpenRec(i, 4)
                Case Is = 0
                    ActiveCell.Offset(1, 3).Value = Rep_Headers(2)
                Case Is = 1
                    ActiveCell.Offset(1, 3).Value = Rep_Headers(3)
                Case Is = 2
                    ActiveCell.Offset(1, 3).Value = Rep_Headers(4)
                Case Is = 3
                    ActiveCell.Offset(1, 3).Value = Rep_Headers(5)
                Case Is = 4
                    ActiveCell.Offset(1, 3).Value = Rep_Headers(6)
                Case Is = 5
                    ActiveCell.Offset(1, 3).Value = Rep_Headers(7)
                Case Is = 6
                    ActiveCell.Offset(1, 3).Value = Rep_Headers(8)
                Case Is = 7
                    ActiveCell.Offset(1, 3).Value = Rep_Headers(9)
            End Select
            Select Case OpenRec(i, 5)
                Case Is = 1
                    ActiveCell.Offset(1, 4).Value = Rep_Headers(19)
                Case Is = 2
                    ActiveCell.Offset(1, 4).Value = Rep_Headers(20)
                Case Is = 3
                    ActiveCell.Offset(1, 4).Value = Rep_Headers(21)
                Case Is = 4
                    ActiveCell.Offset(1, 4).Value = Rep_Headers(22)
                Case Is = 5
                    ActiveCell.Offset(1, 4).Value = Rep_Headers(23)
            End Select
            ActiveCell.Offset(1, 5).Value = areas_affected(OpenList(i))
            ActiveCell.Offset(1, 0).Activate
        Else
        End If
    Next i
Next j
Cells(1, 19).Activate
ActiveCell.Value = "Closed Records"
For i = 1 To 6
ActiveCell.Offset(1, i - 1).Value = Rep_Headers(12 + i)
Next i
ActiveCell.Offset(1, 0).Activate
For j = 1 To 5
    For i = 1 To ClosedRecNum
        If ClosedRec(i, 5) = j Then
            ActiveCell.Offset(1, 0).Value = ClosedRec(i, 1)
            ActiveCell.Offset(1, 1).Value = ClosedRec(i, 2)
            ActiveCell.Offset(1, 2).Value = ClosedRec(i, 3)
            Select Case ClosedRec(i, 4)
                Case Is = 0
                    ActiveCell.Offset(1, 3).Value = Rep_Headers(2)
                Case Is = 1
                    ActiveCell.Offset(1, 3).Value = Rep_Headers(3)
                Case Is = 2
                    ActiveCell.Offset(1, 3).Value = Rep_Headers(4)
                Case Is = 3
                    ActiveCell.Offset(1, 3).Value = Rep_Headers(5)
                Case Is = 4
                    ActiveCell.Offset(1, 3).Value = Rep_Headers(6)
                Case Is = 5
                    ActiveCell.Offset(1, 3).Value = Rep_Headers(7)
                Case Is = 6
                    ActiveCell.Offset(1, 3).Value = Rep_Headers(8)
                Case Is = 7
                    ActiveCell.Offset(1, 3).Value = Rep_Headers(9)
            End Select
            Select Case ClosedRec(i, 5)
                Case Is = 1
                    ActiveCell.Offset(1, 4).Value = Rep_Headers(18)
                Case Is = 2
                    ActiveCell.Offset(1, 4).Value = Rep_Headers(19)
                Case Is = 3
                    ActiveCell.Offset(1, 4).Value = Rep_Headers(20)
                Case Is = 4
                    ActiveCell.Offset(1, 4).Value = Rep_Headers(21)
                Case Is = 5
                    ActiveCell.Offset(1, 4).Value = Rep_Headers(22)
            End Select
            ActiveCell.Offset(1, 5).Value = areas_affected(ClosedList(i))
            ActiveCell.Offset(1, 0).Activate
        Else
        End If
    Next i
Next j
ReplCol = ActiveCell.End(xlToRight).Column
Cells(1, ReplCol + 1).Activate
ActiveCell.Value = "New Records"
For i = 1 To 5
ActiveCell.Offset(1, i - 1).Value = Rep_Headers(12 + i)
Next i
ActiveCell.Offset(1, 0).Activate
For j = 1 To 5
    For i = 1 To NewRecNum
        If NewRec(i, 5) = j Then
            ActiveCell.Offset(1, 0).Value = NewRec(i, 1)
            ActiveCell.Offset(1, 1).Value = NewRec(i, 2)
            ActiveCell.Offset(1, 2).Value = NewRec(i, 3)
            ActiveCell.Offset(1, 3).Value = Rep_Headers(2)
            Select Case NewRec(i, 5)
                Case Is = 1
                    ActiveCell.Offset(1, 4).Value = Rep_Headers(18)
                Case Is = 2
                    ActiveCell.Offset(1, 4).Value = Rep_Headers(19)
                Case Is = 3
                    ActiveCell.Offset(1, 4).Value = Rep_Headers(20)
                Case Is = 4
                    ActiveCell.Offset(1, 4).Value = Rep_Headers(21)
                Case Is = 5
                    ActiveCell.Offset(1, 4).Value = Rep_Headers(22)
            End Select
            ActiveCell.Offset(1, 0).Activate
        Else
        End If
    Next i
Next j

'------------------------------------------------------------------
'Charting
'------------------------------------------------------------------
ChartSheet_Name = ReportSheet_Name & "_Chart"
Sheets.Add after:=Sheets(ReportSheet_Name)
Sheets(Sheets.Count).Select
Sheets(Sheets.Count).Name = ChartSheet_Name
ActiveSheet.Shapes.AddChart.Select
ActiveChart.ChartType = xlColumnStacked
ActiveChart.SeriesCollection.NewSeries
ActiveChart.SeriesCollection(1).Name = Rep_Headers(2)
ActiveChart.SeriesCollection(1).Values = ReportSheet_Name & "!" & "$B$3:$B$7"
ActiveChart.SeriesCollection(1).XValues = ReportSheet_Name & "!" & "$A$3:$A$7"
ActiveChart.SeriesCollection(1).Interior.Color = RGB(79, 129, 189)
ActiveChart.SeriesCollection(1).ApplyDataLabels
ActiveChart.SeriesCollection.NewSeries
ActiveChart.SeriesCollection(2).Name = Rep_Headers(3)
ActiveChart.SeriesCollection(2).Values = ReportSheet_Name & "!" & "$C$3:$C$7"
ActiveChart.SeriesCollection(2).Interior.Color = RGB(255, 192, 0)
ActiveChart.SeriesCollection(2).ApplyDataLabels
ActiveChart.SeriesCollection.NewSeries
ActiveChart.SeriesCollection(3).Name = Rep_Headers(11)
ActiveChart.SeriesCollection(3).Values = ReportSheet_Name & "!" & "$K$3:$K$7"
ActiveChart.SeriesCollection(3).ApplyDataLabels
ActiveChart.SeriesCollection(3).Interior.Color = RGB(192, 80, 77)
ActiveChart.SeriesCollection.NewSeries
ActiveChart.SeriesCollection(4).Values = ReportSheet_Name & "!" & "$L$3:$L$7"
ActiveChart.SeriesCollection(4).ChartType = xlLineMarkers
ActiveChart.SeriesCollection(4).ApplyDataLabels
ActiveChart.SeriesCollection(4).DataLabels.Position = xlLabelPositionAbove
ActiveChart.SeriesCollection(4).MarkerStyle = -4142
ActiveChart.SeriesCollection(4).Format.Fill.Visible = msoFalse
ActiveChart.SeriesCollection(4).Format.Line.Visible = msoFalse
ActiveChart.Legend.LegendEntries(4).Delete
'ActiveChart.ChartStyle = 26
With ActiveChart
    .HasTitle = True
    .ChartTitle.Text = "CQ Open Record by Type and Age (Week " & Week_Num & ", " & Right(Period_End, 4) & ")"
End With
ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
ActiveSheet.Shapes.AddChart.Select
ActiveChart.ChartType = xlColumnStacked
ActiveChart.SeriesCollection.NewSeries
ActiveChart.SeriesCollection(1).Name = Rep_Headers(10)
ActiveChart.SeriesCollection(1).Values = ReportSheet_Name & "!" & "$J$11:$J$15"
ActiveChart.SeriesCollection(1).XValues = ReportSheet_Name & "!" & "$A$11:$A$15"
ActiveChart.SeriesCollection(1).ApplyDataLabels
ActiveChart.SeriesCollection.NewSeries
ActiveChart.SeriesCollection(2).Name = Rep_Headers(11)
ActiveChart.SeriesCollection(2).Values = ReportSheet_Name & "!" & "$K$11:$K$15"
ActiveChart.SeriesCollection(2).ApplyDataLabels
ActiveChart.SeriesCollection.NewSeries
ActiveChart.SeriesCollection(3).Name = Rep_Headers(13)
ActiveChart.SeriesCollection(3).Values = ReportSheet_Name & "!" & "$L$11:$L$15"
ActiveChart.SeriesCollection(3).ChartType = xlLineMarkers
ActiveChart.SeriesCollection(3).ApplyDataLabels
ActiveChart.SeriesCollection(3).DataLabels.Position = xlLabelPositionAbove
ActiveChart.SeriesCollection(3).MarkerStyle = -4142
ActiveChart.SeriesCollection(3).Format.Fill.Visible = msoFalse
ActiveChart.SeriesCollection(3).Format.Line.Visible = msoFalse
ActiveChart.Legend.LegendEntries(3).Delete
'ActiveChart.ChartStyle = 26
With ActiveChart
    .HasTitle = True
    .ChartTitle.Text = "CQ Number of Records Closed on Week " & Week_Num & ", " & Right(Period_End, 4)
End With
ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
End Sub
