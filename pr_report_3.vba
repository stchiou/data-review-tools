Attribute VB_Name = "PR_Status_Report_v3"
Sub PR_Report()
'-----------------------------------------------------------------
'Macro for computing weekly PR Status
'Sean Chiou, version 3, 03/13/2019
'-----------------------------------------------------------------
'Items required:
'1. total opein-categorized by type of records
'2. closed last week
'3. aged > 30 days (bar chart, including data from previous 5 weeks, categorized by types:ER, QAR, LIR, RACAC, INC)
'4. aging up (age > 23 days)
'5. committed to close this week
'6. aged that will close
'------------------------------------------------------------------------------------------------------------------
'Features:
'1. Combine output records with corresponding short description
'2. Computes age of the records
'3. Computes stage of the records based on age
'4. Generate reports
'------------------------------------------------------------------------------------------------------------------
Dim File_1 As String
Dim File_2 As String
Dim Report_Type As Integer
Dim Week_Num As Integer
Dim CutOff As String
Dim Record_Num As Long
Dim DataSheet_Name As String
Dim SnapShot_Name As String
'-------------------------------------------------------
'Fields in raw data
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
Dim discovery_date() As String
Dim date_closed() As String
Dim due_date() As String
Dim original_due_date() As String
Dim number_of_approved_extensions() As Integer
Dim qa_final_app_on() As String
Dim site_qa_approval_on() As String
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
Dim root_cause_1() As String
Dim root_cause_2() As String
Dim root_cause_3() As String
Dim recom_diposition_comments() As String
Dim final_comments() As String
Dim assignable_cause_class() As String
Dim assignable_cause() As String
Dim supplier_name_lot_no() As String
Dim area_disocovered() As String
Dim areas_affected() As String
Dim analyst_personnel_sub_category() As String
Dim pr_state() As String
Dim reason_for_investigating() As String
'-----------------------------------------------------------------
Dim OpenRecNum As Integer
Dim Open_Index() As Integer
Dim OpenList() As Integer
Dim OpenList_Pos As Integer
Dim OpenAge() As Integer
Dim OpenStage() As Integer
Dim OpenRecType() As Integer
Dim OpenRecCount() As Integer
'---------------------------------------------------------------
Dim ClosedRecNum As Integer
Dim Closed_Index() As Integer
Dim ClosedList() As Integer
Dim ClosedList_Pos As Integer
Dim ClosedStage() As Integer
Dim ClosedRecType() As Integer
Dim ClosedRecCount() As Integer
'-----------------------------------------------------------------


Dim temp() As Integer
Dim tempval As Long
Dim OpenRec() As String

Dim ReplCol As Long
Dim ReplRow As Long
Dim CloseSheet_Name As String
Dim i As Integer
Dim j As Integer
Dim age As Integer
Dim stage As Integer
Dim address_1 As String
Dim address_2 As String
'---------------------------------------------------------------------------------
'Capture File Names and Path of Data files
'---------------------------------------------------------------------------------
Report_Type = InputBox("Which type of report you want to generate?" _
    & vbCr & "1. Weekly" _
    & vbCr & "2. Monthly" _
    & vbCr & "3. Quarterly" _
    & vbCr & "4. Annually")
Week_Num = InputBox("Input week number of the year", "WEEK NUMBER")
CutOff = InputBox("Input Cut-off Date for the Report in the format of 'mm/dd/yyyy'", "CUTOFF DATE")
Input1:
    File_1 = Application.GetOpenFilename _
        (Title:="Data File", _
        filefilter:="CSV (Comma delimited) (*.csv),*.csv")
    If MsgBox("File contains records to be processed is " & File_1 & ". Is this correct?", vbYesNo) = vbNo Then
        GoTo Input1:
    Else
    End If
Input2:
    File_2 = Application.GetOpenFilename _
        (Title:="Snapshot File", _
        filefilter:="Worksheet(*.xlsx),*.xlsx")
     If MsgBox("File contains snapshots of the past records is " & File_2 & ". Is this correct?", vbYesNo) = vbNo Then
        GoTo Input2:
    Else
    End If
If MsgBox("These are data files that you select:" _
    & vbCr & File_1 _
    & vbCr & File_2 _
    & vbCr & "Please verify if they are correct.", vbYesNo) = vbNo Then
    GoTo Input1:
Else
End If
DataSheet_Name = Mid(File_1, InStrRev(File_1, "\") + 1, (Len(File_1) - InStrRev(File_1, "\") - 4))
SnapShot_Name = Mid(File_2, InStrRev(File_2, "\") + 1, (Len(File_2) - InStrRev(File_2, "\") - 4))
Window_1 = DataSheet_Name & ".csv"
Window_2 = SnapShot_Name & ".xlsx"
'-------------------------------------------------------------------------------
'Calculate Record Number and redeclare array for raw data
'-------------------------------------------------------------------------------
Workbooks.OpenText Filename:=File_1, local:=True
Workbooks.Open Filename:=File_2, local:=True
Windows(Window_1).Activate
Record_Num = Cells(1, 1).End(xlDown).Row
ReDim pr_id(Record_Num)
ReDim title_short_description(Record_Num)
ReDim responsible_person(Record_Num)
ReDim record_type(Record_Num)
ReDim investigation_type(Record_Num)
ReDim related_records(Record_Num)
ReDim event_code(Record_Num)
ReDim qar_required(Record_Num)
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
ReDim root_cause_1(Record_Num)
ReDim root_cause_2(Record_Num)
ReDim root_cause_3(Record_Num)
ReDim recom_diposition_comments(Record_Num)
ReDim final_comments(Record_Num)
ReDim assignable_cause_class(Record_Num)
ReDim assignable_cause(Record_Num)
ReDim supplier_name_lot_no(Record_Num)
ReDim area_disocovered(Record_Num)
ReDim areas_affected(Record_Num)
ReDim analyst_personnel_sub_category(Record_Num)
ReDim pr_state(Record_Num)
ReDim reason_for_investigation(Record_Num)
For i = 2 To Record_Num
    Cells(i, 1).Activate
    pr_id(i) = ActiveCell.Value
    title_short_description(i) = ActiveCell.Offset(0, 1).Value
    responsible_person(i) = ActiveCell.Offset(0, 2).Value
    record_type(i) = ActiveCell.Offset(0, 3).Value
    investigation_type(i) = ActiveCell.Offset(0, 4).Value
    related_records(i) = ActiveCell.Offset(0, 5).Value
    event_code(i) = ActiveCell.Offset(0, 6).Value
    qar_required(i) = ActiveCell.Offset(0, 7).Value
    special_or_common_cuase(i) = ActiveCell.Offset(0, 8).Value
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
    root_cause_1(i) = ActiveCell.Offset(0, 30).Value
    root_cause_2(i) = ActiveCell.Offset(0, 31).Value
    root_cause_3(i) = ActiveCell.Offset(0, 32).Value
    recom_diposition_comments(i) = ActiveCell.Offset(0, 33).Value
    final_comments(i) = ActiveCell.Offset(0, 34).Value
    assignable_cause_class(i) = ActiveCell.Offset(0, 35).Value
    assignable_cause(i) = ActiveCell.Offset(0, 36).Value
    supplier_name_lot_no(i) = ActiveCell.Offset(0, 37).Value
    area_disocovered(i) = ActiveCell.Offset(0, 38).Value
    areas_affected(i) = ActiveCell.Offset(0, 39).Value
    analyst_personnel_sub_category(i) = ActiveCell.Offset(0, 40).Value
    pr_state(i) = ActiveCell.Offset(0, 41).Value
    reason_for_investigation(i) = ActiveCell.Offset(0, 42).Value
Next i
'------------------------------------------------------------------------------
'Count Number of Open Record
'1. Count all the pr_state that are opened upto cut-off date
'2. Remove the recods from 1. that qa_final_app_on has a value
'3. Remove the records from 2. that site_qa_approval_on has a value
'------------------------------------------------------------------------------
OpenRecNum = 0
ReDim Open_Index(Record_Num)
For i = 2 To Record_Num
  If pr_state(i) <> "Closed" Then
    If pr_state(i) <> "Cancelled" Then
        If pr_state(i) <> "Awaiting SQL Approval" Then
            If InStr("OPUQL", pr_state(i)) = 0 Then
                If discovery_date(i) <= DateValue(CutOff) Then
                    If qa_final_app_on(i) = "" Then
                        If site_qa_approval_on(i) = "" Then
                            OpenRecNum = OpenRecNum + 1
                            Open_Index(i) = i
                        Else
                            OpenRecNum = OpenRecNum
                            Open_Index(i) = 0
                        End If
                    Else
                        OpenRecNum = OpenRecNum
                        Open_Index(i) = 0
                    End If
                Else
                    OpenRecNum = OpenRecNum
                    Open_Index(i) = 0
                End If
            Else
                OpenRecNum = OpenRecNum
                Open_Index(i) = 0
            End If
        Else
            OpenRecNum = OpenRecNum
            Open_Index(i) = 0
        End If
    Else
        OpenRecNum = OpenRecNum
        Open_Index(i) = 0
    End If
  Else
    OpenRecNum = OpenRecNum
    Open_Index(i) = 0
  End If
Next i
'-------------------------------------------------------------------------------
'Fill the list of Open Records with index numbers of the whole data set
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
'Calculate Age , Stage and Type of Open Records
'---------------------------------------------------------------------------------
ReDim OpenAge(OpenRecNum)
ReDim OpenStage(OpenRecNum)
ReDim OpenRecType(OpenRecNum)
For i = 1 To OpenRecNum
    OpenAge(i) = DateValue(CutOff) - DateValue(discovery_date(OpenList(i)))
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
'--------------------------------------------------------------
'Compute Subtotal and Grand Total of the Open Records Matrix
'--------------------------------------------------------------
ReDim OpenRecCount(6, 8)
For i = 0 To 6
    For j = 0 To 8
        OpenRecCount(i, j) = 0
    Next j
Next i
'----------------------------------------------------------------------------
'First dimension
'---------------
'0(n/a); 1(LIR); 2(RAAC); 3(ER); 4(QAR); 5(INC); 6(Total)
'----------------
'Second dimension
'----------------
'0(<23); 1(<30); 2(<60); 3(<90); 4(<120); 5(<150); 6(<180); 7(>=180); 8(total)
'-----------------------------------------------------------------------------
For i = 1 To OpenRecNum
    Select Case OpenRecType(i)
        Case Is = 1
            Select Case OpenStage(i)
                Case Is = 0
                    OpenRecCount(1, 0) = OpenRecCount(1, 0) + 1
                Case Is = 1
                    OpenRecCount(1, 1) = OpenRecCount(1, 1) + 1
                Case Is = 2
                    OpenRecCount(1, 2) = OpenRecCount(1, 2) + 1
                Case Is = 3
                    OpenRecCount(1, 3) = OpenRecCount(1, 3) + 1
                Case Is = 4
                    OpenRecCount(1, 4) = OpenRecCount(1, 4) + 1
                Case Is = 5
                    OpenRecCount(1, 5) = OpenRecCount(1, 5) + 1
                Case Is = 6
                    OpenRecCount(1, 6) = OpenRecCount(1, 6) + 1
                Case Is = 7
                    OpenRecCount(1, 7) = OpenRecCount(1, 7) + 1
            End Select
        Case Is = 2
            Select Case OpenStage(i)
                Case Is = 0
                    OpenRecCount(2, 0) = OpenRecCount(2, 0) + 1
                Case Is = 1
                    OpenRecCount(2, 1) = OpenRecCount(2, 1) + 1
                Case Is = 2
                    OpenRecCount(2, 2) = OpenRecCount(2, 2) + 1
                Case Is = 3
                    OpenRecCount(2, 3) = OpenRecCount(2, 3) + 1
                Case Is = 4
                    OpenRecCount(2, 4) = OpenRecCount(2, 4) + 1
                Case Is = 5
                    OpenRecCount(2, 5) = OpenRecCount(2, 5) + 1
                Case Is = 6
                    OpenRecCount(2, 6) = OpenRecCount(2, 6) + 1
                Case Is = 7
                    OpenRecCount(2, 7) = OpenRecCount(2, 7) + 1
            End Select
        Case Is = 3
            Select Case OpenStage(i)
                Case Is = 0
                    OpenRecCount(3, 0) = OpenRecCount(3, 0) + 1
                Case Is = 1
                    OpenRecCount(3, 1) = OpenRecCount(3, 1) + 1
                Case Is = 2
                    OpenRecCount(3, 2) = OpenRecCount(3, 2) + 1
                Case Is = 3
                    OpenRecCount(3, 3) = OpenRecCount(3, 3) + 1
                Case Is = 4
                    OpenRecCount(3, 4) = OpenRecCount(3, 4) + 1
                Case Is = 5
                    OpenRecCount(3, 5) = OpenRecCount(3, 5) + 1
                Case Is = 6
                    OpenRecCount(3, 6) = OpenRecCount(3, 6) + 1
                Case Is = 7
                    OpenRecCount(3, 7) = OpenRecCount(3, 7) + 1
            End Select
        Case Is = 4
            Select Case OpenStage(i)
                Case Is = 0
                    OpenRecCount(4, 0) = OpenRecCount(4, 0) + 1
                Case Is = 1
                    OpenRecCount(4, 1) = OpenRecCount(4, 1) + 1
                Case Is = 2
                    OpenRecCount(4, 2) = OpenRecCount(4, 2) + 1
                Case Is = 3
                    OpenRecCount(4, 3) = OpenRecCount(4, 3) + 1
                Case Is = 4
                    OpenRecCount(4, 4) = OpenRecCount(4, 4) + 1
                Case Is = 5
                    OpenRecCount(4, 5) = OpenRecCount(4, 5) + 1
                Case Is = 6
                    OpenRecCount(4, 6) = OpenRecCount(4, 6) + 1
                Case Is = 7
                    OpenRecCount(4, 7) = OpenRecCount(4, 7) + 1
            End Select
        Case Is = 5
            Select Case OpenStage(i)
                Case Is = 0
                    OpenRecCount(5, 0) = OpenRecCount(5, 0) + 1
                Case Is = 1
                    OpenRecCount(5, 1) = OpenRecCount(5, 1) + 1
                Case Is = 2
                    OpenRecCount(5, 2) = OpenRecCount(5, 2) + 1
                Case Is = 3
                    OpenRecCount(5, 3) = OpenRecCount(5, 3) + 1
                Case Is = 4
                    OpenRecCount(5, 4) = OpenRecCount(5, 4) + 1
                Case Is = 5
                    OpenRecCount(5, 5) = OpenRecCount(5, 5) + 1
                Case Is = 6
                    OpenRecCount(5, 6) = OpenRecCount(5, 6) + 1
                Case Is = 7
                    OpenRecCount(5, 7) = OpenRecCount(5, 7) + 1
            End Select
    End Select
Next i
'--------------------------------------------------------------
'Calculate Summary of the Opened Records
'--------------------------------------------------------------
For i = 1 To 5
   OpenRecCount(i, 8) = OpenRecCount(i, 0) + OpenRecCount(i, 1) + OpenRecCount(i, 2) _
   + OpenRecCount(i, 3) + OpenRecCount(i, 4) + OpenRecCount(i, 5) + OpenRecCount(i, 6) _
   + OpenRecCount(i, 7)
Next i
For i = 0 To 8
    OpenRecCount(6, i) = OpenRecCount(1, i) + OpenRecCount(2, i) + OpenRecCount(3, i) _
    + OpenRecCount(4, i) + OpenRecCount(5, i)
Next i
'----------------------------------------------------------------
'Identify Closed Record within Specified Time Range
'1. pr_state ="closed", and date_closed >= datevalue(cutoff)-7
'2. qa_final_app_on is not blank, and qa_final_app_on >= datevalue(cutoff)-7
'3. site_qa_approval_on is not blank, and site_qa_approval_on >= datevalue(cutoff)-7
'----------------------------------------------------------------
ClosedRecNum = 0
ReDim Closed_Index(Record_Num)
For i = 2 To Record_Num
    If pr_state(i) = "Closed" Then
        If DateValue(CutOff) - DateValue(date_closed(i)) <= 7 Then
            ClosedRecNum = ClosedRecNum + 1
            Closed_Index(i) = i
        Else 'date_closed(i) >= DateValue(CutOff) - 7
            CloseRecNum = CloseRecNum
            Closed_Index(i) = 0
        End If 'date_closed(i) >= DateValue(CutOff) - 7
    Else 'pr_state(i) = "Closed"
        If qa_final_app_on(i) <> "" Then
            If DateValue(CutOff) - DateValue(qa_final_app_on(i)) <= 7 Then
                ClosedRecNum = ClosedRecNum + 1
                Closed_Index(i) = i
            Else 'DateValue(CutOff) - DateValue(qa_final_app_on(i)) <= 7
                If site_qa_approval_on(i) <> "" Then
                    If DateValue(CutOff) - DateValue(site_qa_approval_on(i)) <= 7 Then
                        ClosedRecNum = ClosedRecNum + 1
                        Closed_Index(i) = i
                    Else 'DateValue(CutOff) - DateValue(site_qa_approval_on(i))
                        ClosedRecNum = ClosedRecNum
                        Closed_Index(i) = 0
                    End If 'DateValue(CutOff) - DateValue(site_qa_approval_on(i))
                Else 'site_qa_approval_on(i) <> "" Then
                    ClosedRecNum = ClosedRecNum
                    Closed_Index(i) = 0
                End If 'site_qa_approval_on(i) <> "" Then
            End If 'DateValue(CutOff) - DateValue(qa_final_app_on(i)) <= 7
        Else 'qa_final_app_on(i) <> ""
            CloseRecNum = CloseRecNum
            Closed_Index(i) = 0
        End If 'qa_final_app_on(i) <> ""
    End If 'pr_state(i) = "Closed"
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
ReDim CloseRecType(ClosedRecNum)
ReDim ClosedRecCount(ClosedRecNum)
For i = 1 To ClosedRecNum
    If qa_final_app_on(ClosedList(i)) <> "" Then
        CloseAge(i) = DateValue(qa_final_app_on(ClosedList(i))) - DateValue(discovery_date(ClosedList(i)))
    Else
        CloseAge(i) = DateValue(site_qa_approval_on(ClosedList(i))) - DateValue(discovery_date(ClosedList(i)))
    End If
    If CloseAge(i) > 30 Then
        CloseStage(i) = 1
    Else
        CloseStage(i) = 0
    End If
'----------------------------------------------------------------
'Closed Recrod Type
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
ReDim ClosedRecCount(6, 2)
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
'0(ontime); 1(aged); 2(Total)
'----------------------------------------------------------------
For i = 1 To ClosedRecNum
    Select Case ClosedRecType(i)
        Case Is = 1
            Select Case CloseStage(i)
                Case Is = 0
                    ClosedRecCount(1, 0) = ClosedRecCount(1, 0) + 1
                Case Is = 1
                    ClosedRecCount(1, 1) = ClosedRecCount(1, 1) + 1
            End Select
        Case Is = 2
            Select Case CloseStage(i)
                Case Is = 0
                    ClosedRecCount(1, 0) = ClosedRecCount(1, 0) + 1
                Case Is = 1
                    ClosedRecCount(1, 1) = ClosedRecCount(1, 1) + 1
            End Select
        Case Is = 3
            Select Case CloseStage(i)
                Case Is = 0
                    ClosedRecCount(1, 0) = ClosedRecCount(1, 0) + 1
                Case Is = 1
                    ClosedRecCount(1, 1) = ClosedRecCount(1, 1) + 1
            End Select
        Case Is = 4
            Select Case CloseStage(i)
                Case Is = 0
                    ClosedRecCount(1, 0) = ClosedRecCount(1, 0) + 1
                Case Is = 1
                    ClosedRecCount(1, 1) = ClosedRecCount(1, 1) + 1
            End Select
        Case Is = 5
            Select Case CloseStage(i)
                Case Is = 0
                    ClosedRecCount(1, 0) = ClosedRecCount(1, 0) + 1
                Case Is = 1
                    ClosedRecCount(1, 1) = ClosedRecCount(1, 1) + 1
            End Select
    End Select
Next i
For i = 1 To 5
    ClosedRecCount(i, 2) = ClosedRecCount(i, 0) + ClosedRecCount(i, 1)
Next i
For i = 0 To 2
    ClosedRecCount(6, i) = ClosedRecCount(1, i) + ClosedRecCount(2, i) _
    + ClosedRecCount(3, i) + ClosedRecCount(4, i) + ClosedRecCount(5, i)
Next i
'----------------------------------------------------------------
'Generate Summary Report
'----------------------------------------------------------------
Sheets.Add after:=Sheets(DataSheet_Name)
Sheets(Sheets.Count).Select
Sheets(Sheets.Count).Name = "Week_" & Week_Num
'----------------------------------------------------------------
'Create Headers Row and Column of the Report
'----------------------------------------------------------------
Worksheets("Week_" & Week_Num).Cells(1, 1).Value = "Record Type"
Worksheets("Week_" & Week_Num).Cells(1, 2).Value = "<23 Days"
Worksheets("Week_" & Week_Num).Cells(1, 3).Value = "24-30 Days"
Worksheets("Week_" & Week_Num).Cells(1, 4).Value = "31-60 Days"
Worksheets("Week_" & Week_Num).Cells(1, 5).Value = "61-90 Days"
Worksheets("Week_" & Week_Num).Cells(1, 6).Value = "91-120 Days"
Worksheets("Week_" & Week_Num).Cells(1, 7).Value = "121-150 Days"
Worksheets("Week_" & Week_Num).Cells(1, 8).Value = "151-180 Days"
Worksheets("Week_" & Week_Num).Cells(1, 9).Value = ">181 Days"
Worksheets("Week_" & Week_Num).Cells(1, 10).Value = "Aged"
Worksheets("Week_" & Week_Num).Cells(1, 11).Value = "Total"
Worksheets("Week_" & Week_Num).Cells(2, 1).Value = "LIR"
Worksheets("Week_" & Week_Num).Cells(3, 1).Value = "RAAC"
Worksheets("Week_" & Week_Num).Cells(4, 1).Value = "ER"
Worksheets("Week_" & Week_Num).Cells(5, 1).Value = "QAR"
Worksheets("Week_" & Week_Num).Cells(6, 1).Value = "INC"
Worksheets("Week_" & Week_Num).Cells(7, 1).Value = "Total"
'----------------------------------------------------------------
'Writing Open Record Matrix
'----------------------------------------------------------------
'For i = 1 To 6
'  For j = 0 To 8
'      Cells(i + 1, j + 2).Value = OpenCount(i, j)
'  Next j
'Next i
''-------------------------------------------------------------
''Update the number of Non-Empty Columns in the Summary Report
''-------------------------------------------------------------
'ReplCol = Cells(1, 1).End(xlToRight).Column
''---------------------------------------------------------------
''Generate Headers for Details Section of the Summary Report
''--------------------------------------------------------------
'For i = 0 To 4
'  Worksheets("Week_" & Week_Num).Cells(1, ReplCol + 4 * i + 1).Value = "Record ID"
'  Worksheets("Week_" & Week_Num).Cells(1, ReplCol + 4 * i + 2).Value = "Short Description"
'  Worksheets("Week_" & Week_Num).Cells(1, ReplCol + 4 * i + 3).Value = "Record Stage"
'  Worksheets("Week_" & Week_Num).Cells(1, ReplCol + 4 * i + 4).Value = "Record Type"
'Next i
''-----------------------------------------------------------------------------------------------
''Create Array to Capture Positions of Where Each Record Being Output in the Summary Spreadsheet
''-----------------------------------------------------------------------------------------------
'OpenCurRec(0, 1) = 2
'OpenCurRec(1, 1) = ReplCol + 1
'OpenCurRec(0, 2) = 2
'OpenCurRec(1, 2) = ReplCol + 5
'OpenCurRec(0, 3) = 2
'OpenCurRec(1, 3) = ReplCol + 9
'OpenCurRec(0, 4) = 2
'OpenCurRec(1, 4) = ReplCol + 13
'OpenCurRec(0, 5) = 2
'OpenCurRec(1, 5) = ReplCol + 17
''----------------------------------------------------------------------------------
''Writing Detail Information of Open Records from Array into Spreadsheet while
''Updating Array that Captured Position of each Record in the Spreadsheet
''----------------------------------------------------------------------------------
'For i = 2 To OpenRecNum
'  If OpenRec(i, 3) = 1 Then
'    Cells(OpenCurRec(0, 1), OpenCurRec(1, 1)).Activate
'    ActiveCell.Value = OpenRec(i, 0)
'    ActiveCell.Offset(0, 1).Value = OpenRec(i, 1)
'    ActiveCell.Offset(0, 2).Value = OpenRec(i, 2)
'    ActiveCell.Offset(0, 3).Value = OpenRec(i, 3)
'    OpenCurRec(0, 1) = OpenCurRec(0, 1) + 1
'    OpenCurRec(1, 1) = OpenCurRec(1, 1)
'  Else
'    If OpenRec(i, 3) = 2 Then
'        Cells(OpenCurRec(0, 2), OpenCurRec(1, 2)).Activate
'        ActiveCell.Value = OpenRec(i, 0)
'        ActiveCell.Offset(0, 1).Value = OpenRec(i, 1)
'        ActiveCell.Offset(0, 2).Value = OpenRec(i, 2)
'        ActiveCell.Offset(0, 3).Value = OpenRec(i, 3)
'        OpenCurRec(0, 2) = OpenCurRec(0, 2) + 1
'        OpenCurRec(1, 2) = OpenCurRec(1, 2)
'    Else
'        If OpenRec(i, 3) = 3 Then
'            Cells(OpenCurRec(0, 3), OpenCurRec(1, 3)).Activate
'            ActiveCell.Value = OpenRec(i, 0)
'            ActiveCell.Offset(0, 1).Value = OpenRec(i, 1)
'            ActiveCell.Offset(0, 2).Value = OpenRec(i, 2)
'            ActiveCell.Offset(0, 3).Value = OpenRec(i, 3)
'            OpenCurRec(0, 3) = OpenCurRec(0, 3) + 1
'            OpenCurRec(1, 3) = OpenCurRec(1, 3)
'        Else
'            If OpenRec(i, 3) = 4 Then
'                Cells(OpenCurRec(0, 4), OpenCurRec(1, 4)).Activate
'                ActiveCell.Value = OpenRec(i, 0)
'                ActiveCell.Offset(0, 1).Value = OpenRec(i, 1)
'                ActiveCell.Offset(0, 2).Value = OpenRec(i, 2)
'                ActiveCell.Offset(0, 3).Value = OpenRec(i, 3)
'                OpenCurRec(0, 4) = OpenCurRec(0, 4) + 1
'                OpenCurRec(1, 4) = OpenCurRec(1, 4)
'            Else
'                If OpenRec(i, 3) = 5 Then
'                    Cells(OpenCurRec(0, 5), OpenCurRec(1, 5)).Activate
'                    ActiveCell.Value = OpenRec(i, 0)
'                    ActiveCell.Offset(0, 1).Value = OpenRec(i, 1)
'                    ActiveCell.Offset(0, 2).Value = OpenRec(i, 2)
'                    ActiveCell.Offset(0, 3).Value = OpenRec(i, 3)
'                    OpenCurRec(0, 5) = OpenCurRec(0, 5) + 1
'                    OpenCurRec(1, 5) = OpenCurRec(1, 5)
'                Else
'                End If
'            End If
'        End If
'    End If
'  End If
'Next i
'ReplCol = Worksheets("Week_" & Week_Num).Cells(1, 1).End(xlToRight).Column
''--------------------------------------------------------------------------
''Open Files Contains Closed Records and Short Description of Closed Records
''Insert Short Descriptions to the Sheet that Contains Closed Records
''--------------------------------------------------------------------------
'CloseSheet_Name = Left(File_3, InStr(File_3, ".") - 1)
'Workbooks.OpenText Filename:="C:\Users\chious\Box Sync\vba-projects\pr-status\week" & Week_Num & "\" & File_3, local:=True
'Workbooks.OpenText Filename:="C:\Users\chious\Box Sync\vba-projects\pr-status\week" & Week_Num & "\" & File_4, local:=True
'Columns("E:E").Select
'Selection.Copy
'Windows(File_3).Activate
'Columns("C:C").Select
'Selection.Insert Shift:=xlToRight
'Worksheets(CloseSheet_Name).Activate
'CloselRow = Cells(1, 1).End(xlDown).Row
'CloselCol = Cells(1, 1).End(xlToRight).Column
''----------------------------------------
''Calculate Age of the Closed Records
''----------------------------------------
'CloseRecNum = CloselRow
''CloseRecNum is the line number of the last line that contain close record;
''Total closed Record Number = CloseRecNum -1
'Cells(1, CloselCol).Value = "Age"
'ReDim CloseAge(CloselRow) As Integer
'ReDim CloseStage(CloselRow) As Integer
'ReDim CloseRecType(CloselRow) As Integer
'For i = 2 To CloselRow
'  CloseAge(i) = Date - Cells(i, 4)
'  Cells(i, CloselCol).Value = CloseAge(i)
'Next i
'Range(Cells(2, CloselCol), Cells(CloselRow, CloselCol)).NumberFormat = "0"
'CloselCol = CloselCol + 1
''----------------------------------------
''create category
''----------------------------------------
'Cells(1, CloselCol).Value = "Stage"
'Cells(1, CloselCol + 1).Value = "Type"
'For i = 2 To CloseRecNum
'      If CloseAge(i) > 30 Then
'          CloseStage(i) = 1
'      Else
'          If CloseAge(i) <= 30 Then
'              CloseStage(i) = 0
'          Else
'          End If
'      End If
'  temp = Cells(i, 11).Value
'  Select Case temp
'      Case "Laboratory Investigations / Laboratory Investigation Report (LIR)"
'          CloseRecType(i) = 1
'      Case "Laboratory Investigations / Readily Apparent Assignable Cause (RAAC)"
'          CloseRecType(i) = 2
'      Case "Manufacturing Investigations / Event Report"
'          CloseRecType(i) = 3
'      Case "Manufacturing Investigations / Quality Assurance Report (QAR)"
'          CloseRecType(i) = 4
'      Case "Manufacturing Investigations / Incident"
'          CloseRecType(i) = 5
'  End Select
'  Cells(i, CloselCol).Value = CloseStage(i)
'  Cells(i, CloselCol + 1).Value = CloseRecType(i)
'Next i
'CloselCol = CloselCol + 2
'ReDim CloseCount(6, 2) As Integer
'ReDim CloseRec(CloseRecNum, 3) As String
'ReDim CloseCurRec(1, 5) As Integer
'For i = 0 To 6
'  For j = 0 To 2
'      CloseCount(i, j) = 0
'  Next j
'Next i
'For i = 2 To CloselRow
'  Select Case CloseRecType(i)
'      Case Is = 1
'          Select Case CloseStage(i)
'              Case Is = 0
'                  CloseCount(1, 0) = CloseCount(1, 0) + 1
'              Case Is = 1
'                  CloseCount(1, 1) = CloseCount(1, 1) + 1
'          End Select
'      Case Is = 2
'          Select Case CloseStage(i)
'              Case Is = 0
'                  CloseCount(2, 0) = CloseCount(2, 0) + 1
'              Case Is = 1
'                  CloseCount(2, 1) = CloseCount(2, 1) + 1
'          End Select
'      Case Is = 3
'          Select Case CloseStage(i)
'              Case Is = 0
'                  CloseCount(3, 0) = CloseCount(3, 0) + 1
'              Case Is = 1
'                  CloseCount(3, 1) = CloseCount(3, 1) + 1
'          End Select
'      Case Is = 4
'          Select Case CloseStage(i)
'              Case Is = 0
'                  CloseCount(4, 0) = CloseCount(4, 0) + 1
'              Case Is = 1
'                  CloseCount(4, 1) = CloseCount(4, 1) + 1
'          End Select
'      Case Is = 5
'          Select Case CloseStage(i)
'              Case Is = 0
'                  CloseCount(5, 0) = CloseCount(5, 0) + 1
'              Case Is = 1
'                  CloseCount(5, 1) = CloseCount(5, 1) + 1
'          End Select
'  End Select
'  CloseRec(i, 0) = Worksheets(CloseSheet_Name).Cells(i, 1).Value
'  CloseRec(i, 1) = Worksheets(CloseSheet_Name).Cells(i, 3).Value
'  CloseRec(i, 2) = CloseStage(i)
'  CloseRec(i, 3) = CloseRecType(i)
'Next i
'For i = 1 To 5
'  CloseCount(i, 2) = CloseCount(i, 0) + CloseCount(i, 1)
'Next i
'For i = 0 To 2
'  CloseCount(6, i) = CloseCount(1, i) + CloseCount(2, i) + CloseCount(3, i) + CloseCount(4, i) + CloseCount(5, i)
'Next i
''---------------------------------------------------------------------------
'ReplCol = ReplCol + 1
'Windows(File_1).Activate
'Worksheets("Week_" & Week_Num).Cells(1, ReplCol).Activate
'ActiveCell.Value = "Recod Type"
'ActiveCell.Offset(0, 1).Value = "On Time"
'ActiveCell.Offset(0, 2).Value = "Aged"
'ActiveCell.Offset(0, 3).Value = "Total"
'ActiveCell.Offset(1, 0).Value = "LIR"
'ActiveCell.Offset(2, 0).Value = "RAAC"
'ActiveCell.Offset(3, 0).Value = "ER"
'ActiveCell.Offset(4, 0).Value = "QAR"
'ActiveCell.Offset(5, 0).Value = "INC"
'ActiveCell.Offset(6, 0).Value = "Total"
'For i = 1 To 6
'  For j = 0 To 2
'      ActiveCell.Offset(i, j + 1).Offset.Value = CloseCount(i, j)
'  Next
'Next i
'ReplCol = Cells(1, 1).End(xlToRight).Column + 1
'For i = 0 To 4
'  Worksheets("Week_" & Week_Num).Cells(1, ReplCol + 4 * i).Value = "Record ID"
'  Worksheets("Week_" & Week_Num).Cells(1, ReplCol + 4 * i + 1).Value = "Short Description"
'  Worksheets("Week_" & Week_Num).Cells(1, ReplCol + 4 * i + 2).Value = "Record Stage"
'  Worksheets("Week_" & Week_Num).Cells(1, ReplCol + 4 * i + 3).Value = "Record Type"
'Next i
'CloseCurRec(0, 1) = 2
'CloseCurRec(1, 1) = ReplCol
'CloseCurRec(0, 2) = 2
'CloseCurRec(1, 2) = CloseCurRec(1, 1) + 4
'CloseCurRec(0, 3) = 2
'CloseCurRec(1, 3) = CloseCurRec(1, 2) + 4
'CloseCurRec(0, 4) = 2
'CloseCurRec(1, 4) = CloseCurRec(1, 3) + 4
'CloseCurRec(0, 5) = 2
'CloseCurRec(1, 5) = CloseCurRec(1, 4) + 4
'For i = 2 To CloseRecNum
'  If CloseRec(i, 3) = 1 Then
'      Cells(CloseCurRec(0, 1), CloseCurRec(1, 1)).Activate
'      ActiveCell.Value = CloseRec(i, 0)
'      ActiveCell.Offset(0, 1).Value = CloseRec(i, 1)
'      ActiveCell.Offset(0, 2).Value = CloseRec(i, 2)
'      ActiveCell.Offset(0, 3).Value = CloseRec(i, 3)
'      CloseCurRec(0, 1) = CloseCurRec(0, 1) + 1
'      CloseCurRec(1, 1) = CloseCurRec(1, 1)
'  Else
'      If CloseRec(i, 3) = 2 Then
'          Cells(CloseCurRec(0, 2), CloseCurRec(1, 2)).Activate
'          ActiveCell.Value = CloseRec(i, 0)
'          ActiveCell.Offset(0, 1).Value = CloseRec(i, 1)
'          ActiveCell.Offset(0, 2).Value = CloseRec(i, 2)
'          ActiveCell.Offset(0, 3).Value = CloseRec(i, 3)
'          CloseCurRec(0, 2) = CloseCurRec(0, 2) + 1
'          CloseCurRec(1, 2) = CloseCurRec(1, 2)
'      Else
'          If CloseRec(i, 3) = 3 Then
'              Cells(CloseCurRec(0, 3), CloseCurRec(1, 3)).Activate
'              ActiveCell.Value = CloseRec(i, 0)
'              ActiveCell.Offset(0, 1).Value = CloseRec(i, 1)
'              ActiveCell.Offset(0, 2).Value = CloseRec(i, 2)
'              ActiveCell.Offset(0, 3).Value = CloseRec(i, 3)
'              CloseCurRec(0, 3) = CloseCurRec(0, 3) + 1
'              CloseCurRec(1, 3) = CloseCurRec(1, 3)
'          Else
'              If CloseRec(i, 3) = 4 Then
'                  Cells(CloseCurRec(0, 4), CloseCurRec(1, 4)).Activate
'                  ActiveCell.Value = CloseRec(i, 0)
'                  ActiveCell.Offset(0, 1).Value = CloseRec(i, 1)
'                  ActiveCell.Offset(0, 2).Value = CloseRec(i, 2)
'                  ActiveCell.Offset(0, 3).Value = CloseRec(i, 3)
'                  CloseCurRec(0, 4) = CloseCurRec(0, 4) + 1
'                  CloseCurRec(1, 4) = CloseCurRec(1, 4)
'              Else
'                  If CloseRec(i, 3) = 5 Then
'                      Cells(CloseCurRec(0, 5), CloseCurRec(1, 5)).Activate
'                      ActiveCell.Value = CloseRec(i, 0)
'                      ActiveCell.Offset(0, 1).Value = CloseRec(i, 1)
'                      ActiveCell.Offset(0, 2).Value = CloseRec(i, 2)
'                      ActiveCell.Offset(0, 3).Value = CloseRec(i, 3)
'                      CloseCurRec(0, 5) = CloseCurRec(0, 5) + 1
'                      CloseCurRec(1, 5) = CloseCurRec(1, 5)
'                  Else
'                  End If
'              End If
'          End If
'      End If
'  End If
'Next i
'Worksheets("Week_" & Week_Num).Cells(1, 1).Activate
'ActiveCell.EntireRow.Insert
'Cells(1, 1).Value = "Open Records"
'Cells(1, 12).Value = "Open LIR"
'Cells(1, 16).Value = "Open RAAC"
'Cells(1, 20).Value = "Open ER"
'Cells(1, 24).Value = "Open QAR"
'Cells(1, 28).Value = "Open INC"
'Cells(1, 32).Value = "Closed Records"
'Cells(1, 36).Value = "Closed LIR"
'Cells(1, 40).Value = "Closed RAAC"
'Cells(1, 44).Value = "Closed ER"
'Cells(1, 48).Value = "Closed QAR"
'Cells(1, 52).Value = "Closed INC"
'address_1 = Cells(1, 1).Address(rowabsolute:=False, columnabsolute:=False)
'address_2 = Cells(1, 11).Address(rowabsolute:=False, columnabsolute:=False)
'Range(address_1 & ":" & address_2).Select
'Selection.Merge
'For i = 3 To 13
'    address_1 = Cells(1, 4 * i).Address(rowabsolute:=False, columnabsolute:=False)
'    address_2 = Cells(1, 4 * i + 3).Address(rowabsolute:=False, columnabsolute:=False)
'    Range(address_1 & ":" & address_2).Select
'    Selection.Merge
'Next i
'Worksheets("Week_" & WeekNum).Move
End Sub
