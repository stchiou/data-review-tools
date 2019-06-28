Attribute VB_Name = "Module1"
Sub commitment()
'-----------------------------------------
'Variables for Commitment List
'-----------------------------------------
Dim ComFile As String
Dim ComID() As String
Dim ComWeek() As Integer
Dim ComYear() As Integer
Dim ComStart() As Date
Dim ComEnd() As Date
Dim ComShortDes() As String
Dim ComAge() As Integer
Dim ComStatus() As String
'------------------------------------------
'Variables for reading QTS data
'------------------------------------------
Dim File_1 As String
Dim Record_Num As Long
Dim DataSheet_Name As String
Dim Window_1 As String
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
'-------------------------------------------------------------------------------
'Calculate Record Number and redeclare array for raw data
'-------------------------------------------------------------------------------
Input_data_file:
    File_1 = Application.GetOpenFilename _
        (Title:="Data File", _
        filefilter:="CSV (Comma delimited) (*.csv),*.csv")
    If MsgBox("File contains records to be processed is " & File_1 & ". Is this correct?", vbYesNo) = vbNo Then
        GoTo Input_data_file:
    Else
    End If
DataSheet_Name = Mid(File_1, InStrRev(File_1, "\") + 1, (Len(File_1) - InStrRev(File_1, "\") - 4))
Window_1 = DataSheet_Name & ".csv"
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
'---------------------------------------------------------------------------------
'
End Sub
