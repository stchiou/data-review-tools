Attribute VB_Name = "Module1"
Sub reviewer_score()
'--------------------------------------------------------------------------------
'variables for store input
'--------------------------------------------------------------------------------
    Dim entry_date() As Date
    Dim reviewer() As String
    Dim lot() As Integer
    Dim assay() As Integer
    Dim potency() As Integer
    Dim impurity() As Integer
    Dim id() As Integer
    Dim possible_score() As Integer
    Dim penalty() As Long
    Dim score() As Long
    Dim record_num As Long
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
    Dim review_score() As Long
    Dim review_score() As Long
'-----------------------------------------------------------------------------
'Prepare for data entry
'-----------------------------------------------------------------------------
    Worksheets("Sheet1").Activate
    Cells(1, 1).Value = "Entry Date"
    Cells(1, 2).Value = "Name"
    Cells(1, 3).Value = "Number of Lots"
    Cells(1, 4).Value = "Assay"
    Cells(1, 5).Value = "Potency"
    Cells(1, 6).Value = "Impurity"
    Cells(1, 7).Value = "ID"
    Cells(1, 8).Value = "Possible Scores"
    Cells(1, 9).Value = "Penalty"
    Cells(1, 10).Value = "Final Score"
    
'-----------------------------------------------------------------------------
'Validation, List
'   With Selection.Validation
'        .Delete
'        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
'        xlBetween, Formula1:="=Sheet3!$A$1:$A$27"
'        .IgnoreBlank = True
'        .InCellDropdown = True
'        .InputTitle = ""
'        .ErrorTitle = ""
'        .InputMessage = ""
'        .ErrorMessage = ""
'        .ShowInput = True
'        .ShowError = True
'    End With
End Sub
