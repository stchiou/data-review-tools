Attribute VB_Name = "Module1"
Sub reviewer_score()
'-----------------------------------------------------------------------------
'Prepare for data entry
'-----------------------------------------------------------------------------
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim btn As Button
    Set wb = ActiveWorkbook
    On Error Resume Next
    Set ws = wb.Sheets("Data_Entry")
    On Error GoTo 0
    If Not ws Is Nothing Then
        MsgBox "The Sheet called " & "Data_Entry" & " already existed in the workbook.", vbExclamation, "Sheet Already Exists!"
        GoTo Entry_Prompt
    Else
        Set ws = wb.Sheets.Add(after:=wb.Sheets(wb.Sheets.Count))
        ws.name = "Data_Entry"
    End If
    Cells(1, 1).Value = "Review Date"
    Cells(1, 2).Value = "Name"
    Cells(1, 3).Value = "Number of Lots"
    Cells(1, 4).Value = "Number of Potency/Impurity in each lot"
    Cells(1, 5).Value = "Number of Potency in each lot"
    Cells(1, 6).Value = "Number of Impurity in each lot"
    Cells(1, 7).Value = "Number of Assay in each lot"
    Cells(1, 8).Value = "Number of ID in each lot"
    Cells(1, 9).Value = "Possible Scores"
    Cells(1, 10).Value = "Penalty"
    Cells(1, 11).Value = "Final Score"
    Range("B2").Select
    With Selection.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=Names!$A$1:$A$27"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Selection.AutoFill Destination:=Range("B2:B1048576"), Type:=xlFillDefault
    Range("B2").End(xlDown).Select
    Set btn = ActiveSheet.Buttons.Add(Range("N2").Left, 100, 50, 100)
    ActiveSheet.Buttons.Select
    With Selection
    .OnAction = "Compute"
    .Characters.Text = "Calculate"
    .Font.Bold = True
    End With
Entry_Prompt:
    Worksheets("Data_Entry").Activate
    Cells(1, 1).Select
    MsgBox ("Enter Data in columns A-G. Click the 'Calculate' button to compute values for columns H-J.")
End Sub
Sub Compute()
'--------------------------------------------------------------------------------
'variables for store input
'--------------------------------------------------------------------------------
    Dim entry_date() As Date
    Dim reviewer() As String
    Dim lot() As Integer
    Dim pot_imp() As Integer
    Dim potency() As Integer
    Dim impurity() As Integer
    Dim assay() As Integer
    Dim id() As Integer
    Dim possible_score() As Integer
    Dim penalty() As Long
    Dim score() As Long
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
Worksheets("Data_Entry").Activate
Cells(1, 1).Activate
record_num = ActiveSheet.UsedRange.Rows.Count
    ReDim entry_date(record_num) As Date
    ReDim reviewer(record_num) As String
    ReDim lot(record_num) As Integer
    ReDim pot_imp(record_num) As Integer
    ReDim potency(record_num) As Integer
    ReDim impurity(record_num) As Integer
    ReDim assay(record_num) As Integer
    ReDim id(record_num) As Integer
    ReDim possible_score(record_num) As Integer
    ReDim penalty(record_num) As Long
    ReDim score(record_num) As Long
    For i = 2 To record_num
      Cells(2, 1).Activate
      entry_date(i) = ActiveCell.Value
      reviewer(i) = ActiveCell.Offset(0, 1).Value
      lot(i) = ActiveCell.Offset(0, 2).Value
      pot_imp(i) = ActiveCell.Offset(0, 3).Value
      potency(i) = ActiveCell.Offset(0, 4).Value
      impurity(i) = ActiveCell.Offset(0, 5).Value
      assay(i) = ActiveCell.Offset(0, 6).Value
      id(i) = ActiveCell.Offset(0, 7).Value
    
    Next i
    

MsgBox ("This is it!!!")
End Sub

