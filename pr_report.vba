Attribute VB_Name = "PR_Status_Report"
Sub PR_Report()
'-----------------------------------------------------------------
'Macro for computing weekly PR Status
'Sean Chiou, version 1.0, 02/14/2019
'-----------------------------------------------------------------
'Items required:
'1. total opein-categorized by type of records
'2. closed last week
'3. aged > 30 days (bar chart, including data from previous 5 weeks, categorized by types:ER, QAR, LIR, RACAC, INC)
'4. aging up (age > 23 days)
'5. committed to close this week
'6. aged that will close
'------------------------------------------------------------------------------------------------------------------
    Dim OpenType_QAR As Integer
    Dim OpenType_LIR As Integer
    Dim OpenType_RAAC As Integer
    Dim OpenType_ER As Integer
    Dim OpenType_INC As Integer
    Dim OpenIRow As Integer
    Dim OpenICol As Integer
    Dim ClosedType_QAR As Integer
    Dim ClosedType_LIR As Integer
    Dim ClosedType_RAAC As Integer
    Dim ClosedType_ER As Integer
    Dim ClosedType_INC As Integer
    Dim ClosedIRow As Integer
    Dim ClosedICol As Integer
    Dim i As Integer
    Dim age As Integer
    Dim stage As Integer
    Worksheets("open").Activate
    OpenIRow = Cells(1, 1).End(xlDown).Row
    OpenICol = Cells(1, 1).End(xlToRight).Column
    '------------------------------------------
    'Removing approved record
    '------------------------------------------
    For i = 2 To OpenIRow
        If Cells(i, 6).Value > 0 Then
            Rows(i).EntireRow.Delete
            i = i - 1
            OpenIRow = OpenIRow - 1
        Else
        End If
    Next i
    For i = 2 To OpenIRow
        If Cells(i, 7).Value > 0 Then
            Rows(i).EntireRow.Delete
            i = i - 1
            OpenIRow = OpenIRow - 1
        Else
        End If
    Next i
  '----------------------------------------
  'Calculate Age
  '----------------------------------------
  Cells(1, OpenICol).Value = "Age"
  For i = 2 To OpenIRow
    Cells(i, OpenICol).Value = Date - Cells(i, 4)
  Next i
  Range(Cells(2, OpenICol), Cells(OpenIRow, OpenICol)).NumberFormat = "0"
End Sub
