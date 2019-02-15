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
    Dim RecordType As Integer
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
    Columns(6).EntireColumn.Delete
    Columns(6).EntireColumn.Delete
    OpenICol = OpenICol - 2
  '----------------------------------------
  'Calculate Age
  '----------------------------------------
  Cells(1, OpenICol).Value = "Age"
  For i = 2 To OpenIRow
    Cells(i, OpenICol).Value = Date - Cells(i, 4)
  Next i
  Range(Cells(2, OpenICol), Cells(OpenIRow, OpenICol)).NumberFormat = "0"
  OpenICol = OpenICol + 1
  '----------------------------------------
  'create category
  '----------------------------------------
  Cells(1, OpenICol).Value = "Age Category"
  For i = 2 To OpenIRow
    age = Cells(i, OpenICol - 1).Value
    stage = Application.WorksheetFunction.Floor(age / 30, 1)
    Select Case stage
        Case Is > 6
            Cells(i, OpenICol).Value = 7
        Case Is > 5
            Cells(i, OpenICol).Value = 6
        Case Is > 4
            Cells(i, OpenICol).Value = 5
        Case Is > 3
            Cells(i, OpenICol).Value = 4
        Case Is > 2
            Cells(i, OpenICol).Value = 3
        Case Is > 1
            Cells(i, OpenICol).Value = 2
        Case Is > 0
            Cells(i, OpenICol).Value = 1
    End Select
    If age < 30 Then
        If age >= 23 Then
            Cells(i, OpenICol).Value = 0.5
        Else
            Cells(i, OpenICol).Value = 0
        End If
    Else
    End If
  Next i
  '------------------------------------------
  'Calculating Open Records by Age and Type
  '------------------------------------------
  OpenType_LIR = Application.WorksheetFunction.CountIf(Range("$I$2:$I" & OpenIRow), "Laboratory Investigations / Laboratory Investigation Report (LIR)")
  OpenType_RAAC = Application.WorksheetFunction.CountIf(Range("$I$2:$I" & OpenIRow), "Laboratory Investigations / Readily Apparent Assignable Cause (RAAC)")
  OpenType_ER = Application.WorksheetFunction.CountIf(Range("$I$2:$I" & OpenIRow), "Manufacturing Investigations / Event Report")
  OpenType_INC = Application.WorksheetFunction.CountIf(Range("$I$2:$I" & OpenIRow), "Manufacturing Investigations / Incident")
  OpenType_QAR = Application.WorksheetFunction.CountIf(Range("$I$2:$I" & OpenIRow), "Manufacturing Investigations / Event Report")
  Cells(1, OpenICol + 1).Value = OpenType_LIR
  Cells(2, OpenICol + 1).Value = OpenType_RAAC
  Cells(3, OpenICol + 1).Value = OpenType_ER
  Cells(4, OpenICol + 1).Value = OpenType_INC
  Cells(5, OpenICol + 1).Value = OpenType_QAR
End Sub
