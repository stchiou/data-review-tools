Attribute VB_Name = "PR_Status_Report"
Sub PR_Report()
'-----------------------------------------------------------------
'Macro for computing weekly PR Status
'Sean Chiou, version 1.1, 02/19/2019
'-----------------------------------------------------------------
'Items required:
'1. total opein-categorized by type of records
'2. closed last week
'3. aged > 30 days (bar chart, including data from previous 5 weeks, categorized by types:ER, QAR, LIR, RACAC, INC)
'4. aging up (age > 23 days)
'5. committed to close this week
'6. aged that will close
'------------------------------------------------------------------------------------------------------------------
    Dim OpenCount() As Integer
    Dim OpenAge() As Integer
    Dim OpenStage() As Integer
    Dim OpenIRow As Integer
    Dim OpenICol As Integer
    Dim RecType() As Integer
    Dim temp As String
    Dim ClosedType_QAR() As Integer
    Dim ClosedType_LIR() As Integer
    Dim ClosedType_RAAC() As Integer
    Dim ClosedType_ER() As Integer
    Dim ClosedType_INC() As Integer
    Dim ClosedAge() As Integer
    Dim ClosedIRow As Integer
    Dim ClosedICol As Integer
    Dim i As Integer
    Dim j As Integer
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
  ReDim OpenAge(OpenIRow) As Integer
  ReDim OpenStage(OpenIRow) As Integer
  ReDim RecType(OpenIRow) As Integer
  For i = 2 To OpenIRow
    OpenAge(i) = Date - Cells(i, 4)
    Cells(i, OpenICol).Value = OpenAge(i)
  Next i
  Range(Cells(2, OpenICol), Cells(OpenIRow, OpenICol)).NumberFormat = "0"
  OpenICol = OpenICol + 1
  '----------------------------------------
  'create category
  '----------------------------------------
  Cells(1, OpenICol).Value = "Stage"
  Cells(1, OpenICol + 1).Value = "Type"
  
  For i = 2 To OpenIRow
    OpenStage(i) = Application.WorksheetFunction.Floor(OpenAge(i) / 30, 1)
    Select Case OpenStage(i)
        Case Is > 5
            OpenStage(i) = 7
        Case Is > 4
            OpenStage(i) = 6
        Case Is > 3
            OpenStage(i) = 5
        Case Is > 2
            OpenStage(i) = 4
        Case Is > 1
            OpenStage(i) = 3
        Case Is > 0
    End Select
    If OpenAge(i) < 30 Then
        If OpenAge(i) >= 23 Then
            OpenStage(i) = 1
        Else
            OpenStage(i) = 0
        End If
    Else
    End If
    temp = Cells(i, 9).Value
    Select Case temp
        Case "Laboratory Investigations / Laboratory Investigation Report (LIR)"
            RecType(i) = 1
        Case "Laboratory Investigations / Readily Apparent Assignable Cause (RAAC)"
            RecType(i) = 2
        Case "Manufacturing Investigations / Event Report"
            RecType(i) = 3
        Case "Manufacturing Investigations / Quality Assurance Report (QAR)"
            RecType(i) = 4
        Case "Manufacturing Investigations / Incident"
            RecType(i) = 5
    End Select
    Cells(i, OpenICol).Value = OpenStage(i)
    Cells(i, OpenICol + 1).Value = RecType(i)
  Next i
  OpenICol = OpenICol + 2
  '--------------------------------------------------------------------------------
  'Calculating Open Records by Age and Type
  '--------------------------------------------------------------------------------
  'Text for Record types:
  'LIR:     "Laboratory Investigations / Laboratory Investigation Report (LIR)"
  'RAAC:    "Laboratory Investigations / Readily Apparent Assignable Cause (RAAC)"
  'ER:      "Manufacturing Investigations / Event Report"
  'QAR:     "Manufacturing Investigations / Quality Assurance Report (QAR)"
  'INC:     "Manufacturing Investigations / Incident"
  '------------------------------
  'Array subscripts:
  '-----------------------------------
  'First Dimension | Second Dimension
  '-----------------------------------
  '1: LIR          | 0: < 30
  '2: RAAC         | 1: 23-30
  '3: ER           | 2: 31-60
  '4: QAR          | 3: 61-90
  '5: INC          | 4: 91-120
  '                | 5: 121-150
  '                | 6: 151-180
  '                | 7: >180
  '---------------------------------------------------------------------------------
  ReDim OpenCount(5, 7) As Integer
  For i = 2 To OpenIRow
    
  Next i
  '------------------------------------------
  'Codes from V 1.0
  '<23
  '------------------------------------------
'  OpenType_LIR_0 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Laboratory Investigations / Laboratory Investigation Report (LIR)", Range("$P$2:$P" & OpenIRow), 0)
'  OpenType_RAAC_0 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Laboratory Investigations / Readily Apparent Assignable Cause (RAAC)", Range("$P$2:$P" & OpenIRow), 0)
'  OpenType_ER_0 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Manufacturing Investigations / Event Report", Range("$P$2:$P" & OpenIRow), 0)
'  OpenType_INC_0 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Manufacturing Investigations / Incident", Range("$P$2:$P" & OpenIRow), 0)
'  OpenType_QAR_0 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Manufacturing Investigations / Quality Assurance Report (QAR)", Range("$P$2:$P" & OpenIRow), 0)
'  '------------------------------------------
'  'Aging
'  '------------------------------------------
'  OpenType_LIR_0_5 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Laboratory Investigations / Laboratory Investigation Report (LIR)", Range("$P$2:$P" & OpenIRow), 0.5)
'  OpenType_RAAC_0_5 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Laboratory Investigations / Readily Apparent Assignable Cause (RAAC)", Range("$P$2:$P" & OpenIRow), 0.5)
'  OpenType_ER_0_5 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Manufacturing Investigations / Event Report", Range("$P$2:$P" & OpenIRow), 0.5)
'  OpenType_INC_0_5 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Manufacturing Investigations / Incident", Range("$P$2:$P" & OpenIRow), 0.5)
'  OpenType_QAR_0_5 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Manufacturing Investigations / Quality Assurance Report (QAR)", Range("$P$2:$P" & OpenIRow), 0.5)
'  '-----------------------------------
'  '31-60
'  '-----------------------------------
'  OpenType_LIR_1 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Laboratory Investigations / Laboratory Investigation Report (LIR)", Range("$P$2:$P" & OpenIRow), 1)
'  OpenType_RAAC_1 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Laboratory Investigations / Readily Apparent Assignable Cause (RAAC)", Range("$P$2:$P" & OpenIRow), 1)
'  OpenType_ER_1 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Manufacturing Investigations / Event Report", Range("$P$2:$P" & OpenIRow), 1)
'  OpenType_INC_1 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Manufacturing Investigations / Incident", Range("$P$2:$P" & OpenIRow), 1)
'  OpenType_QAR_1 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Manufacturing Investigations / Quality Assurance Report (QAR)", Range("$P$2:$P" & OpenIRow), 1)
'  '-------------------------------------
'  '61-90
'  '-------------------------------------
'  OpenType_LIR_2 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Laboratory Investigations / Laboratory Investigation Report (LIR)", Range("$P$2:$P" & OpenIRow), 2)
'  OpenType_RAAC_2 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Laboratory Investigations / Readily Apparent Assignable Cause (RAAC)", Range("$P$2:$P" & OpenIRow), 2)
'  OpenType_ER_2 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Manufacturing Investigations / Event Report", Range("$P$2:$P" & OpenIRow), 2)
'  OpenType_INC_2 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Manufacturing Investigations / Incident", Range("$P$2:$P" & OpenIRow), 2)
'  OpenType_QAR_2 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Manufacturing Investigations / Quality Assurance Report (QAR)", Range("$P$2:$P" & OpenIRow), 2)
'  '-------------------------------------
'  '91-120
'  '-------------------------------------
'  OpenType_LIR_3 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Laboratory Investigations / Laboratory Investigation Report (LIR)", Range("$P$2:$P" & OpenIRow), 3)
'  OpenType_RAAC_3 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Laboratory Investigations / Readily Apparent Assignable Cause (RAAC)", Range("$P$2:$P" & OpenIRow), 3)
'  OpenType_ER_3 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Manufacturing Investigations / Event Report", Range("$P$2:$P" & OpenIRow), 3)
'  OpenType_INC_3 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Manufacturing Investigations / Incident", Range("$P$2:$P" & OpenIRow), 3)
'  OpenType_QAR_3 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Manufacturing Investigations / Quality Assurance Report (QAR)", Range("$P$2:$P" & OpenIRow), 3)
'  '--------------------------------------
'  '121-150
'  '--------------------------------------
'  OpenType_LIR_4 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Laboratory Investigations / Laboratory Investigation Report (LIR)", Range("$P$2:$P" & OpenIRow), 4)
'  OpenType_RAAC_4 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Laboratory Investigations / Readily Apparent Assignable Cause (RAAC)", Range("$P$2:$P" & OpenIRow), 4)
'  OpenType_ER_4 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Manufacturing Investigations / Event Report", Range("$P$2:$P" & OpenIRow), 4)
'  OpenType_INC_4 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Manufacturing Investigations / Incident", Range("$P$2:$P" & OpenIRow), 4)
'  OpenType_QAR_4 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Manufacturing Investigations / Quality Assurance Report (QAR)", Range("$P$2:$P" & OpenIRow), 4)
'  '--------------------------------------
'  '151-180
'  '--------------------------------------
'  OpenType_LIR_5 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Laboratory Investigations / Laboratory Investigation Report (LIR)", Range("$P$2:$P" & OpenIRow), 5)
'  OpenType_RAAC_5 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Laboratory Investigations / Readily Apparent Assignable Cause (RAAC)", Range("$P$2:$P" & OpenIRow), 5)
'  OpenType_ER_5 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Manufacturing Investigations / Event Report", Range("$P$2:$P" & OpenIRow), 5)
'  OpenType_INC_5 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Manufacturing Investigations / Incident", Range("$P$2:$P" & OpenIRow), 5)
'  OpenType_QAR_5 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Manufacturing Investigations / Quality Assurance Report (QAR)", Range("$P$2:$P" & OpenIRow), 5)
'  '--------------------------------------
'  '>181
'  '--------------------------------------
'  OpenType_LIR_6 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Laboratory Investigations / Laboratory Investigation Report (LIR)", Range("$P$2:$P" & OpenIRow), 6)
'  OpenType_RAAC_6 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Laboratory Investigations / Readily Apparent Assignable Cause (RAAC)", Range("$P$2:$P" & OpenIRow), 6)
'  OpenType_ER_6 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Manufacturing Investigations / Event Report", Range("$P$2:$P" & OpenIRow), 6)
'  OpenType_INC_6 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Manufacturing Investigations / Incident", Range("$P$2:$P" & OpenIRow), 6)
'  OpenType_QAR_6 = Application.WorksheetFunction.CountIfs(Range("$I$2:$I" & OpenIRow), _
'  "Manufacturing Investigations / Quality Assurance Report (QAR)", Range("$P$2:$P" & OpenIRow), 6)
'  '------------------------------------------------
'  'Write Counts to Spreadsheet
'  '------------------------------------------------
'
'
'  Cells(1, OpenICol + 1).Value = "LIR"
'  Cells(1, OpenICol + 2).Value = "<23"
'  Cells(1, OpenICol + 3).Value = OpenType_LIR_0
'  Cells(2, OpenICol + 1).Value = "LIR"
'  Cells(2, OpenICol + 2).Value = "Aging"
'  Cells(2, OpenICol + 3).Value = OpenType_LIR_0_5
'  Cells(3, OpenICol + 1).Value = "LIR"
'  Cells(3, OpenICol + 2).Value = "31-60 days"
'  Cells(3, OpenICol + 3).Value = OpenType_LIR_1
'  Cells(4, OpenICol + 1).Value = "LIR"
'  Cells(4, OpenICol + 2).Value = "61-90 days"
'  Cells(4, OpenICol + 3).Value = OpenType_LIR_2
'  Cells(5, OpenICol + 1).Value = "LIR"
'  Cells(5, OpenICol + 2).Value = "91-120 days"
'  Cells(5, OpenICol + 3).Value = OpenType_LIR_3
'  Cells(6, OpenICol + 1).Value = "LIR"
'  Cells(6, OpenICol + 2).Value = "21-150 days"
'  Cells(6, OpenICol + 3).Value = OpenType_LIR_4
'  Cells(7, OpenICol + 1).Value = "LIR"
'  Cells(7, OpenICol + 2).Value = "151-180 days"
'  Cells(7, OpenICol + 3).Value = OpenType_LIR_5
'  Cells(8, OpenICol + 1).Value = "LIR"
'  Cells(8, OpenICol + 2).Value = ">180 days"
'  Cells(8, OpenICol + 3).Value = OpenType_LIR_6
'  'RAAC
'  Cells(9, OpenICol + 1).Value = "RAAC"
'  Cells(9, OpenICol + 2).Value = "<23"
'  Cells(9, OpenICol + 3).Value = OpenType_RAAC_0
'  Cells(10, OpenICol + 1).Value = "RAAC"
'  Cells(10, OpenICol + 2).Value = "Aging"
'  Cells(10, OpenICol + 3).Value = OpenType_RAAC_0_5
'  Cells(11, OpenICol + 1).Value = "RAAC"
'  Cells(11, OpenICol + 2).Value = "31-60 days"
'  Cells(11, OpenICol + 3).Value = OpenType_RAAC_1
'  Cells(12, OpenICol + 1).Value = "RAAC"
'  Cells(12, OpenICol + 2).Value = "61-90 days"
'  Cells(12, OpenICol + 3).Value = OpenType_RAAC_2
'  Cells(13, OpenICol + 1).Value = "RAAC"
'  Cells(13, OpenICol + 2).Value = "91-120 days"
'  Cells(13, OpenICol + 3).Value = OpenType_RAAC_3
'  Cells(14, OpenICol + 1).Value = "RAAC"
'  Cells(14, OpenICol + 2).Value = "121-150 days"
'  Cells(14, OpenICol + 3).Value = OpenType_RAAC_4
'  Cells(15, OpenICol + 1).Value = "RAAC"
'  Cells(15, OpenICol + 2).Value = "151-180 days"
'  Cells(15, OpenICol + 3).Value = OpenType_RAAC_5
'  Cells(16, OpenICol + 1).Value = "RAAC"
'  Cells(16, OpenICol + 2).Value = ">180 days"
'  Cells(16, OpenICol + 3).Value = OpenType_RAAC_6
'  'ER
'  Cells(17, OpenICol + 1).Value = "ER"
'  Cells(17, OpenICol + 2).Value = "<23"
'  Cells(17, OpenICol + 3).Value = OpenType_ER_0
'  Cells(18, OpenICol + 1).Value = "ER"
'  Cells(18, OpenICol + 2).Value = "Aging"
'  Cells(18, OpenICol + 3).Value = OpenType_ER_0_5
'  Cells(19, OpenICol + 1).Value = "ER"
'  Cells(19, OpenICol + 2).Value = "31-60 days"
'  Cells(19, OpenICol + 3).Value = OpenType_ER_1
'  Cells(20, OpenICol + 1).Value = "ER"
'  Cells(20, OpenICol + 2).Value = "61-90 days"
'  Cells(20, OpenICol + 3).Value = OpenType_ER_2
'  Cells(21, OpenICol + 1).Value = "ER"
'  Cells(21, OpenICol + 2).Value = "91-120 days"
'  Cells(21, OpenICol + 3).Value = OpenType_ER_3
'  Cells(22, OpenICol + 1).Value = "ER"
'  Cells(22, OpenICol + 2).Value = "121-150 days"
'  Cells(22, OpenICol + 3).Value = OpenType_ER_4
'  Cells(23, OpenICol + 1).Value = "ER"
'  Cells(23, OpenICol + 2).Value = "151-180 days"
'  Cells(23, OpenICol + 3).Value = OpenType_ER_5
'  Cells(24, OpenICol + 1).Value = "ER"
'  Cells(24, OpenICol + 2).Value = ">180 days"
'  Cells(24, OpenICol + 3).Value = OpenType_ER_6
'  'QAR
'  Cells(25, OpenICol + 1).Value = "QAR"
'  Cells(25, OpenICol + 2).Value = "<23"
'  Cells(25, OpenICol + 3).Value = OpenType_QAR_0
'  Cells(26, OpenICol + 1).Value = "QAR"
'  Cells(26, OpenICol + 2).Value = "Aging"
'  Cells(26, OpenICol + 3).Value = OpenType_QAR_0_5
'  Cells(27, OpenICol + 1).Value = "QAR"
'  Cells(27, OpenICol + 2).Value = "31-60 days"
'  Cells(27, OpenICol + 3).Value = OpenType_QAR_1
'  Cells(28, OpenICol + 1).Value = "QAR"
'  Cells(28, OpenICol + 2).Value = "61-90 days"
'  Cells(28, OpenICol + 3).Value = OpenType_QAR_2
'  Cells(29, OpenICol + 1).Value = "QAR"
'  Cells(29, OpenICol + 2).Value = "91-120 days"
'  Cells(29, OpenICol + 3).Value = OpenType_QAR_3
'  Cells(30, OpenICol + 1).Value = "QAR"
'  Cells(30, OpenICol + 2).Value = "121-150 days"
'  Cells(30, OpenICol + 3).Value = OpenType_QAR_4
'  Cells(31, OpenICol + 1).Value = "QAR"
'  Cells(31, OpenICol + 2).Value = "151-180 days"
'  Cells(31, OpenICol + 3).Value = OpenType_QAR_5
'  Cells(32, OpenICol + 1).Value = "QAR"
'  Cells(32, OpenICol + 2).Value = "> 180 days"
'  Cells(32, OpenICol + 3).Value = OpenType_QAR_6
'  'INC
'  Cells(33, OpenICol + 1).Value = "INC"
'  Cells(33, OpenICol + 2).Value = "<23"
'  Cells(33, OpenICol + 3).Value = OpenType_INC_0
'  Cells(34, OpenICol + 1).Value = "INC"
'  Cells(34, OpenICol + 2).Value = "Aging"
'  Cells(34, OpenICol + 3).Value = OpenType_INC_0_5
'  Cells(35, OpenICol + 1).Value = "INC"
'  Cells(35, OpenICol + 2).Value = "31-60 days"
'  Cells(35, OpenICol + 3).Value = OpenType_INC_1
'  Cells(36, OpenICol + 1).Value = "INC"
'  Cells(36, OpenICol + 2).Value = "61-90 days"
'  Cells(36, OpenICol + 3).Value = OpenType_INC_2
'  Cells(37, OpenICol + 1).Value = "INC"
'  Cells(37, OpenICol + 2).Value = "91-120 days"
'  Cells(37, OpenICol + 3).Value = OpenType_INC_3
'  Cells(38, OpenICol + 1).Value = "INC"
'  Cells(38, OpenICol + 2).Value = "121-150 days"
'  Cells(38, OpenICol + 3).Value = OpenType_INC_4
'  Cells(39, OpenICol + 1).Value = "INC"
'  Cells(39, OpenICol + 2).Value = "151-180 days"
'  Cells(39, OpenICol + 3).Value = OpenType_INC_5
'  Cells(40, OpenICol + 1).Value = "INC"
'  Cells(40, OpenICol + 2).Value = ">180 days"
'  Cells(40, OpenICol + 3).Value = OpenType_INC_6
End Sub
