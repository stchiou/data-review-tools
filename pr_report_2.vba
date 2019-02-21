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
    Dim File_1 As String
    Dim File_2 As String
    Dim week_num As Integer
    Dim OpenRecNum As Long
    Dim Sheet_Name As String
    Dim OpenCount() As Integer
    Dim OpenAge() As Integer
    Dim OpenStage() As Integer
    Dim OpenIRow As Integer
    Dim OpenICol As Integer
    Dim RecType() As Integer
    Dim CurRec() As Range
    Dim temp As String
    Dim tempval As Long
    Dim address1 As String
    Dim address2 As String
    Dim OpenRec() As String
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
   '---------------------------------------------------------------------------------
    File_1 = InputBox("Input filename and file extension of the data file to be processed")
    File_2 = InputBox("Input filename and file extension that contains short descriptions of the records")
    week_num = InputBox("Input week number of the year")
    Sheet_Name = Left(File_1, InStr(File_1, ".") - 1)
    Workbooks.OpenText Filename:="C:\Users\chious\Box Sync\vba-projects\pr-status\week" & week_num & "\" & File_1, local:=True
    Workbooks.OpenText Filename:="C:\Users\chious\Box Sync\vba-projects\pr-status\week" & week_num & "\" & File_2, local:=True
    Columns("E:E").Select
    Selection.Copy
    Windows(File_1).Activate
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Worksheets(Sheet_Name).Activate
    OpenIRow = Cells(1, 1).End(xlDown).Row
    OpenICol = Cells(1, 1).End(xlToRight).Column
    '------------------------------------------
    'Removing approved record
    '------------------------------------------
    For i = 2 To OpenIRow
        temp = Cells(i, 9).Value
        If InStr(temp, "Awaiting SQL Approval") > 0 Then
        Else
            If InStr(temp, "OPUQL") > 0 Then
            Else
                tempval = Cells(i, 6)
                If tempval > 0 Then
                    Rows(i).EntireRow.Delete
                    i = i - 1
                    OpenIRow = OpenIRow - 1
                Else
                    tempval = Cells(i, 7)
                    If tempval > 0 Then
                        Rows(i).EntireRow.Delete
                        i = i - 1
                        OpenIRow = OpenIRow - 1
                    Else
                    End If
                End If
            End If
        End If
    Next i
  '----------------------------------------
  'Calculate Age
  '----------------------------------------
  OpenRecNum = OpenIRow
  'OpenRecNum is the line number of the last line that contain open record;
  'Total Record Number = OpenRecNum -1
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
  
  For i = 2 To OpenRecNum
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
    temp = Cells(i, 11).Value
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
  'Array subscripts
  '------------------------------
  'Array Name: OpenRec
  '----------------------------------
  'First Dimension | Second Dimension
  '----------------------------------
  '2-OpenRecNum:   | 0: Record Number
  '                | 1: Short Description
  'Row on          | 2: Record Stage
  'spreadsheet     | 3: Record Type
  '----------------------------------
  'Array Name: OpenCount
  '-----------------------------------
  'First Dimension | Second Dimension
  '-----------------------------------
  '                | 0: < 30
  '1: LIR          | 1: 23-30
  '2: RAAC         | 2: 31-60
  '3: ER           | 3: 61-90
  '4: QAR          | 4: 91-120
  '5: INC          | 5: 121-150
  '6: stage total  | 6: 151-180
  '                | 7: >180
  '                | 8: Aged record Total
  '                | 9: Type Total
  '---------------------------------------------------------------------------------
  ReDim CurRec(5) As Range
  ReDim OpenRec(OpenRecNum, 3) As String
  ReDim OpenCount(6, 9) As Integer
  For i = 0 To 5
    For j = 0 To 9
        OpenCount(i, j) = 0
    Next j
  Next i
  For i = 2 To OpenIRow
    Select Case RecType(i)
        Case Is = 1
            Select Case OpenStage(i)
                Case Is = 0
                    OpenCount(1, 0) = OpenCount(1, 0) + 1
                Case Is = 1
                    OpenCount(1, 1) = OpenCount(1, 1) + 1
                Case Is = 2
                    OpenCount(1, 2) = OpenCount(1, 2) + 1
                Case Is = 3
                    OpenCount(1, 3) = OpenCount(1, 3) + 1
                Case Is = 4
                    OpenCount(1, 4) = OpenCount(1, 4) + 1
                Case Is = 5
                    OpenCount(1, 5) = OpenCount(1, 5) + 1
                Case Is = 6
                    OpenCount(1, 6) = OpenCount(1, 6) + 1
                Case Is = 7
                    OpenCount(1, 7) = OpenCount(1, 7) + 1
            End Select
        Case Is = 2
            Select Case OpenStage(i)
                Case Is = 0
                    OpenCount(2, 0) = OpenCount(2, 0) + 1
                Case Is = 1
                    OpenCount(2, 1) = OpenCount(2, 1) + 1
                Case Is = 2
                    OpenCount(2, 2) = OpenCount(2, 2) + 1
                Case Is = 3
                    OpenCount(2, 3) = OpenCount(2, 3) + 1
                Case Is = 4
                    OpenCount(2, 4) = OpenCount(2, 4) + 1
                Case Is = 5
                    OpenCount(2, 5) = OpenCount(2, 5) + 1
                Case Is = 6
                    OpenCount(2, 6) = OpenCount(2, 6) + 1
                Case Is = 7
                    OpenCount(2, 7) = OpenCount(2, 7) + 1
            End Select
        Case Is = 3
            Select Case OpenStage(i)
                Case Is = 0
                    OpenCount(3, 0) = OpenCount(3, 0) + 1
                Case Is = 1
                    OpenCount(3, 1) = OpenCount(3, 1) + 1
                Case Is = 2
                    OpenCount(3, 2) = OpenCount(3, 2) + 1
                Case Is = 3
                    OpenCount(3, 3) = OpenCount(3, 3) + 1
                Case Is = 4
                    OpenCount(3, 4) = OpenCount(3, 4) + 1
                Case Is = 5
                    OpenCount(3, 5) = OpenCount(3, 5) + 1
                Case Is = 6
                    OpenCount(3, 6) = OpenCount(3, 6) + 1
                Case Is = 7
                    OpenCount(3, 7) = OpenCount(3, 7) + 1
            End Select
        Case Is = 4
            Select Case OpenStage(i)
                Case Is = 0
                    OpenCount(4, 0) = OpenCount(4, 0) + 1
                Case Is = 1
                    OpenCount(4, 1) = OpenCount(4, 1) + 1
                Case Is = 2
                    OpenCount(4, 2) = OpenCount(4, 2) + 1
                Case Is = 3
                    OpenCount(4, 3) = OpenCount(4, 3) + 1
                Case Is = 4
                    OpenCount(4, 4) = OpenCount(4, 4) + 1
                Case Is = 5
                    OpenCount(4, 5) = OpenCount(4, 5) + 1
                Case Is = 6
                    OpenCount(4, 6) = OpenCount(4, 6) + 1
                Case Is = 7
                    OpenCount(4, 7) = OpenCount(4, 7) + 1
            End Select
        Case Is = 5
            Select Case OpenStage(i)
                Case Is = 0
                    OpenCount(5, 0) = OpenCount(5, 0) + 1
                Case Is = 1
                    OpenCount(5, 1) = OpenCount(5, 1) + 1
                Case Is = 2
                    OpenCount(5, 2) = OpenCount(5, 2) + 1
                Case Is = 3
                    OpenCount(5, 3) = OpenCount(5, 3) + 1
                Case Is = 4
                    OpenCount(5, 4) = OpenCount(5, 4) + 1
                Case Is = 5
                    OpenCount(5, 5) = OpenCount(5, 5) + 1
                Case Is = 6
                    OpenCount(5, 6) = OpenCount(5, 6) + 1
                Case Is = 7
                    OpenCount(5, 7) = OpenCount(5, 7) + 1
            End Select
    End Select
    OpenRec(i, 0) = Worksheets(Sheet_Name).Cells(i, 1).Value
    OpenRec(i, 1) = Worksheets(Sheet_Name).Cells(i, 3).Value
    OpenRec(i, 2) = OpenStage(i)
    OpenRec(i, 3) = RecType(i)
  Next i
  For i = 1 To 5
        OpenCount(i, 8) = OpenCount(i, 2) + OpenCount(i, 3) + OpenCount(i, 4) + OpenCount(i, 5) + OpenCount(i, 6) + OpenCount(i, 7)
        OpenCount(i, 9) = OpenCount(i, 0) + OpenCount(i, 1) + OpenCount(i, 8)
  Next i
  For i = 0 To 9
    OpenCount(6, i) = OpenCount(1, i) + OpenCount(2, i) + OpenCount(3, i) + OpenCount(4, i) + OpenCount(5, i)
  Next i
  
  Sheets.Add after:=Sheets(Sheet_Name)
  Sheets(Sheets.Count).Select
  Sheets(Sheets.Count).Name = "Results"
  Worksheets("Results").Cells(1, 1).Value = "Record Type"
  Worksheets("Results").Cells(1, 2).Value = "<23 Days"
  Worksheets("Results").Cells(1, 3).Value = "24-30 Days"
  Worksheets("Results").Cells(1, 4).Value = "31-60 Days"
  Worksheets("Results").Cells(1, 5).Value = "61-90 Days"
  Worksheets("Results").Cells(1, 6).Value = "91-120 Days"
  Worksheets("Results").Cells(1, 7).Value = "121-150 Days"
  Worksheets("Results").Cells(1, 8).Value = "151-180 Days"
  Worksheets("Results").Cells(1, 9).Value = ">181 Days"
  Worksheets("Results").Cells(1, 10).Value = "Aged"
  Worksheets("Results").Cells(1, 11).Value = "Total"
  Worksheets("Results").Cells(2, 1).Value = "LIR"
  Worksheets("Results").Cells(3, 1).Value = "RAAC"
  Worksheets("Results").Cells(4, 1).Value = "ER"
  Worksheets("Results").Cells(5, 1).Value = "QAR"
  Worksheets("Results").Cells(6, 1).Value = "INC"
  Worksheets("Results").Cells(7, 1).Value = "Total"
  For i = 1 To 6
    For j = 0 To 9
        Cells(i + 1, j + 2).Value = OpenCount(i, j)
    Next j
  Next i
  OpenICol = Cells(1, 1).End(xlToRight).Column
  For i = 0 To 4
    Worksheets("Results").Cells(1, OpenICol + 4 * i + 1).Value = "Record ID"
    Worksheets("Results").Cells(1, OpenICol + 4 * i + 2).Value = "Short Description"
    Worksheets("Results").Cells(1, OpenICol + 4 * i + 3).Value = "Record Stage"
    Worksheets("Results").Cells(1, OpenICol + 4 * i + 4).Value = "Record Type"
  Next i
  CurRec(1) = Cells(2, OpenICol)
  CurRec(2) = CurRec(1).Offset(0, 4)
  CurRec(3) = CurRec(1).Offset(0, 8)
  CurRec(4) = CurRec(1).Offset(0, 12)
  CurRec(5) = CurRec(1).Offset(0, 16)
  For i = 2 To OpenRecNum
    If OpenRec(i, 3) = 1 Then
        CurRec(1).Value = OpenRec(i, 0)
        CurRec(1).Offset(0, 1).Value = OpenRec(i, 1)
        CurRec(1).Offset(0, 2).Value = OpenRec(i, 2)
        CurRec(1).Offset(0, 3).Value = OpenRec(i, 3)
        CurRec(1) = CurRec(1).Offset(1, 0)
    Else
        If OpenRec(i, 3) = 2 Then
            CurRec(2).Value = OpenRec(i, 0)
            CurRec(2).Offset(0, 1).Value = OpenRec(i, 1)
            CurRec(2).Offset(0, 2).Value = OpenRec(i, 2)
            CurRec(2).Offset(0, 3).Value = OpenRec(i, 3)
            CurRec(2) = CurRec(2).Offset(1, 0)
        Else
            If OpenRec(i, 3) = 3 Then
                CurRec(3).Value = OpenRec(i, 0)
                CurRec(3).Offset(0, 1).Value = OpenRec(i, 1)
                CurRec(3).Offset(0, 2).Value = OpenRec(i, 2)
                CurRec(3).Offset(0, 3).Value = OpenRec(i, 3)
                CurRec(3) = CurRec(3).Offset(1, 0)
            Else
                If OpenRec(i, 3) = 4 Then
                    CurRec(4).Value = OpenRec(i, 0)
                    CurRec(4).Offset(0, 1).Value = OpenRec(i, 1)
                    CurRec(4).Offset(0, 2).Value = OpenRec(i, 2)
                    CurRec(4).Offset(0, 3).Value = OpenRec(i, 3)
                    CurRec(4) = CurRec(4).Offset(1, 0)
                Else
                End If
            End If
        End If
    End If
  Next i
'Cells(1, 1).Activate
'  ActiveCell.EntireRow.Insert
'  address1 = Cells(1, 1).Address(rowabsolute:=False, columnabsolute:=False)
'  address2 = Cells(1, OpenICol).Address(rowabsolute:=False, columnabsolute:=False)
'  Range(address1 & ":" & address2).Select
'  Selection.Merge
End Sub
