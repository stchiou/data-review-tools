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
'Features:
'1. Combine output records with corresponding short description
'2. Computes age of the records
'3. Computes stage of the records based on age
'4. Generate reports
'------------------------------------------------------------------------------------------------------------------
Dim File_1 As String
Dim File_2 As String
Dim File_3 As String
Dim File_4 As String
Dim week_num As Integer
Dim OpenRecNum As Long
Dim OpenSheet_Name As String
Dim OpenCount() As Integer
Dim OpenAge() As Integer
Dim OpenStage() As Integer
Dim OpenlRow As Integer
Dim OpenlCol As Integer
Dim OpenRecType() As Integer
Dim OpenCurRec() As Integer
Dim temp As String
Dim tempval As Long
Dim OpenRec() As String
Dim CloseRecNum As Long
Dim CloselRow As Integer
Dim CloselCol As Integer
Dim CloseCount() As Integer
Dim CloseAge() As Integer
Dim CloseStage() As Integer
Dim CloseRecType() As Integer
Dim CloseCurRec() As Integer
Dim CloseRec() As String
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
week_num = InputBox("Input week number of the year", "WEEK NUMBER")
File_1 = InputBox("Input filename and file extension of the Open Records data file of week " & week_num & " to be processed", "OPEN RECORDS")
File_2 = InputBox("Input filename and file extension of the file contains short descriptions of the Open Records of week " & week_num, "OPEN RECORDS SHORT DESCRIPTION")
File_3 = InputBox("Input filename and file extension of the Closed Records data file of week " & week_num & " to be processed", "CLOSED RECORDS")
File_4 = InputBox("Input filename and file extension of the file contains short descriptions of the Closed Record of week " & week_num, "CLOSED RECORDS SHORT DESCRIPTION")
OpenSheet_Name = Left(File_1, InStr(File_1, ".") - 1)
'--------------------------------------------------------------------------------
'Combine Short Description to Record File for Open Records
'--------------------------------------------------------------------------------
Workbooks.OpenText Filename:="C:\Users\chious\Box Sync\vba-projects\pr-status\week" & week_num & "\" & File_1, local:=True
Workbooks.OpenText Filename:="C:\Users\chious\Box Sync\vba-projects\pr-status\week" & week_num & "\" & File_2, local:=True
Columns("E:E").Select
Selection.Copy
Windows(File_1).Activate
Columns("C:C").Select
Selection.Insert Shift:=xlToRight
Worksheets(OpenSheet_Name).Activate
OpenlRow = Cells(1, 1).End(xlDown).Row
OpenlCol = Cells(1, 1).End(xlToRight).Column
'------------------------------------------------------------------------------
'Removing approved record on Open Records
'------------------------------------------------------------------------------
For i = 2 To OpenlRow
      temp = Cells(i, 9).Value
      If InStr(temp, "Awaiting SQL Approval") > 0 Then
      Else
          If InStr(temp, "OPUQL") > 0 Then
          Else
              tempval = Cells(i, 6)
              If tempval > 0 Then
                  Rows(i).EntireRow.Delete
                  i = i - 1
                  OpenlRow = OpenlRow - 1
              Else
                  tempval = Cells(i, 7)
                  If tempval > 0 Then
                      Rows(i).EntireRow.Delete
                      i = i - 1
                      OpenlRow = OpenlRow - 1
                  Else
                  End If
              End If
          End If
      End If
  Next i
'---------------------------------------------------------------------------------
'Calculate Age of Open Records
'---------------------------------------------------------------------------------
OpenRecNum = OpenlRow
'OpenRecNum is the line number of the last line that contain open record;
'Total Record Number = OpenRecNum -1
Cells(1, OpenlCol).Value = "Age"
ReDim OpenAge(OpenlRow) As Integer
ReDim OpenStage(OpenlRow) As Integer
ReDim OpenRecType(OpenlRow) As Integer
For i = 2 To OpenlRow
  OpenAge(i) = Date - Cells(i, 4)
  Cells(i, OpenlCol).Value = OpenAge(i)
Next i
Range(Cells(2, OpenlCol), Cells(OpenlRow, OpenlCol)).NumberFormat = "0"
OpenlCol = OpenlCol + 1
'--------------------------------------------------------------------------------
'Assign Stage of Open Record Based on Age of the Records
'Assign Type of Records in Numerical Formats
'--------------------------------------------------------------------------------
Cells(1, OpenlCol).Value = "Stage"
Cells(1, OpenlCol + 1).Value = "Type"
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
  Cells(i, OpenlCol).Value = OpenStage(i)
  Cells(i, OpenlCol + 1).Value = OpenRecType(i)
Next i
OpenlCol = OpenlCol + 2
'--------------------------------------------------------------------------------
'Computing Open Records by Age and Type then Store the Results in Array
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
ReDim OpenCurRec(1, 5) As Integer
ReDim OpenRec(OpenRecNum, 3) As String
ReDim OpenCount(6, 9) As Integer
'--------------------------------------------------
'Reset the Array that Store Open Record Counts
'--------------------------------------------------
For i = 0 To 5
  For j = 0 To 9
      OpenCount(i, j) = 0
  Next j
Next i
'---------------------------------------------------------------------------------------------
'Assigning Record Type, Counting Numbers of Each Record Type,
'and Storing Results in an Array
'---------------------------------------------------------------------------------------------
For i = 2 To OpenlRow
  Select Case OpenRecType(i)
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
    '-------------------------------------------------------------
    'Capture Open Records Information into Array
    '-------------------------------------------------------------
  OpenRec(i, 0) = Worksheets(OpenSheet_Name).Cells(i, 1).Value
  OpenRec(i, 1) = Worksheets(OpenSheet_Name).Cells(i, 3).Value
  OpenRec(i, 2) = OpenStage(i)
  OpenRec(i, 3) = OpenRecType(i)
Next i
'--------------------------------------------------------------
'Compute Subtotal and Grand Total of the Open Records Matrix
'--------------------------------------------------------------
For i = 1 To 5
      OpenCount(i, 8) = OpenCount(i, 2) + OpenCount(i, 3) + OpenCount(i, 4) + OpenCount(i, 5) + OpenCount(i, 6) + OpenCount(i, 7)
      OpenCount(i, 9) = OpenCount(i, 0) + OpenCount(i, 1) + OpenCount(i, 8)
Next i
For i = 0 To 9
  OpenCount(6, i) = OpenCount(1, i) + OpenCount(2, i) + OpenCount(3, i) + OpenCount(4, i) + OpenCount(5, i)
Next i
'----------------------------------------------------------------
'Generate Summary Report
'----------------------------------------------------------------
Sheets.Add after:=Sheets(OpenSheet_Name)
Sheets(Sheets.Count).Select
Sheets(Sheets.Count).Name = "Week_" & week_num
'----------------------------------------------------------------
'Create Headers Row and Column of the Report
'----------------------------------------------------------------
Worksheets("Week_" & week_num).Cells(1, 1).Value = "Record Type"
Worksheets("Week_" & week_num).Cells(1, 2).Value = "<23 Days"
Worksheets("Week_" & week_num).Cells(1, 3).Value = "24-30 Days"
Worksheets("Week_" & week_num).Cells(1, 4).Value = "31-60 Days"
Worksheets("Week_" & week_num).Cells(1, 5).Value = "61-90 Days"
Worksheets("Week_" & week_num).Cells(1, 6).Value = "91-120 Days"
Worksheets("Week_" & week_num).Cells(1, 7).Value = "121-150 Days"
Worksheets("Week_" & week_num).Cells(1, 8).Value = "151-180 Days"
Worksheets("Week_" & week_num).Cells(1, 9).Value = ">181 Days"
Worksheets("Week_" & week_num).Cells(1, 10).Value = "Aged"
Worksheets("Week_" & week_num).Cells(1, 11).Value = "Total"
Worksheets("Week_" & week_num).Cells(2, 1).Value = "LIR"
Worksheets("Week_" & week_num).Cells(3, 1).Value = "RAAC"
Worksheets("Week_" & week_num).Cells(4, 1).Value = "ER"
Worksheets("Week_" & week_num).Cells(5, 1).Value = "QAR"
Worksheets("Week_" & week_num).Cells(6, 1).Value = "INC"
Worksheets("Week_" & week_num).Cells(7, 1).Value = "Total"
'----------------------------------------------------------------
'Writing Open Record Matrix
'----------------------------------------------------------------
For i = 1 To 6
  For j = 0 To 9
      Cells(i + 1, j + 2).Value = OpenCount(i, j)
  Next j
Next i
'-------------------------------------------------------------
'Update the number of Non-Empty Columns in the Summary Report
'-------------------------------------------------------------
ReplCol = Cells(1, 1).End(xlToRight).Column
'---------------------------------------------------------------
'Generate Headers for Details Section of the Summary Report
'--------------------------------------------------------------
For i = 0 To 4
  Worksheets("Week_" & week_num).Cells(1, ReplCol + 4 * i + 1).Value = "Record ID"
  Worksheets("Week_" & week_num).Cells(1, ReplCol + 4 * i + 2).Value = "Short Description"
  Worksheets("Week_" & week_num).Cells(1, ReplCol + 4 * i + 3).Value = "Record Stage"
  Worksheets("Week_" & week_num).Cells(1, ReplCol + 4 * i + 4).Value = "Record Type"
Next i
'-----------------------------------------------------------------------------------------------
'Create Array to Capture Positions of Where Each Record Being Output in the Summary Spreadsheet
'-----------------------------------------------------------------------------------------------
OpenCurRec(0, 1) = 2
OpenCurRec(1, 1) = ReplCol + 1
OpenCurRec(0, 2) = 2
OpenCurRec(1, 2) = ReplCol + 5
OpenCurRec(0, 3) = 2
OpenCurRec(1, 3) = ReplCol + 9
OpenCurRec(0, 4) = 2
OpenCurRec(1, 4) = ReplCol + 13
OpenCurRec(0, 5) = 2
OpenCurRec(1, 5) = ReplCol + 17
'----------------------------------------------------------------------------------
'Writing Detail Information of Open Records from Array into Spreadsheet while
'Updating Array that Captured Position of each Record in the Spreadsheet
'----------------------------------------------------------------------------------
For i = 2 To OpenRecNum
  If OpenRec(i, 3) = 1 Then
    Cells(OpenCurRec(0, 1), OpenCurRec(1, 1)).Activate
    ActiveCell.Value = OpenRec(i, 0)
    ActiveCell.Offset(0, 1).Value = OpenRec(i, 1)
    ActiveCell.Offset(0, 2).Value = OpenRec(i, 2)
    ActiveCell.Offset(0, 3).Value = OpenRec(i, 3)
    OpenCurRec(0, 1) = OpenCurRec(0, 1) + 1
    OpenCurRec(1, 1) = OpenCurRec(1, 1)
  Else
    If OpenRec(i, 3) = 2 Then
        Cells(OpenCurRec(0, 2), OpenCurRec(1, 2)).Activate
        ActiveCell.Value = OpenRec(i, 0)
        ActiveCell.Offset(0, 1).Value = OpenRec(i, 1)
        ActiveCell.Offset(0, 2).Value = OpenRec(i, 2)
        ActiveCell.Offset(0, 3).Value = OpenRec(i, 3)
        OpenCurRec(0, 2) = OpenCurRec(0, 2) + 1
        OpenCurRec(1, 2) = OpenCurRec(1, 2)
    Else
        If OpenRec(i, 3) = 3 Then
            Cells(OpenCurRec(0, 3), OpenCurRec(1, 3)).Activate
            ActiveCell.Value = OpenRec(i, 0)
            ActiveCell.Offset(0, 1).Value = OpenRec(i, 1)
            ActiveCell.Offset(0, 2).Value = OpenRec(i, 2)
            ActiveCell.Offset(0, 3).Value = OpenRec(i, 3)
            OpenCurRec(0, 3) = OpenCurRec(0, 3) + 1
            OpenCurRec(1, 3) = OpenCurRec(1, 3)
        Else
            If OpenRec(i, 3) = 4 Then
                Cells(OpenCurRec(0, 4), OpenCurRec(1, 4)).Activate
                ActiveCell.Value = OpenRec(i, 0)
                ActiveCell.Offset(0, 1).Value = OpenRec(i, 1)
                ActiveCell.Offset(0, 2).Value = OpenRec(i, 2)
                ActiveCell.Offset(0, 3).Value = OpenRec(i, 3)
                OpenCurRec(0, 4) = OpenCurRec(0, 4) + 1
                OpenCurRec(1, 4) = OpenCurRec(1, 4)
            Else
                If OpenRec(i, 3) = 5 Then
                    Cells(OpenCurRec(0, 5), OpenCurRec(1, 5)).Activate
                    ActiveCell.Value = OpenRec(i, 0)
                    ActiveCell.Offset(0, 1).Value = OpenRec(i, 1)
                    ActiveCell.Offset(0, 2).Value = OpenRec(i, 2)
                    ActiveCell.Offset(0, 3).Value = OpenRec(i, 3)
                    OpenCurRec(0, 5) = OpenCurRec(0, 5) + 1
                    OpenCurRec(1, 5) = OpenCurRec(1, 5)
                Else
                End If
            End If
        End If
    End If
  End If
Next i
ReplCol = Worksheets("Week_" & week_num).Cells(1, 1).End(xlToRight).Column
'--------------------------------------------------------------------------
'Open Files Contains Closed Records and Short Description of Closed Records
'Insert Short Descriptions to the Sheet that Contains Closed Records
'--------------------------------------------------------------------------
CloseSheet_Name = Left(File_3, InStr(File_3, ".") - 1)
Workbooks.OpenText Filename:="C:\Users\chious\Box Sync\vba-projects\pr-status\week" & week_num & "\" & File_3, local:=True
Workbooks.OpenText Filename:="C:\Users\chious\Box Sync\vba-projects\pr-status\week" & week_num & "\" & File_4, local:=True
Columns("E:E").Select
Selection.Copy
Windows(File_3).Activate
Columns("C:C").Select
Selection.Insert Shift:=xlToRight
Worksheets(CloseSheet_Name).Activate
CloselRow = Cells(1, 1).End(xlDown).Row
CloselCol = Cells(1, 1).End(xlToRight).Column
'----------------------------------------
'Calculate Age of the Closed Records
'----------------------------------------
CloseRecNum = CloselRow
'CloseRecNum is the line number of the last line that contain close record;
'Total closed Record Number = CloseRecNum -1
Cells(1, CloselCol).Value = "Age"
ReDim CloseAge(CloselRow) As Integer
ReDim CloseStage(CloselRow) As Integer
ReDim CloseRecType(CloselRow) As Integer
For i = 2 To CloselRow
  CloseAge(i) = Date - Cells(i, 4)
  Cells(i, CloselCol).Value = CloseAge(i)
Next i
Range(Cells(2, CloselCol), Cells(CloselRow, CloselCol)).NumberFormat = "0"
CloselCol = CloselCol + 1
'----------------------------------------
'create category
'----------------------------------------
Cells(1, CloselCol).Value = "Stage"
Cells(1, CloselCol + 1).Value = "Type"
For i = 2 To CloseRecNum
      If CloseAge(i) > 30 Then
          CloseStage(i) = 1
      Else
          If CloseAge(i) <= 30 Then
              CloseStage(i) = 0
          Else
          End If
      End If
  temp = Cells(i, 11).Value
  Select Case temp
      Case "Laboratory Investigations / Laboratory Investigation Report (LIR)"
          CloseRecType(i) = 1
      Case "Laboratory Investigations / Readily Apparent Assignable Cause (RAAC)"
          CloseRecType(i) = 2
      Case "Manufacturing Investigations / Event Report"
          CloseRecType(i) = 3
      Case "Manufacturing Investigations / Quality Assurance Report (QAR)"
          CloseRecType(i) = 4
      Case "Manufacturing Investigations / Incident"
          CloseRecType(i) = 5
  End Select
  Cells(i, CloselCol).Value = CloseStage(i)
  Cells(i, CloselCol + 1).Value = CloseRecType(i)
Next i
CloselCol = CloselCol + 2
ReDim CloseCount(6, 2) As Integer
ReDim CloseRec(CloseRecNum, 3) As String
ReDim CloseCurRec(1, 5) As Integer
For i = 0 To 6
  For j = 0 To 2
      CloseCount(i, j) = 0
  Next j
Next i
For i = 2 To CloselRow
  Select Case CloseRecType(i)
      Case Is = 1
          Select Case CloseStage(i)
              Case Is = 0
                  CloseCount(1, 0) = CloseCount(1, 0) + 1
              Case Is = 1
                  CloseCount(1, 1) = CloseCount(1, 1) + 1
          End Select
      Case Is = 2
          Select Case CloseStage(i)
              Case Is = 0
                  CloseCount(2, 0) = CloseCount(2, 0) + 1
              Case Is = 1
                  CloseCount(2, 1) = CloseCount(2, 1) + 1
          End Select
      Case Is = 3
          Select Case CloseStage(i)
              Case Is = 0
                  CloseCount(3, 0) = CloseCount(3, 0) + 1
              Case Is = 1
                  CloseCount(3, 1) = CloseCount(3, 1) + 1
          End Select
      Case Is = 4
          Select Case CloseStage(i)
              Case Is = 0
                  CloseCount(4, 0) = CloseCount(4, 0) + 1
              Case Is = 1
                  CloseCount(4, 1) = CloseCount(4, 1) + 1
          End Select
      Case Is = 5
          Select Case CloseStage(i)
              Case Is = 0
                  CloseCount(5, 0) = CloseCount(5, 0) + 1
              Case Is = 1
                  CloseCount(5, 1) = CloseCount(5, 1) + 1
          End Select
  End Select
  CloseRec(i, 0) = Worksheets(CloseSheet_Name).Cells(i, 1).Value
  CloseRec(i, 1) = Worksheets(CloseSheet_Name).Cells(i, 3).Value
  CloseRec(i, 2) = CloseStage(i)
  CloseRec(i, 3) = CloseRecType(i)
Next i
For i = 1 To 5
  CloseCount(i, 2) = CloseCount(i, 0) + CloseCount(i, 1)
Next i
For i = 0 To 2
  CloseCount(6, i) = CloseCount(1, i) + CloseCount(2, i) + CloseCount(3, i) + CloseCount(4, i) + CloseCount(5, i)
Next i
'---------------------------------------------------------------------------
ReplCol = ReplCol + 1
Windows(File_1).Activate
Worksheets("Week_" & week_num).Cells(1, ReplCol).Activate
ActiveCell.Value = "Recod Type"
ActiveCell.Offset(0, 1).Value = "On Time"
ActiveCell.Offset(0, 2).Value = "Aged"
ActiveCell.Offset(0, 3).Value = "Total"
ActiveCell.Offset(1, 0).Value = "LIR"
ActiveCell.Offset(2, 0).Value = "RAAC"
ActiveCell.Offset(3, 0).Value = "ER"
ActiveCell.Offset(4, 0).Value = "QAR"
ActiveCell.Offset(5, 0).Value = "INC"
ActiveCell.Offset(6, 0).Value = "Total"
For i = 1 To 6
  For j = 0 To 2
      ActiveCell.Offset(i, j + 1).Offset.Value = CloseCount(i, j)
  Next
Next i
ReplCol = Cells(1, 1).End(xlToRight).Column + 1
For i = 0 To 4
  Worksheets("Week_" & week_num).Cells(1, ReplCol + 4 * i).Value = "Record ID"
  Worksheets("Week_" & week_num).Cells(1, ReplCol + 4 * i + 1).Value = "Short Description"
  Worksheets("Week_" & week_num).Cells(1, ReplCol + 4 * i + 2).Value = "Record Stage"
  Worksheets("Week_" & week_num).Cells(1, ReplCol + 4 * i + 3).Value = "Record Type"
Next i
CloseCurRec(0, 1) = 2
CloseCurRec(1, 1) = ReplCol
CloseCurRec(0, 2) = 2
CloseCurRec(1, 2) = CloseCurRec(1, 1) + 4
CloseCurRec(0, 3) = 2
CloseCurRec(1, 3) = CloseCurRec(1, 2) + 4
CloseCurRec(0, 4) = 2
CloseCurRec(1, 4) = CloseCurRec(1, 3) + 4
CloseCurRec(0, 5) = 2
CloseCurRec(1, 5) = CloseCurRec(1, 4) + 4
For i = 2 To CloseRecNum
  If CloseRec(i, 3) = 1 Then
      Cells(CloseCurRec(0, 1), CloseCurRec(1, 1)).Activate
      ActiveCell.Value = CloseRec(i, 0)
      ActiveCell.Offset(0, 1).Value = CloseRec(i, 1)
      ActiveCell.Offset(0, 2).Value = CloseRec(i, 2)
      ActiveCell.Offset(0, 3).Value = CloseRec(i, 3)
      CloseCurRec(0, 1) = CloseCurRec(0, 1) + 1
      CloseCurRec(1, 1) = CloseCurRec(1, 1)
  Else
      If CloseRec(i, 3) = 2 Then
          Cells(CloseCurRec(0, 2), CloseCurRec(1, 2)).Activate
          ActiveCell.Value = CloseRec(i, 0)
          ActiveCell.Offset(0, 1).Value = CloseRec(i, 1)
          ActiveCell.Offset(0, 2).Value = CloseRec(i, 2)
          ActiveCell.Offset(0, 3).Value = CloseRec(i, 3)
          CloseCurRec(0, 2) = CloseCurRec(0, 2) + 1
          CloseCurRec(1, 2) = CloseCurRec(1, 2)
      Else
          If CloseRec(i, 3) = 3 Then
              Cells(CloseCurRec(0, 3), CloseCurRec(1, 3)).Activate
              ActiveCell.Value = CloseRec(i, 0)
              ActiveCell.Offset(0, 1).Value = CloseRec(i, 1)
              ActiveCell.Offset(0, 2).Value = CloseRec(i, 2)
              ActiveCell.Offset(0, 3).Value = CloseRec(i, 3)
              CloseCurRec(0, 3) = CloseCurRec(0, 3) + 1
              CloseCurRec(1, 3) = CloseCurRec(1, 3)
          Else
              If CloseRec(i, 3) = 4 Then
                  Cells(CloseCurRec(0, 4), CloseCurRec(1, 4)).Activate
                  ActiveCell.Value = CloseRec(i, 0)
                  ActiveCell.Offset(0, 1).Value = CloseRec(i, 1)
                  ActiveCell.Offset(0, 2).Value = CloseRec(i, 2)
                  ActiveCell.Offset(0, 3).Value = CloseRec(i, 3)
                  CloseCurRec(0, 4) = CloseCurRec(0, 4) + 1
                  CloseCurRec(1, 4) = CloseCurRec(1, 4)
              Else
                  If CloseRec(i, 3) = 5 Then
                      Cells(CloseCurRec(0, 5), CloseCurRec(1, 5)).Activate
                      ActiveCell.Value = CloseRec(i, 0)
                      ActiveCell.Offset(0, 1).Value = CloseRec(i, 1)
                      ActiveCell.Offset(0, 2).Value = CloseRec(i, 2)
                      ActiveCell.Offset(0, 3).Value = CloseRec(i, 3)
                      CloseCurRec(0, 5) = CloseCurRec(0, 5) + 1
                      CloseCurRec(1, 5) = CloseCurRec(1, 5)
                  Else
                  End If
              End If
          End If
      End If
  End If
Next i
Worksheets("Week_" & week_num).Cells(1, 1).Activate
ActiveCell.EntireRow.Insert
Cells(1, 1).Value = "Open Records"
Cells(1, 12).Value = "Open LIR"
Cells(1, 16).Value = "Open RAAC"
Cells(1, 20).Value = "Open ER"
Cells(1, 24).Value = "Open QAR"
Cells(1, 28).Value = "Open INC"
Cells(1, 32).Value = "Closed Records"
Cells(1, 36).Value = "Closed LIR"
Cells(1, 40).Value = "Closed RAAC"
Cells(1, 44).Value = "Closed ER"
Cells(1, 48).Value = "Closed QAR"
Cells(1, 52).Value = "Closed INC"
address_1 = Cells(1, 1).Address(rowabsolute:=False, columnabsolute:=False)
address_2 = Cells(1, 11).Address(rowabsolute:=False, columnabsolute:=False)
Range(address_1 & ":" & address_2).Select
Selection.Merge
For i = 3 To 13
    address_1 = Cells(1, 4 * i).Address(rowabsolute:=False, columnabsolute:=False)
    address_2 = Cells(1, 4 * i + 3).Address(rowabsolute:=False, columnabsolute:=False)
    Range(address_1 & ":" & address_2).Select
    Selection.Merge
Next i

'---------------------------------------------------------------
'Saving Weekly Report
'---------------------------------------------------------------
Sheets("Week_" & week_num).Move
Worksheets("Week_" & week_num).Activate
ActiveWorkbook.SaveAs Filename:="Week_" & week_num & "_summary"


'-----------------------------------------------------------------
'Output Weekly Results
'-----------------------------------------------------------------

End Sub
