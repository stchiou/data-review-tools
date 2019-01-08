Attribute VB_Name = "Module1"
Sub Data_Review()
    Dim LastRow As Integer
    Dim curRow As Integer
    Dim i As Integer
    Dim j As Integer
    Dim col_g(100) As String
    Dim nb(100) As Integer
    Dim pg(100) As Integer
    Dim col_j(100) As String
    Dim col_m(100) As String
    Dim p1 As Integer
    Dim p2 As Integer
    Dim p_rev as String
    Dim rev as String
    Dim rel as String
    'Function for finding Data Reviewer name'
    Function FindRole (comment)
      Dim comment as string
      Dim rev_start as Integer
      Dim rel_start as Integer
      Dim separator as Integer
      separator=Instr(comment,"     ")
      'Find Data Reviewer'
      rev_start=Instr(comment,"Data reviewer ")
      If rev_start=0 Then
        comment=left(comment,separator+5,len(comment)-separator-5)
      Else
        If rev_start > separator Then
          comment=left(comment,separator+5,len(comment)-separator-5)
          separator=Instr(comment,"     ")
        Else
          comment=left(comment,separator+5,len(comment)-separator-5)
        End If
          rev=mid(comment,rev_start+14,10)
      End If
      'Find Releaser'
      rel_start=Inst()
    End Function
    Function FindRel ()
    End Function
    Function FindPR()
    End Function
    'Create a new sheet for consolidated data'
    Sheets.Add after:=Sheets("supplement")
    Sheets(Sheets.Count).Select
    Sheets(Sheets.Count).Name = "Data"
    'Remove extra blank lines on the spreadsheet'
    On Error Resume Next
        Range(1, 1).Select
        Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    'Determine number of record in the sheet'
    Worksheets("QA Data").Select
    LastRow = Cells(1, 1).End(xlDown).Row
    'Make a copy of Date and Method from the source sheet to the target sheet'
    Worksheets("QA Data").Range(Cells(1, 5), Cells(LastRow, 5)).Copy _
    Destination:=Worksheets("Data").Range("A1")
    Worksheets("QA Data").Range(Cells(1, 12), Cells(LastRow, 12)).Copy _
    Destination:=Worksheets("Data").Range("B1")
    Worksheets("QA Data").Range(Cells(1, 6), Cells(LastRow, 6)).Copy _
    Destination:=Worksheets("Data").Range("E1")
    Worksheets("QA Data").Range(Cells(1, 8), Cells(LastRow, 8)).Copy _
    Destination:=Worksheets("Data").Range("F1")
    'Create Headers for Notebook and Page Number columns'
    Worksheets("Data").Cells(1, 3).Value = "Note Book"
    Worksheets("Data").Cells(1, 4).Value = "Page"
    'Parse Notebook and Page Number from the source sheet to the target sheet'
    For i = 2 To LastRow
        Cells(i, 7).Select
        col_g(i) = Cells(i, 7).Value
        p1 = InStr(col_g(i), "Book ")
        p2 = InStr(col_g(i), "page ")
        nb(i) = Mid(Cells(i, 7).Value, p1 + 5, 5)
        pg(i) = Mid(Cells(i, 7).Value, p2 + 5, 2)
        Worksheets("Data").Cells(i, 3).Value = nb(i)
        Worksheets("Data").Cells(i, 4).Value = pg(i)
        col_j(i) = cells(i, 10).Value

    Next i







End Sub
