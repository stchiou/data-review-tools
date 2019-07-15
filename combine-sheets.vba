Attribute VB_Name = "Module1"
Sub CountFiles()
Dim xFolder As String
Dim xPath As String
Dim xCount As Long
Dim xFiDialog As FileDialog
Dim xFile As String
Dim FName(1000) As String
Dim SName(1000) As String
Set xFiDialog = Application.FileDialog(msoFileDialogFolderPicker)
If xFiDialog.Show = -1 Then
xFolder = xFiDialog.SelectedItems(1)
End If
If xFolder = "" Then Exit Sub
xPath = xFolder & "\*.csv"
xFile = Dir(xPath)
Do While xFile <> ""
    xCount = xCount + 1
    xFile = Dir()
    FName(xCount) = xFile
    If Len(FName(xCount)) <> 0 Then
        SName(xCount) = Left(FName(xCount), Len(FName(xCount)) - 4)
        Workbooks.Open FileName:=xFolder & "\" & FName(xCount)
        Sheets(SName(xCount)).Select
        Sheets(SName(xCount)).Move Before:=Workbooks("Book1").Sheets(1)
    Else: Exit Sub
    End If
Loop
End Sub
