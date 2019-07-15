Attribute VB_Name = "Module1"
Sub CountFiles()
Dim xFolder As String
Dim xPath As String
Dim xCount As Long
Dim xFiDialog As FileDialog
Dim xFile As String
Dim FileName(200) As String
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
    FileName(xCount) = xFile
Loop

MsgBox xCount & " files found"
End Sub
