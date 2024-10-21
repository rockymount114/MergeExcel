Set fso = CreateObject("Scripting.FileSystemObject")
sFolderPath = GetFolderPath()
sFilePath = sFolderPath & "\MergeExcel.txt"

If fso.FileExists(sFilePath) = False Then
  MsgBox "Could not file configuration file: " & sFilePath
  WScript.Quit
End If

Dim oExcel: Set oExcel = CreateObject("Excel.Application")
oExcel.Visible = True
oExcel.DisplayAlerts = false
Set oMasterWorkbook = oExcel.Workbooks.Add()
Set oMasterSheet = oMasterWorkbook.Worksheets("Sheet1")
oMasterSheet.Name = "temp_delete"
oMasterWorkbook.Worksheets("Sheet2").Delete
oMasterWorkbook.Worksheets("Sheet3").Delete

Set oFile = fso.OpenTextFile(sFilePath, 1)   
Do until oFile.AtEndOfStream
  sFilePath = Replace(oFile.ReadLine,"""","")
  
  If fso.FileExists(sFilePath) Then
    Set oWorkBook = oExcel.Workbooks.Open(sFilePath)
    
    For Each oSheet in oWorkBook.Worksheets
      oSheet.Copy oMasterSheet
      'oSht.Move , oSheet
    Next
    
    oWorkBook.Close()
  End If
Loop
oFile.Close

oMasterSheet.Delete
MsgBox "Done"
          
Function GetFolderPath()
	Dim oFile 'As Scripting.File
	Set oFile = fso.GetFile(WScript.ScriptFullName)
	GetFolderPath = oFile.ParentFolder
End Function
