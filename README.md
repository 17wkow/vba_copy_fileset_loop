# vba_copy_fileset_loop

Sub LoopAllExcelFilesInFolder()
'Loop through all Excel files in a user specified folder and perform a set task on them

Dim wb As Workbook
Dim myPath As String
Dim myfile As String
Dim myExtension As String
Dim FldrPicker As FileDialog

'Optimize Macro Speed
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Application.Calculation = xlCalculationManual

'Retrieve Target Folder Path From User
  Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

    With FldrPicker
      .Title = "Select A Target Folder"
      .AllowMultiSelect = False
        If .Show <> -1 Then GoTo NextCode
        myPath = .SelectedItems(1) & "\"
    End With

'In Case of Cancel
NextCode:
  myPath = myPath
  If myPath = "" Then GoTo ResetSettings

'Target File Extension (must include wildcard "*")
  myExtension = "*.xls*"

'Target Path with Ending Extention
  myfile = Dir(myPath & myExtension)

'Loop through each Excel file in folder
  Do While myfile <> ""
    'Set variable equal to opened workbook
      Set wb = Workbooks.Open(Filename:=myPath & myfile)
      Set y = Workbooks.Open("C:\Users\Admin\Desktop\Book3.xlsx")
      
      
    'Ensure Workbook has opened before moving on to next line of code
      DoEvents
    
    'Copy from file wb to y !!!!!!
      wb.Worksheets("Sheet1").Range("B2").Copy
      y.Worksheets("Sheet1").Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
      
      'Save and Close Workbook
      wb.Close savechanges:=True
      y.Close savechanges:=True
      
      'Ensure Workbook has closed before moving on to next line of code
      DoEvents
      
      'Get next file name
      myfile = Dir
  Loop
      
    'Ensure Workbook has closed before moving on to next line of code
      DoEvents


'Message Box when tasks are completed
  MsgBox "Task Complete!"

ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub
