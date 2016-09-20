Attribute VB_Name = "Module3"
Sub Button8_Click()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    
    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    diaFolder.AllowMultiSelect = False
    diaFolder.Show
    'MsgBox diaFolder.SelectedItems(1)
    
    Path = diaFolder.SelectedItems(1) & "\Route Adresses.txt"
    
    Set oFile = fso.CreateTextFile(Path)
    
    readSheetNum = InputBox("Please Input the Sheet Number of the Daily Task List")
    num = Int(readSheetNum)
    
    Worksheets(num).Select
    
    colNum = 1
    rowNum = 5
    
    'colNum will be column number of address
    While Trim(Cells(rowNum, colNum).Value) <> "Address" And colNum < 23
        colNum = colNum + 1
    Wend
    
    rowNum = 6
    
    While (Cells(rowNum, colNum).Value <> "")
        If (Cells(rowNum, colNum).Value <> "MISSING DATA") Then
            oFile.WriteLine Cells(rowNum, colNum).Value
        End If
        rowNum = rowNum + 1
    Wend
    
    oFile.Close
    Set fso = Nothing
    Set oFile = Nothing
End Sub

'will rely on the MISSING DATA TAG
