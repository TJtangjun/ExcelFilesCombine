Rem necessary reference ： Microsoft scripting runtime
Sub GetEveryFiles()
    Dim FileManager As New FileSystemObject
    Dim ThisFolder As Folder
    Dim ThisFile As File
    Dim FolderPicker As FileDialog
    Set FolderPicker = Application.FileDialog(msoFileDialogFolderPicker)
    FolderPicker.Show
    If FolderPicker.SelectedItems.Count = 0 Then Exit Sub
    Set ThisFolder = FileManager.GetFolder(FolderPicker.SelectedItems(1))
    Rem create a empty workbook file
             Dim NewWorkbook As Workbook
            Set NewWorkbook = Workbooks.Add
            Dim NewWorksheet As Worksheet
            Set NewWorksheet = NewWorkbook.Worksheets(1)
    Rem loop in files
    For Each ThisFile In ThisFolder.Files
        If Right(ThisFile.Name, 4) = "xlsx" And (Not ThisFile.Name Like "*_total.xlsx") Then
            Rem open file
            Dim CurrentBook As Workbook
            Set CurrentBook = Workbooks.Open(ThisFile.Path)
            Dim CurrentSheet As Worksheet
            Set CurrentSheet = CurrentBook.Worksheets(1)
           Rem copy and paste
            CurrentSheet.UsedRange.Copy
            Dim TargetPostion As Range
            Set TargetPostion = NewWorksheet.Range("A1048576").End(xlUp).Offset(1)
            If TargetPostion.Row = 2 Then Set TargetPostion = NewWorksheet.Range("A1")
            TargetPostion.PasteSpecial xlPasteValuesAndNumberFormats
            Application.CutCopyMode = 0
            Rem close and save
            CurrentBook.Close False
        End If
    Next
            Rem file name format: date+random number + folder name+total
            NewWorkbook.SaveAs ThisFolder & "\" & Format(Date, "yyyyMMdd_") & WorksheetFunction.RandBetween(1, 100) & ThisFolder.Name & "_total.xlsx"
            NewWorkbook.Close
    MsgBox "done."
End Sub
