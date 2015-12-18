Set fso = CreateObject("Scripting.FileSystemObject")
startFolder = "."

Set startFolderObj = fso.GetFolder(startFolder)

Set folders = startFolderObj.SubFolders

For Each folder in folders
  'Wscript.Echo folder.Path

    ' get excel files in path
    For Each file in folder.Files
      'Wscript.Echo file.Path


      If fso.GetExtensionName(file.Path) = "xlsx" Then



        Wscript.Echo file.Path
        excelFile = file.Path

        Set objXL = CreateObject("Excel.Application")
        Set objWorkBook = objXL.Workbooks.Open(excelFile)
        objXL.DisplayAlerts = False

        rem loop over worksheets
        For Each sheet In objWorkBook.Sheets

          If objXL.Application.WorksheetFunction.CountA(sheet.Cells) <> 0 Then
          rem sheet.Rows(1).delete ' this will remove Row 1 or the header Row
            sheet.SaveAs startFolderObj.Path & "\" & folder.Name & "_" & sheet.Name & ".csv", 6 'CSV
          End If
        Next

        rem clean up
        objWorkBook.Close
        objXL.quit
        Set objXL = Nothing
        Set objWorkBook = Nothing




      End If
    Next
Next

  Set fso = Nothing
