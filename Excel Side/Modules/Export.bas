Option Explicit

Public Sub JSON()

    Dim savename As String
    Dim MyFile As String
    Dim lcolumn As Long
    Dim lrow As Long
    
    savename = "equipment_list.json"
    'MyFile = Application.DefaultFilePath & "\" & savename
    MyFile = GetLocalPath(Application.ActiveWorkbook.Path) & "/" & savename
    
    Open MyFile For Output As #1
    Dim wkb As Workbook
    Dim wks As Worksheet
    
    Set wkb = ThisWorkbook
    Set wks = wkb.Sheets(1)
    
    'Find last column
    lcolumn = wks.Cells(1, Columns.Count).End(xlToLeft).column
    
    'Find last row
    lrow = wks.Range("A1").SpecialCells(xlCellTypeLastCell).Row
    
    Dim headers() As String
    ReDim headers(lcolumn)
    
    Dim i As Integer
    
    ' Remove spaced from headers to use in JSON key:value pairs
    For i = 1 To lcolumn
        headers(i) = Replace(wks.Cells(1, i), " ", "")
    Next i
    
    ' Used for double quotes
    Dim dq As String
    dq = """"
    
    Dim j As Integer
    Dim unit_array As String
    Dim cell_value As String
    
    ' Category column is currently a constant
    ' Todo: programmatically find the Category column
    Dim CATEGORY_COLUMN As Integer
    CATEGORY_COLUMN = 6 'Category is the 6th column
    
    ' Loop through equipment list
    For j = 2 To lrow
        ' Loop through columns
        For i = 1 To lcolumn
            cell_value = wks.Cells(j, i)
            
            If i <> lcolumn Then
                ' Add each header and its contents on the row with a comma
                unit_array = unit_array & dq & headers(i) & dq & ":" & dq & cell_value & dq & ", "
            Else
                ' If it's the last column, don't add the comma
                unit_array = unit_array & dq & headers(i) & dq & ":" & dq & cell_value & dq
            End If
            
            ' Clear the variable for the next row
            cell_value = ""
        Next i

        ' Add unit string to JSON file
        Print #1, "{" & unit_array & "},"
        
        'reset variable for next row
        unit_array = ""
        
    Next j
    
    ' Close the file
    Close #1
    
    ' The following code sets up the JSON file to be edited like a string
    ' I copy-pasted the code so I'm not quite sure how it works
    Dim sBuf As String
    Dim sTemp As String
    Dim iFileNum As Integer
    
    iFileNum = FreeFile
    Open MyFile For Input As iFileNum
    
    Do Until EOF(iFileNum)
        Line Input #iFileNum, sBuf
        sTemp = sTemp & sBuf & vbCrLf
    Loop
    Close iFileNum
    
    ' Delete the final comma and whitespace and wrap the JSON file with brackets
    sTemp = "[" & Left(Trim(sTemp), Len(Trim(sTemp)) - 3) & "]"
    
    iFileNum = FreeFile
    Open MyFile For Output As iFileNum
    Print #iFileNum, sTemp 'I think this is the code that replaces the contents of the JSON file with the modified text above
    
    Close iFileNum
    
    MsgBox "JSON file was exported successfully"

End Sub

Public Sub CSV()

    Dim FlSv As Variant
    Dim MyFile As String
    Dim sh As Worksheet
    Dim currentsheetname As Variant
    
    'Save current sheet's name to variable
    currentsheetname = Application.ActiveSheet.name
    
    'activate current sheet and copy it
    Set sh = Sheets(currentsheetname)
    sh.Copy
    
    'add extension to sheet name
    MyFile = currentsheetname & ".csv"
    
    'open up the save as window to verify file name and choose a destination to save
    FlSv = Application.GetSaveAsFilename(MyFile, "Excel Files (*.csv), *.csv)", , "Enter your file name")
    
    'if user cancels the operation, close the copy and stop code
    If FlSv = False Then
        ActiveWorkbook.Close False
        Exit Sub
    End If

End Sub
