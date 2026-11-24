Option Explicit

' Global Variables
Public NewWorkbook As Workbook
Public SourceWorkbook As Workbook
Public SiteTypes

Public Sub PrepWorkbook()

    ' This global stores the current workbook as the new workbook
    Set NewWorkbook = Application.ActiveWorkbook
    
    ' This variable stores the Site Types used in the NewSite Form dropdown
    ReDim SiteTypes(5)
    SiteTypes(0) = "AMZL"
    SiteTypes(1) = "ATS"
    SiteTypes(2) = "GSF"
    SiteTypes(3) = "Logistics"
    SiteTypes(4) = "NACF"
    SiteTypes(5) = "NASC"

End Sub

' Function returns the column number of the header the user inputs
' Public Function ColumnLookup(lastcolumnletter As String, SearchItem As String) As Integer
Public Function ColumnLookup(TargetSheet As Worksheet, SearchItem As String, Optional HeaderRow As Integer) As Integer

    Dim HeaderRange As Range
    Dim FoundHeader As Range
    
    ' Make 1 the default value for HeaderRow
    If HeaderRow = 0 Then
        HeaderRow = 1
    End If
    
    ' Find the column of the last header
    Dim LastColumn
    LastColumn = NumberToColumn(TargetSheet.Cells(HeaderRow, TargetSheet.Columns.Count).End(xlToLeft).column)
    
    ' Define range to search
    Set HeaderRange = Range("A" & HeaderRow & ":" & LastColumn & HeaderRow)
    
    ' Search through range to find column of where field name is found
    Set FoundHeader = HeaderRange.Find(SearchItem, , , xlWhole, , , False, , False)
    If Not FoundHeader Is Nothing Then
        ColumnLookup = FoundHeader.column
    End If

End Function

' Function takes column number as input and returns the column letter
Public Function NumberToColumn(column As Integer) As String

    Dim vArr
    vArr = Split(Cells(1, column).Address(True, False), "$")
    NumberToColumn = vArr(0)

End Function

' Function takes rows as input and returns the number of columns in range
Public Function CountColumns(RowRange) As Integer

    ' User can enter a range, an integer, or a string as a row range input
    Select Case VarType(RowRange)
    Case 8204 ' Range
        CountColumns = Application.CountA(RowRange)
        
    Case 2  ' Integer
        CountColumns = Application.CountA(Rows(RowRange))
        
    Case 8  ' String
        If InStr(1, RowRange, ":") > 0 Then
            CountColumns = Application.CountA(Range(RowRange))
        Else
            CountColumns = Application.CountA(Range(RowRange & ":" & RowRange))
        End If
        
    Case Else
        CountColumns = -1
        
    End Select

End Function

' Function returns true if worksheet name exists
Public Function FindSheet(name As String, Optional Workbook As Workbook) As Boolean
    
    ' Initialize
    Dim wsSheet As Worksheet
    FindSheet = False
    
    If Workbook Is Nothing Then
        Set Workbook = ActiveWorkbook
    End If
    
    ' Loop through every worksheet
    For Each wsSheet In Workbook.Worksheets
        If wsSheet.name = name Then
            ' Set function to true if sheet found
            FindSheet = True
            GoTo exit_loop
        End If
    Next
exit_loop:
End Function

' This function returns a workbook the user chooses from a file picker window
' This function returns null if user did not choose a file
Public Function RequestWorkbook(Optional FileType As String) As Workbook

    ' Setup variable
    Dim FileName As String
    
    ' handlecase where filetype is blank
    If FileType = "" Then
        FileType = "file"
    End If
    
    ' Get the customer workbook
    Do
        ' Show the file choice dialog
        FileName = Application.GetOpenFilename( _
            FileFilter:="Excel files (*.xlsx;*.xlsm;*.xltx),*.xlsx;*.xlsm;*.xltx", _
            Title:="Please Select a " & FileType & " to Import...", _
            MultiSelect:=False)
            
        ' User clicked Cancel
        If FileName = "" Then Exit Function
        
        ' Check if selected file is the active workbook
        If CStr(FileName) = GetLocalPath(ThisWorkbook.FullName) Then
            MsgBox "You cannot select the active workbook! Please choose another file.", vbExclamation, "Invalid Selection"
            FileName = "" ' Reset to loop again
        End If
    Loop While FileName = ""
    
    ' Return workbook user chose, null if canceled
    If FileName <> "False" Then
        Set SourceWorkbook = Application.Workbooks.Open(FileName)
        Set RequestWorkbook = SourceWorkbook
    Else
        End
    End If

End Function

' Source - https://stackoverflow.com/a/6688482
' Posted by Tim Williams, modified by community. See post 'Timeline' for change history
' Retrieved 2026-07-22, License - CC BY-SA 4.0

Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    WorksheetExists = (wb.Worksheets(shtName).name = shtName)
End Function

Public Function CompareVersions(SourceWorkbook As Workbook, TargetWorkbook As Workbook) As Boolean

    Dim TargetVersion, SourceVersion
    
    TargetVersion = Application.Run("'" & TargetWorkbook.name & "'!GetVersion")
    
    On Error GoTo old_equipment_list
    SourceVersion = Application.Run("'" & SourceWorkbook.name & "'!GetVersion")
    
    CompareVersions = SourceWorkbook = TargetWorkbook
    Exit Function
    
old_equipment_list:
    CompareVersions = False

End Function

Sub RenameSelf(NewName As String)
    Dim oldPath As String, newPath As String
    
    ' Ensure the workbook is saved before renaming
    If ActiveWorkbook.Path = "" Then
        MsgBox "Please save the workbook first."
        Exit Sub
    End If
    
    oldPath = ActiveWorkbook.FullName
    newPath = ActiveWorkbook.Path & "\" & NewName ' Change name here
    
    ' Save copy to the new name
    ActiveWorkbook.SaveAs FileName:=newPath, FileFormat:=ActiveWorkbook.FileFormat
    
    'Close Source workbook
    SourceWorkbook.Close
    
    ' Delete the original file
    Kill GetLocalPath(oldPath)
End Sub

Public Sub AddDropdown(Range As Range, formula As String)

    With Range(Range).Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:=formula
             
        .InCellDropdown = True
    End With

End Sub

' Searches a worksheet for a string and returns the row or column of the last instance
' TargetType returns Row if user enters 0, Column if user enters 1
' Returns 0 if no results are found
Public Function FindLast(SearchSheet As Worksheet, TargetType As String, Optional SearchTerm As String)

    ' If user doesn't provide a SearchTerm, default to "*"
    If SearchTerm = "" Then
        SearchTerm = "*"
    End If
    
    Dim LastCell As Range
    Set LastCell = SearchSheet.Cells.Find(SearchTerm, SearchOrder:=xlByRows, searchDirection:=xlPrevious)
    
    Select Case TargetType
        Case "Row"
            FindLast = LastCell.Row
        Case "Column"
            FindLast = LastCell.column
        Case Else
            FindLast = 0
    End Select

End Function

Public Function FindLastRow(SearchSheet As Worksheet, SearchColumn As String, SearchTerm As String) As Integer
    
    ' Make sure SearchSheet is active
    SearchSheet.Activate
    
    ' Find the last instance of SearchTerm in SearchColumn, then return the row
    FindLastRow = SearchSheet.Columns(SearchColumn).Find(What:=SearchTerm, _
                                                    LookIn:=xlValues, _
                                                    SearchOrder:=xlByRows, _
                                                    searchDirection:=xlPrevious).Row

End Function

Public Function CountTableRows(SourceSheet As Worksheet, TableName As String) As Integer

    Dim tbl As ListObject
    
    ' To-do: figure out how to avoid hard-coded table name
    Set tbl = SourceSheet.ListObjects(TableName)
    
    ' Return count
    CountTableRows = tbl.ListRows.Count

End Function

