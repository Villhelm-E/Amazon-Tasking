Option Explicit

' Global Variables
Public NewWorkbook As Workbook

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
    ' LastColumn = NumberToColumn(CountColumns(Range(HeaderRow & ":" & HeaderRow)))
    LastColumn = NumberToColumn(TargetSheet.Cells(HeaderRow, TargetSheet.Columns.Count).End(xlToLeft).column)
    
    ' Define range to search
    Set HeaderRange = Range("A" & HeaderRow & ":" & LastColumn & HeaderRow)
    
    ' Search through range to find column of where field name is found
    Set FoundHeader = HeaderRange.find(SearchItem, , , xlWhole, , , False, , False)
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
