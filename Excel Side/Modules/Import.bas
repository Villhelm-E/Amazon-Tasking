Option Explicit

Public Sub RequestFile()

    ' Get customer workbook...
    Dim filter As String
    Dim caption As String
    Dim OldEquipmentListFilename As String
    Dim OldEquipmentList As Workbook
    
    ' get the customer workbook
    filter = "Excel files (*.xlsx;*.xlsm; etc.),*.xls*"
    caption = "Please Select a Master Equipment List to Import..."
    
    'This asks for a document
    OldEquipmentListFilename = Application.GetOpenFilename(filter, , caption)
    
    'Only transfer data if user chose a document.
    If OldEquipmentListFilename <> "" Then
        Set OldEquipmentList = Application.Workbooks.Open(OldEquipmentListFilename)
        
        Call UpdateEquipmentList(OldEquipmentList)
        
        ' Close customer workbook
        OldEquipmentList.Close
        
        ' Inform the user that the import completed.
        MsgBox "Master Equipment List has been updated. Please delete the old one. You can then rename this one if you like.", vbOKOnly, "Update Status"
    End If
    

End Sub

' This function manually calls an import function for each worksheet in the workbook
Private Sub UpdateEquipmentList(SourceWorkbook As Workbook)

    Application.ScreenUpdating = False

    ' Update the Address Sheet
    Call ImportAddress(SourceWorkbook, SourceWorkbook.Sheets("Addresses"), NewWorkbook.Sheets("Addresses"))
    
    ' Update the Filters and Belts Sheet
'    Call ImportFiltersBelts(SourceWorkbook, SourceWorkbook.Sheets("Filters and Belts"), NewWorkbook.Sheets("Filters and Belts"))
    
    ' Update the Unit Types Sheet
    Call ImportUnitTypes(SourceWorkbook, SourceWorkbook.Sheets("Unit Types"), NewWorkbook.Sheets("Unit Types"))
    
    ' Update the Equipment List Sheet
    Call ImportEquipmentList(SourceWorkbook, SourceWorkbook.Sheets("Equipment List"), NewWorkbook.Sheets("Equipment List"))
    
    Application.ScreenUpdating = True

End Sub

' Function copies the column with the ColumnHeader to the TargetWorksheet
' The default pastes as values, KeepFormatting boolean will paste formatting
Private Sub MigrateColumn(ColumnHeader As String, SourceSheet As Worksheet, TargetSheet As Worksheet, Optional HeaderRow As Integer, Optional ApplyFormula As Boolean)

    ' Define variables
    Dim SourceColumn As Integer, TargetColumn As Integer
    Dim StartRow As Integer
    
    SourceSheet.Activate
    SourceColumn = ColumnLookup(SourceSheet, ColumnHeader, HeaderRow)
    
    TargetSheet.Activate
    TargetColumn = ColumnLookup(TargetSheet, ColumnHeader, HeaderRow)
    
    ' Start copying from the second row, ignoring the header
    StartRow = 2
    
    ' Find last row with data
    Dim LastRow As Integer
    SourceSheet.Activate
    LastRow = Cells(SourceSheet.Rows.Count, NumberToColumn(SourceColumn)).End(xlUp).Row
    
    ' If header doesn't exist in TargetSheet, add to the end of the headers
    If TargetColumn = 0 Then
        TargetColumn = CountColumns("1:1") + 1
        StartRow = 1
    End If
    
    ' Copy range from Old Workbook if column is not empty
    If LastRow > 1 Then
        SourceSheet.Activate
        SourceSheet.Range(Cells(StartRow, SourceColumn).Address, Cells(LastRow, SourceColumn).Address).Copy
        
        ' Paste range in New Workbook
        TargetSheet.Activate
        
        ' If ApplyFormula is true, then add a formula to the range instead of paste
        If ApplyFormula Then
            TargetSheet.Range(Cells(StartRow, SourceColumn).Address, Cells(LastRow, SourceColumn).Address).Formula = "=INDEX(Addresses!$B:$B,MATCH($A2,Addresses!$A:$A,0))"
        Else
            TargetSheet.Range(Cells(StartRow, TargetColumn).Address).PasteSpecial xlPasteValues
        End If
    End If
    
    Application.CutCopyMode = False

End Sub

' This function copies the equipment list from source worksheet and pastes in target worksheet
Private Sub ImportEquipmentList(SourceWorkbook As Workbook, SourceSheet As Worksheet, TargetSheet As Worksheet)
    
    ' Define variables
    Dim Columns As Integer
    SourceSheet.Activate
    Columns = CountColumns("1:1")
    
    ' Loop through columns
    Dim i As Integer
    For i = 1 To Columns
        ' If column header is "Site Type", enter formula instead of copy/paste
        If Cells(1, i).Value = "Site Type" Then
            SourceSheet.Activate
            Call MigrateColumn(Cells(1, i).Value, SourceSheet, TargetSheet, , True)
        Else
            ' Copy and paste from sourcesheet to targetsheet
            SourceSheet.Activate
            Call MigrateColumn(Cells(1, i).Value, SourceSheet, TargetSheet)
        End If
    Next i
    
    ' Final cleanup
    Application.Range("A1").Select

End Sub

' This function copies the addresses from source worksheet and pastes in target worksheet
Private Sub ImportAddress(SourceWorkbook As Workbook, SourceSheet As Worksheet, TargetSheet As Worksheet)

    ' Define variables
    Dim Columns As Integer
    SourceSheet.Activate
    Columns = CountColumns("1:1")
    
    ' Loop through columns
    Dim i As Integer
    For i = 1 To Columns
        SourceSheet.Activate
        Call MigrateColumn(Cells(1, i).Value, SourceSheet, TargetSheet)
    Next i
    
    Application.Range("A1").Select

End Sub

' This function copies the unit types from source worksheet and pastes in target worksheet
Private Sub ImportUnitTypes(SourceWorkbook As Workbook, SourceSheet As Worksheet, TargetSheet As Worksheet)

    ' Define variables
    Dim LastColumn
    LastColumn = SourceSheet.Cells(1, SourceSheet.Columns.Count).End(xlToLeft).column
    
    ' Age Threshold is the only value that needs to be migrated for now
    SourceSheet.Activate
    Call MigrateColumn(Cells(1, LastColumn).Value, SourceSheet, TargetSheet)
    
    Application.Range("A1").Select

End Sub

Private Sub ImportFiltersBelts(SourceWorkbook As Workbook, SourceSheet As Worksheet, TargetSheet As Worksheet)

    

End Sub
