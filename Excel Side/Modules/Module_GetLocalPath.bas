Option Explicit

Public Sub UpdateMasterEquipmentList()

    Dim OldEquipmentList As Workbook
    Dim Response
    Dim source As String
    Set SourceWorkbook = RequestWorkbook("Master Equipment List")
    source = SourceWorkbook.name
    
    'Only transfer data if user chose a document.
    If Not SourceWorkbook Is Nothing Then
    
        ' Check if chosen file is a master equipment list
        If WorksheetExists("Filters and Belts", SourceWorkbook) Then
            Dim SourceWorksheet As Worksheet
            Set SourceWorksheet = SourceWorkbook.Sheets("Equipment List")
            
            ' Check if equipment list is blank
            If FindLast(SourceWorksheet, "Row") > 1 Then
                ' Start the import
                Call UpdateEquipmentList(SourceWorkbook)
    
                ' Close customer workbook
                ' SourceWorkbook.Close
        
                ' Inform the user that the import completed.
                Response = MsgBox("Master Equipment List has been imported. Do you want to delete the old version?", vbYesNo, "Update Status")
                
                If Response = vbYes Then
                    ' delete old file
                    ' rename this one
                    Call RenameSelf(source)
                    MsgBox "Equipment list has been updated.", vbOKOnly, "Update"
                Else
                    MsgBox "The old master equipment list, '" & source & "', was not deleted." & vbNewLine & vbNewLine & _
                    "The new equipment list is now '" & ThisWorkbook.name & "'.", vbOKOnly, "Old Equipment List was not deleted"
                End If
            Else
                ' If equipment list is blank, let the user know and stop
                MsgBox "The chosen Master Equipment List is empty. Nothing to import."
            End If
        Else
            MsgBox "Not an Equipment List. Please choose another file.", vbOKOnly, "Invalid file"
            SourceWorkbook.Close
            End
        End If

        ' Close customer workbook
        SourceWorkbook.Close
    Else
        MsgBox "No file chosen", vbOKOnly, "Update Canceled"
    End If

End Sub

Public Sub AddSite()

    ' Open the equipment list template file
    Dim wb As Workbook
    Set SourceWorkbook = RequestWorkbook("Equipment List Template")
    
    Dim SourceSheet As Worksheet
    Dim TargetSheet As Worksheet
    Dim LastRow As Integer
    
    ' Check if chosen file is an equipment list template
    If SourceWorkbook.Sheets(1).name = "Equipment List" And SourceWorkbook.Sheets(2).name = "Unit Types" And SourceWorkbook.Sheets("Equipment List").Cells(1, 1).Value = "Site: " And _
    SourceWorkbook.Sheets("Equipment List").Cells(1, 3).Value = "Tech: " And SourceWorkbook.Sheets("Equipment List").Cells(1, 7).Value = "Date: " Then
        
        ' Grab the site name from the equipment list
        Dim SiteName As String
        SiteName = SourceWorkbook.Worksheets("Equipment List").Cells(1, 2).Value
        
        ' Check if the site is already in the equipment list
        Dim ColumnCount As Integer, RowCount As Integer
        
        Set SourceSheet = SourceWorkbook.Sheets("Equipment List")
        Set TargetSheet = ThisWorkbook.Sheets("Equipment List")
            
        'Check if template is blank
        If FindLast(SourceSheet, "Row") > 2 Then
            
            LastRow = FindLastRow(TargetSheet, "A", "*")
            
            If WorksheetFunction.CountIf(Range("A:A"), SiteName) > 0 Then
                'Ask user if they are importing a split equipment list
                Dim user_val
                user_val = MsgBox(SiteName & " is already in the equipment list. Are you importing a split equipment list?", vbYesNo, SiteName & " already exists")
                
                If user_val = vbNo Then
                    MsgBox SiteName & " is already in the equipment list."
                Else
                    ' Disable screen updating
                    Application.ScreenUpdating = False
                    
                    Set SourceSheet = SourceWorkbook.Sheets("Equipment List")
                    Set TargetSheet = ThisWorkbook.Sheets("Equipment List")
                    
                    ' Copy every column from SourceSheet into TargetSheet
                    Call MigrateEquipmentList(SourceSheet, TargetSheet)
                    
                    ' Add the site name, address, and site type to Addresses worksheet
                    Dim UnitCount As Integer, StartRow As Integer, SiteRow As Integer
                    Dim SiteAddress As String, SiteType As String
                    
                    'Count number of units in template
                    SourceSheet.Activate
                    UnitCount = FindLast(SourceSheet, "Row") - 2
                    
                    ' Activate Addresses sheet to find site address info
                    Set TargetSheet = ThisWorkbook.Sheets("Addresses")
                    TargetSheet.Activate
                    
                    ' Lookup site and address info
                    SiteRow = FindLastRow(TargetSheet, "A", SiteName)
                    SiteType = Cells(SiteRow, 2)
                    SiteAddress = Cells(SiteRow, 3)
                    
                    ' Calculate range for migrating
                    StartRow = LastRow + 1
                    LastRow = StartRow + UnitCount - 1
                    
                    Set TargetSheet = ThisWorkbook.Sheets("Equipment List")
                    TargetSheet.Activate
                    
                    ' Migrate the address info
                    Call MigrateAddress(StartRow, LastRow, SiteName, SiteAddress, SiteType)
                    
                    ' Final cleanup
                    SourceWorkbook.Close
                    
                    ' Go to the appropriate row
                    Set TargetSheet = ThisWorkbook.Worksheets("Equipment List")
                    TargetSheet.Activate
                    
                    Application.GoTo Reference:=Range("A" & LastRow - 14), Scroll:=True
                    Application.GoTo Reference:=Range("A" & LastRow + 1), Scroll:=False
                    
                    ' Enable screen updating
                    Application.ScreenUpdating = True
                End If
            Else
                ' Show the import menu
                NewSite.Show
            End If
        Else
            ' If the equipment list is blank, let the user know and stop.
            MsgBox "The chosen equipment list is blank. Nothing to import."
        End If
    Else
        ' Cancel the import if user didn't choose a valid file
        SourceWorkbook.Close
        MsgBox "This is not an Equipment List Template. Please choose a different file.", vbOKOnly, "Invalid file"
        End
    End If

End Sub

Public Sub MigrateEquipmentList(SourceSheet As Worksheet, TargetSheet As Worksheet)

    ' Variables
    Dim ColumnCount As Integer, RowCount As Integer, LastRow As Integer
    
    ' Count number of columns in source sheet
    SourceSheet.Activate
    ColumnCount = CountColumns("2:2")
    
    ' Count number of rows in SourceSheet
    TargetSheet.Activate
    RowCount = FindLast(SourceSheet, "Row") - 2
    
    ' Find last row with a Site Name
    LastRow = FindLast(TargetSheet, "Row")
    
    ' Special case if equipment list is empty
    If LastRow > 1 Then
        ' Insert rows so that cells have dropdowns
        Range(LastRow + 1 & ":" & LastRow + RowCount).EntireRow.Insert
        ' But skip a row so that if the heade row is the only row,
        ' Have to delete the skipped row so that the pasted rows are first
        Range(LastRow + 1 & ":" & LastRow + 1).EntireRow.Delete
    Else
        ' If equipment list is empty, have to start at row 3 to avoid the header formatting applying to all inserted rows
        Range("3:" & LastRow + RowCount).EntireRow.Insert
        Range("2:2").EntireRow.Delete
    End If
    
    ' Loop through columns
    Dim i As Integer
    For i = 1 To ColumnCount
        ' Copy and paste from SourceSheet to TargetSheet
        SourceSheet.Activate
        Call MigrateColumn(Cells(2, i).Value, SourceSheet, TargetSheet, 2, 3, LastRow + 1)
    Next i

End Sub

Public Sub MigrateAddress(StartRow As Integer, LastRow As Integer, SiteName As String, SiteAddress As String, SiteType As String)

    Range(Cells(StartRow, 1), Cells(LastRow, 1)).Value = SiteName
    Range(Cells(StartRow, 2), Cells(LastRow, 2)).Value = SiteAddress
    Range(Cells(StartRow, 3), Cells(LastRow, 3)).formula = SiteType

End Sub

' This function manually calls an import function for each worksheet in the workbook
Private Sub UpdateEquipmentList(SourceWorkbook As Workbook)

    Application.ScreenUpdating = False

    ' Update the Address Sheet
    Call ImportAddress(SourceWorkbook, SourceWorkbook.Sheets("Addresses"), ThisWorkbook.Sheets("Addresses"))
    
    ' Update the Filters and Belts Sheet
'    Call ImportFiltersBelts(SourceWorkbook, SourceWorkbook.Sheets("Filters and Belts"), ThisWorkbook.Sheets("Filters and Belts"))
    
    ' Update the Unit Types Sheet
    Call ImportUnitTypes(SourceWorkbook, SourceWorkbook.Sheets("Unit Types"), ThisWorkbook.Sheets("Unit Types"))
    
    ' Update the Equipment List Sheet
    Call ImportEquipmentList(SourceWorkbook, SourceWorkbook.Sheets("Equipment List"), ThisWorkbook.Sheets("Equipment List"))
    
    Application.ScreenUpdating = True

End Sub

' Function copies the column with the ColumnHeader to the TargetWorksheet
' The default pastes as values, KeepFormatting boolean will paste formatting
Public Function MigrateColumn(ColumnHeader As String, SourceSheet As Worksheet, TargetSheet As Worksheet, Optional HeaderRow As Integer, Optional CopyStartRow As Integer, Optional PasteStartRow As Integer, Optional ApplyFormula As Boolean)

    ' Define variables
    Dim SourceColumn As Integer, TargetColumn As Integer
    
    SourceSheet.Activate
    SourceColumn = ColumnLookup(SourceSheet, ColumnHeader, HeaderRow)
    
    TargetSheet.Activate
    TargetColumn = ColumnLookup(TargetSheet, ColumnHeader, 1) 'changed 1 to 2
    
    ' Find the last row with data
    Dim tbl As ListObject
    Dim LastRow As Long
    
    SourceSheet.Activate
    LastRow = FindLast(SourceSheet, "Row")
    
    ' If header doesn't exist in TargetSheet, add to the end of the headers
    If TargetColumn = 0 Then
        TargetColumn = CountColumns("1:1") + 1
        CopyStartRow = 1
    End If
    
    ' If CopyStartRow is blank, assume row below header
    If CopyStartRow = 0 Then
        CopyStartRow = HeaderRow + 1
    End If
    
    ' If PasteStartRow is blank, assume row below header
    If PasteStartRow = 0 Then
        PasteStartRow = HeaderRow + 1
    End If
    
    ' Copy range from SourceSheet if column is not empty
    If LastRow > 1 Then
        SourceSheet.Activate
        SourceSheet.Range(Cells(CopyStartRow, SourceColumn).Address, Cells(LastRow, SourceColumn).Address).Copy
        
        ' Paste range in New Workbook
        TargetSheet.Activate
        
        ' If ApplyFormula is true, then add a formula to the range instead of paste
        If ApplyFormula Then
            TargetSheet.Range(Cells(PasteStartRow, SourceColumn).Address, Cells(LastRow, SourceColumn).Address).formula = "=INDEX(Addresses!$B:$B,MATCH($A2,Addresses!$A:$A,0))"
        Else
            TargetSheet.Range(Cells(PasteStartRow, TargetColumn).Address).PasteSpecial xlPasteValues
        End If
    End If
    
    Application.CutCopyMode = False

End Function

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
            Call MigrateColumn(Cells(1, i).Value, SourceSheet, TargetSheet, , , , True)
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

    ' Probably not necessary

End Sub
