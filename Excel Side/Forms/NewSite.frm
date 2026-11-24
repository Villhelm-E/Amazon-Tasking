VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewSite 
   Caption         =   "Add a new site"
   ClientHeight    =   1875
   ClientLeft      =   120
   ClientTop       =   460
   ClientWidth     =   10660
   OleObjectBlob   =   "Untitled.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SourceWB As String ' Private internal variable
Private pWB As Workbook

Public Sub SetObjects(ByVal wb As Workbook, Optional ByVal lo As ListObject)
    Set pWB = wb
End Sub

' Property Let writes the variable to the form
Public Property Let Property_SiteName(ByVal NewValue As String)

    SourceWB = NewValue

    ' You can update your form UI directly inside the property assignment
    Me.SiteName.Value = SourceWB

End Property

Public Property Get Property_SiteName() As String

    Property_SiteName = SourceWB

End Property

Private Sub SiteAddressLabel_Click()

End Sub

Private Sub UserForm_Initialize()

    ' Closing the source workbook at the wrong time will stop the global array SiteTypes
    ' from getting its values
    ' PrepWorkbook restores the required variables
    If IsEmpty(SiteTypes) Then
        PrepWorkbook
    End If
    
    ' Populate the Site Type listbox
    Dim i As Integer
    For i = 0 To UBound(SiteTypes) - LBound(SiteTypes)
        SiteType.AddItem SiteTypes(i)
    Next
    
    ' Pull site name from equipment list template
    SourceWorkbook.Worksheets("Equipment List").Activate
    Me.SiteName.Value = Cells(1, 2).Value
    
End Sub

Private Sub bAddSite_Click()
    
    ' Define variables
    Dim SourceSheet As Worksheet
    Dim TargetSheet As Worksheet
    Dim ColumnCount As Integer, RowCount As Integer, LastRow As Long
    Dim place As Integer
    Dim UnitCount As Integer, StartRow As Integer, EndRow As Integer
    Dim SiteName As String, SiteAddress As String
    
    ' Check for required fields
    If Me.SiteName.Value <> "" And Me.SiteAddress.Value <> "" And Me.SiteType.Value <> "" Then
        ' Disable screen updating
        Application.ScreenUpdating = False
        
        Set SourceSheet = SourceWorkbook.Sheets("Equipment List")
        Set TargetSheet = ThisWorkbook.Sheets("Equipment List")
        
        ' Copy every column from SourceSheet into TargetSheet
        Call MigrateEquipmentList(SourceSheet, TargetSheet)
        
        ' Add the site name, address, and site type to the units
        SiteName = Me.SiteName
        SiteAddress = Me.SiteAddress
        
        ' Calculate range for migrating
        UnitCount = FindLast(SourceSheet, "Row") - 2
        StartRow = FindLastRow(TargetSheet, "A", "*") + 1
        EndRow = StartRow + UnitCount - 1
        
        ' Migrate the values
        TargetSheet.Activate
        Call MigrateAddress(StartRow, EndRow, SiteName, SiteAddress, "=INDEX(Addresses!$B:$B,MATCH($A" & StartRow & ",Addresses!$A:$A,0))")
        
        ' Add the site name, address, and site type to Addresses worksheet
        Dim AddressCount As Integer
        Set TargetSheet = ThisWorkbook.Sheets("Addresses")
        TargetSheet.Activate
        AddressCount = Cells(TargetSheet.Rows.Count, "A").End(xlUp).Row
        
        Range("A" & AddressCount + 1).Value = SiteName
        Range("B" & AddressCount + 1).Value = Me.SiteType.Value
        Range("C" & AddressCount + 1).Value = SiteAddress
        
        place = ThisWorkbook.Worksheets("Equipment List").Cells(ThisWorkbook.Worksheets("Equipment List").Rows.Count, "A").End(xlUp).Row + 1
        
        ' Final cleanup
        Unload NewSite
        
        ' Go to the appropriate row
        Set TargetSheet = ThisWorkbook.Worksheets("Equipment List")
        TargetSheet.Activate
        
        Application.GoTo Reference:=Range("A" & place - 15), Scroll:=True
        Application.GoTo Reference:=Range("A" & place), Scroll:=False
        
        ' Enable screen updating
        Application.ScreenUpdating = False
    Else
        MsgBox "Please enter all required fields.", vbOKOnly, "Missing Fields"
    End If

End Sub

Private Sub bCancel_Click()

    SourceWorkbook.Close
    ThisWorkbook.Activate
    Unload NewSite
    
End Sub

Private Sub UserForm_Terminate()

    SourceWorkbook.Close

End Sub
Attribute VB_Name = "NewSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SourceWB As String ' Private internal variable
Private pWB As Workbook

Public Sub SetObjects(ByVal wb As Workbook, Optional ByVal lo As ListObject)
    Set pWB = wb
End Sub

' Property Let writes the variable to the form
Public Property Let Property_SiteName(ByVal NewValue As String)

    SourceWB = NewValue

    ' You can update your form UI directly inside the property assignment
    Me.SiteName.Value = SourceWB

End Property

Public Property Get Property_SiteName() As String

    Property_SiteName = SourceWB

End Property

Private Sub SiteAddressLabel_Click()

End Sub

Private Sub UserForm_Initialize()

    ' Closing the source workbook at the wrong time will stop the global array SiteTypes
    ' from getting its values
    ' PrepWorkbook restores the required variables
    If IsEmpty(SiteTypes) Then
        PrepWorkbook
    End If
    
    ' Populate the Site Type listbox
    Dim i As Integer
    For i = 0 To UBound(SiteTypes) - LBound(SiteTypes)
        SiteType.AddItem SiteTypes(i)
    Next
    
    ' Pull site name from equipment list template
    SourceWorkbook.Worksheets("Equipment List").Activate
    Me.SiteName.Value = Cells(1, 2).Value
    
End Sub

Private Sub bAddSite_Click()
    
    ' Define variables
    Dim SourceSheet As Worksheet
    Dim TargetSheet As Worksheet
    Dim ColumnCount As Integer, RowCount As Integer, LastRow As Long
    Dim place As Integer
    Dim UnitCount As Integer, StartRow As Integer, EndRow As Integer
    Dim SiteName As String, SiteAddress As String
    
    ' Check for required fields
    If Me.SiteName.Value <> "" And Me.SiteAddress.Value <> "" And Me.SiteType.Value <> "" Then
        ' Disable screen updating
        Application.ScreenUpdating = False
        
        Set SourceSheet = SourceWorkbook.Sheets("Equipment List")
        Set TargetSheet = ThisWorkbook.Sheets("Equipment List")
        
        ' Copy every column from SourceSheet into TargetSheet
        Call MigrateEquipmentList(SourceSheet, TargetSheet)
        
        ' Add the site name, address, and site type to the units
        SiteName = Me.SiteName
        SiteAddress = Me.SiteAddress
        
        ' Calculate range for migrating
        UnitCount = FindLast(SourceSheet, "Row") - 2
        StartRow = FindLastRow(TargetSheet, "A", "*") + 1
        EndRow = StartRow + UnitCount - 1
        
        ' Migrate the values
        TargetSheet.Activate
        Call MigrateAddress(StartRow, EndRow, SiteName, SiteAddress, "=INDEX(Addresses!$B:$B,MATCH($A" & StartRow & ",Addresses!$A:$A,0))")
        
        ' Add the site name, address, and site type to Addresses worksheet
        Dim AddressCount As Integer
        Set TargetSheet = ThisWorkbook.Sheets("Addresses")
        TargetSheet.Activate
        AddressCount = Cells(TargetSheet.Rows.Count, "A").End(xlUp).Row
        
        Range("A" & AddressCount + 1).Value = SiteName
        Range("B" & AddressCount + 1).Value = Me.SiteType.Value
        Range("C" & AddressCount + 1).Value = SiteAddress
        
        place = ThisWorkbook.Worksheets("Equipment List").Cells(ThisWorkbook.Worksheets("Equipment List").Rows.Count, "A").End(xlUp).Row + 1
        
        ' Final cleanup
        Unload NewSite
        
        ' Go to the appropriate row
        Set TargetSheet = ThisWorkbook.Worksheets("Equipment List")
        TargetSheet.Activate
        
        Application.GoTo Reference:=Range("A" & place - 15), Scroll:=True
        Application.GoTo Reference:=Range("A" & place), Scroll:=False
        
        ' Enable screen updating
        Application.ScreenUpdating = False
    Else
        MsgBox "Please enter all required fields.", vbOKOnly, "Missing Fields"
    End If

End Sub

Private Sub bCancel_Click()

    SourceWorkbook.Close
    ThisWorkbook.Activate
    Unload NewSite
    
End Sub

Private Sub UserForm_Terminate()

    SourceWorkbook.Close

End Sub
