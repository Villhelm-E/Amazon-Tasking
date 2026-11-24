Public Sub CallbackExportJSON(control As IRibbonControl)

    JSON

End Sub

Public Sub CallbackImportSite(control As IRibbonControl)

    AddSite

End Sub

Public Sub CallbackImportEquipmentList(control As IRibbonControl)

    ' Only run the sub if the equipment list is empty
    If Application.CountA(Range("A2").EntireRow) = 0 Then
        UpdateMasterEquipmentList
    Else
        MsgBox "Equipment List has already been updated.", vbOKOnly, "Already updated"
    End If

End Sub

Public Sub CallbackVersion(control As IRibbonControl)

    MsgBox "The current version is " & GetVersion, vbOKOnly, "Current Version"

End Sub
