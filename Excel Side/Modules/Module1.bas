Sub debug_print_array_formulas_all_sheets()
    Dim ws As Worksheet
    Dim rRange As Range, cell As Range
    Dim tot As Integer

    For Each ws In ThisWorkbook.Sheets
        On Error GoTo NotFound
        ws.Unprotect
        Dim v As Variant
        v = ws.UsedRange.HasFormula
        If IsNull(v) Or v Then

            Set rRange = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
            tot = 0
            For Each cell In rRange
                If cell.HasArray Then
                    MsgBox ws.name & ":" & cell.Address & " " & cell.Formula
                    tot = tot + 1
                End If
             Next cell
             If tot > 0 Then
                MsgBox "total number of array formula on " & ws.name & ": " & tot
                tot = 0
             End If
NotFound:
        End If
    Next ws
End Sub
