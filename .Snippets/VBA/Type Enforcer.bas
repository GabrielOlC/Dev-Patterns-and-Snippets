' -------------------------------------------------------------------------
' Source: ROK app - IDs relationship
' Module: vsForceTextFormat
' Purpose: Forces selected range values to String type to prevent Excel 
'          from auto-converting number IDs into Integers :)
' Context: Critical for Game ID preservation for look ups during manual input
' -------------------------------------------------------------------------
Sub vsForceTextFormat()
    On Error Goto ErrHandler

    If TypeName(Selection) <> "Range" Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim vCellValue As String
    Dim cell As Range    
        For Each cell In Selection
            vCellValue = CStr(cell.Value)
            cell.Value = vCellValue
        Next cell

Clean:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    MsgBox "Error formatting text: " & Err.Description, vbCritical
    Resume Clean:

End Sub