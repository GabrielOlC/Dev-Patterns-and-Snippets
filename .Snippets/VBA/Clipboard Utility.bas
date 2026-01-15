' -------------------------------------------------------------------------
' Source: ROK app - Excel report V8+
' Function: vfCopyToClipboard
' Purpose: Low-level utility to push string data into the MSForms DataObject.
' Context: Shared utility used across ROK App, ERPM, and automated reporting tools.
' Requirements: library MSForms.DataObject
' -------------------------------------------------------------------------
Function vfCopyToClipboard(ByVal bvValues As String) As Boolean
    On Error Goto ErrHandler

    Dim vClipboardData As New MSForms.DataObject
    With vClipboardData
        .SetText bvValues
        .PutInClipboard
    End With

    vfCopyToClipboard = True
    Exit Function

ErrHandler:
    Debug.Print "Clipboard Error (vfCopyToClipboard): " & Err.Description
    vfCopyToClipboard = False
End Function