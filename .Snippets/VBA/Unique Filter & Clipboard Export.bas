' -------------------------------------------------------------------------
' Source: ROK app - Excel report V11+
' Module: vsExportUniqueIDs
' Purpose: Extracts unique Key/Value pairs from a filtered ListObject and 
'          exports them directly to the system clipboard.
' Context: Used to rapidly extract IDs for cross-referencing against 
'          external sources
' Requirements: vfCopyToClipboard
' -------------------------------------------------------------------------
Sub vsExportUniqueIDs(bvSheet As Worksheet, bvTableName As String, bvColKey As String, bvColValToDict As String)
    On Error Goto ErrHandler

    Dim sTb As ListObject
    Dim sDict As Object
    Dim vRngName As Range, vRngID As Range
    Dim vData As String


    ' Initialize Dictionary
    Set sDict = CreateObject("Scripting.Dictionary")
    Set sTb = bvSheet.ListObjects(bvTableName)

    ' Isolate visible data (respecting user filters)
    On Error Resume Next ' Handle empty filter result

    Set vRngID = sTb.ListColumns("bvColKey").DataBodyRange.SpecialCells(xlCellTypeVisible) ' xlCellTypeVisible = not filtered out
    Set vRngName = sTb.ListColumns("bvColValToDict").DataBodyRange.SpecialCells(xlCellTypeVisible)

    On Error Goto ErrHandler

    If vRngID Is Nothing Then Exit Sub

    ' Extraction Logic
    Dim vCell As Range
    For Each vCell In vRngID
        If Not sDict.Exists(vCell.Value) And vCell.Value <> "" Then 'if key already not added and not empty'
            sDict.Add vCell.Value, vRngName.Cells(vCell.Row - vRngName.Cells(1).Row + 1).Value
        End If
    Next vCell

    ' Format Data
    Dim vKey As Variant
    For Each vKey In sDict.Keys
        vData = vData & sDict(vKey) & vbTab & vKey & vbCrLf
    Next vKey

    ' Send to Clipboard
    If Not vfCopyToClipboard(vData) Then Err.Raise 999, , "Clipboard Failure"

    MsgBox "Unique IDs extracted to clipboard.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error extracting IDs: " & Err.Description, vbCritical
End Sub