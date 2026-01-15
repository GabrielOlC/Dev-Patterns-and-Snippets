## VBA snippets

***

### The Table Resizer (Strategic for ETL)

```visual-basic
' -------------------------------------------------------------------------
' Source: Relatório Financeiro Shopee/Relatório Financeiro Shopee (RFS)
' Purpose: Dynamically resizes an Excel ListObject to a single row.
'          Crucial for clearing data tables while preserving formulas/formats 
'          before importing new datasets via SQL/PowerQuery.
' -------------------------------------------------------------------------
Function vfResizeListObject(bvSheetCodeName As Worksheet, bvTableName As String) As Boolean
   vfResizeListObject = True
   On Error GoTo Errhandler

    ' Dynamically resize the specified table to 1 row (Header + 1 Data Row)
    bvSheetCodeName.ListObjects(bvTableName).Resize _
        Intersect(bvSheetCodeName.ListObjects(bvTableName).Range, _
        bvSheetCodeName.Rows(bvSheetCodeName.ListObjects(bvTableName).Range.Row & ":" & _
        bvSheetCodeName.ListObjects(bvTableName).Range.Row + 1))

Exit Function
Errhandler:
    Debug.Print "Error resizing table(vfResizeListObject): " & Err.Description
    vfResizeListObject = False
End Function
```

### Hex to RGB Converter

```visual-basic
' -------------------------------------------------------------------------
' Source: Calculadora de precificação Sacos para Lixo (CPSL)
' Function: vfHexToRGB
' Purpose: Converts Hexadecimal color codes (Web standard) into Long (Excel RGB).
'          Enables the use of modern UI color palettes within VBA forms.
' Context: Developed to allow dynamic UI theming based on product categories.
' -------------------------------------------------------------------------
Function vfHexToRGB(ByVal bvHexColor As String) As Long
    On Error GoTo ErrHandler

    Dim r As Long, g As Long, b As Long

    ' Clean input
    If Left(bvHexColor, 1) = "#" Then bvHexColor = Mid(bvHexColor, 2)

    ' Parse Hex channels to Decimal
    r = Val("&H" & Mid(bvHexColor, 1, 2))
    g = Val("&H" & Mid(bvHexColor, 3, 2))
    b = Val("&H" & Mid(bvHexColor, 5, 2))

    vfHexToRGB = RGB(r, g, b)

Exit Function
ErrHandler:
    Debug.Print "Error in vfHexToRGB: " & Err.Description
    vfHexToRGB = 0 ' Return black on error
End Function
```

### Unique Filter & Clipboard Export

```visual-basic
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
    On Error GoTo ErrHandler

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

    On Error GoTo ErrHandler

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
```

### Clipboard Utility

```visual-basic
' -------------------------------------------------------------------------
' Source: ROK app - Excel report V8+
' Function: vfCopyToClipboard
' Purpose: Low-level utility to push string data into the MSForms DataObject.
' Context: Shared utility used across ROK App, ERPM, and automated reporting tools.
' Requirements: library MSForms.DataObject
' -------------------------------------------------------------------------
Function vfCopyToClipboard(ByVal bvValues As String) As Boolean
    On Error GoTo ErrHandler

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
```

### Type Enforcer (Data Cleaning)

```visual-basic
' -------------------------------------------------------------------------
' Source: ROK app - IDs relationship
' Module: vsForceTextFormat
' Purpose: Forces selected range values to String type to prevent Excel 
'          from auto-converting number IDs into Integers :)
' Context: Critical for Game ID preservation for look ups during manual input
' -------------------------------------------------------------------------
Sub vsForceTextFormat()
    On Error GoTo ErrHandler

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
```
