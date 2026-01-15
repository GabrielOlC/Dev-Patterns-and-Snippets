' -------------------------------------------------------------------------
' Source: Relatório Financeiro Shopee/Relatório Financeiro Shopee (RFS)
' Purpose: Dynamically resizes an Excel ListObject to a single row.
'          Crucial for clearing data tables while preserving formulas/formats 
'          before importing new datasets via SQL/PowerQuery.
' -------------------------------------------------------------------------
Function vfResizeListObject(bvSheetCodeName As Worksheet, bvTableName As String) As Boolean
   vfResizeListObject = True
   On Error Goto Errhandler

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