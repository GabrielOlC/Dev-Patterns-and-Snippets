' -------------------------------------------------------------------------
' Source: Calculadora de precificação Sacos para Lixo (CPSL)
' Function: vfHexToRGB
' Purpose: Converts Hexadecimal color codes (Web standard) into Long (Excel RGB).
'          Enables the use of modern UI color palettes within VBA forms.
' Context: Developed to allow dynamic UI theming based on product categories.
' -------------------------------------------------------------------------
Function vfHexToRGB(ByVal bvHexColor As String) As Long
    On Error Goto ErrHandler

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