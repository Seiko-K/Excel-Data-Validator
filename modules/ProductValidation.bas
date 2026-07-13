Option Explicit

'==================================================
' Product Validation Module
' Validate product catalog data
'==================================================

Sub ValidateProductData()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ActiveSheet

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow

        ' Missing SKU
        If Trim(ws.Cells(i, 1).Value) = "" Then
            ws.Cells(i, 1).Interior.Color = RGB(255, 199, 206)
        End If

        ' Missing Product Name
        If Trim(ws.Cells(i, 2).Value) = "" Then
            ws.Cells(i, 2).Interior.Color = RGB(255, 199, 206)
        End If

        ' Invalid Price
        If Not IsNumeric(ws.Cells(i, 3).Value) _
           Or ws.Cells(i, 3).Value < 0 Then

            ws.Cells(i, 3).Interior.Color = RGB(255, 199, 206)

        End If

    Next i

    MsgBox "Product validation completed.", vbInformation

End Sub
