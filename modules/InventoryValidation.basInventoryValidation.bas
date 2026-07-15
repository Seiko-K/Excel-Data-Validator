Option Explicit

'==================================================
' Inventory Validation Module
' Validate inventory data
'==================================================

Sub ValidateInventoryData()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim stockQuantity As Variant
    Dim reorderLevel As Variant

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

        stockQuantity = ws.Cells(i, 3).Value
        reorderLevel = ws.Cells(i, 4).Value

        ' Invalid Stock Quantity
        If Not IsNumeric(stockQuantity) Then

            ws.Cells(i, 3).Interior.Color = RGB(255, 199, 206)

        ElseIf stockQuantity < 0 Then

            ws.Cells(i, 3).Interior.Color = RGB(255, 199, 206)

        End If

        ' Invalid Reorder Level
        If Not IsNumeric(reorderLevel) Then

            ws.Cells(i, 4).Interior.Color = RGB(255, 199, 206)

        ElseIf reorderLevel < 0 Then

            ws.Cells(i, 4).Interior.Color = RGB(255, 199, 206)

        End If

    Next i

    MsgBox "Inventory validation completed.", vbInformation

End Sub
