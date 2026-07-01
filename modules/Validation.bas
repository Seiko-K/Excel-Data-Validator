Option Explicit

'==================================================
' Customer Validation Workflow
' Detect missing customer names and email addresses
' Highlight invalid records
'==================================================

Sub ValidateCustomerData()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ActiveSheet

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow

        If ws.Cells(i, 2).Value = "" Then
            ws.Cells(i, 2).Interior.Color = vbRed
        End If

        If ws.Cells(i, 3).Value = "" Then
            ws.Cells(i, 3).Interior.Color = vbYellow
        End If

    Next i

End Sub


'==================================================
' Duplicate Detection Workflow
' Highlight duplicate Customer IDs
' Uses Excel conditional formatting
'==================================================

Sub HighlightDuplicates()

    Dim rng As Range

    Set rng = Range("A2:A100")

    rng.FormatConditions.Delete

    rng.FormatConditions.AddUniqueValues

    rng.FormatConditions(
        rng.FormatConditions.Count
    ).DupeUnique = xlDuplicate

    rng.FormatConditions(
        rng.FormatConditions.Count
    ).Interior.Color = RGB(255, 200, 200)

End Sub
