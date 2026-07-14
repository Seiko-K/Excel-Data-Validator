Option Explicit

'==================================================
' Research Validation Module
' Validate research metadata
'==================================================

Sub ValidateResearchData()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim publicationYear As Variant

    Set ws = ActiveSheet

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow

        ' Missing DOI
        If Trim(ws.Cells(i, 1).Value) = "" Then
            ws.Cells(i, 1).Interior.Color = RGB(255, 199, 206)
        End If

        ' Missing Title
        If Trim(ws.Cells(i, 2).Value) = "" Then
            ws.Cells(i, 2).Interior.Color = RGB(255, 199, 206)
        End If

        ' Invalid Publication Year
        publicationYear = ws.Cells(i, 3).Value

        If Not IsNumeric(publicationYear) Then

            ws.Cells(i, 3).Interior.Color = RGB(255, 199, 206)

        ElseIf publicationYear < 1900 _
            Or publicationYear > Year(Date) Then

            ws.Cells(i, 3).Interior.Color = RGB(255, 199, 206)

        End If

    Next i

    MsgBox "Research validation completed.", vbInformation

End Sub
