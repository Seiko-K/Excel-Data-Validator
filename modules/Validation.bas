Option Explicit

'==================================================
' Cleanup Workflow
' Clear validation highlights and conditional formatting
' Reset the worksheet before running validation again
'==================================================
Sub ClearValidationHighlights()

    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ActiveSheet

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ws.Range("A2:C" & lastRow).Interior.ColorIndex = xlNone
    ws.Range("A2:A" & lastRow).FormatConditions.Delete

End Sub


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
' Email Validation Workflow
' Detect invalid email addresses
'==================================================
Sub ValidateEmails()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row

    Dim i As Long
    Dim emailAddress As String

    For i = 2 To lastRow

        emailAddress = Trim(ws.Cells(i, 3).Value)

        If emailAddress <> "" Then

            If Not IsValidEmailDomain(emailAddress) Then

                ws.Cells(i, 3).Interior.Color = RGB(255, 199, 206)

            End If

        End If

    Next i

End Sub


Function IsValidEmailDomain(ByVal emailAddress As String) As Boolean

    Dim atPosition As Long
    Dim domainPart As String

    emailAddress = Trim(emailAddress)

    If emailAddress = "" Then
        IsValidEmailDomain = False
        Exit Function
    End If

    atPosition = InStr(1, emailAddress, "@")

    If atPosition <= 1 Then
        IsValidEmailDomain = False
        Exit Function
    End If

    If atPosition <> InStrRev(emailAddress, "@") Then
        IsValidEmailDomain = False
        Exit Function
    End If

    domainPart = Mid(emailAddress, atPosition + 1)

    If domainPart = "" Then
        IsValidEmailDomain = False
        Exit Function
    End If

    If InStr(1, domainPart, ".") = 0 Then
        IsValidEmailDomain = False
        Exit Function
    End If

    If Left(domainPart, 1) = "." Or Right(domainPart, 1) = "." Then
        IsValidEmailDomain = False
        Exit Function
    End If

    If InStr(1, domainPart, "..") > 0 Then
        IsValidEmailDomain = False
        Exit Function
    End If

    IsValidEmailDomain = True

End Function


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

    rng.FormatConditions(rng.FormatConditions.Count).DupeUnique = xlDuplicate
    rng.FormatConditions(rng.FormatConditions.Count).Interior.Color = RGB(255, 200, 200)

End Sub


'==================================================
' Report Generation Workflow
' Display validation summary
'==================================================

Sub GenerateValidationReport()

    Dim ws As Worksheet
    Dim lastRow As Long

    Dim MissingNames As Long
    Dim MissingEmails As Long

    Set ws = ActiveSheet

    lastRow = ws.Cells(
        ws.Rows.Count,
        1
    ).End(xlUp).Row

    Dim i As Long

    For i = 2 To lastRow

        If ws.Cells(i, 2).Value = "" Then

            MissingNames =
                MissingNames + 1

        End If

        If ws.Cells(i, 3).Value = "" Then

            MissingEmails =
                MissingEmails + 1

        End If

    Next i

    MsgBox _

        "Validation Summary" _
        & vbCrLf _
        & vbCrLf _
        & "Missing Names : " _
        & MissingNames _
        & vbCrLf _
        & "Missing Emails : " _
        & MissingEmails,

        vbInformation

End Sub


Sub GenerateCSVReport()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim filePath As String
    filePath = ThisWorkbook.Path & "\validation_report.csv"

    Dim fileNum As Integer
    fileNum = FreeFile

    Open filePath For Output As #fileNum

    Print #fileNum, "Row,IssueType,Column,Value"

    Dim i As Long

    For i = 2 To lastRow

        If Trim(ws.Cells(i, 2).Value) = "" Then
            Print #fileNum, i & ",Missing,Name,"
        End If

        If Trim(ws.Cells(i, 3).Value) = "" Then
            Print #fileNum, i & ",Missing,Email,"
        End If

        If InStr(ws.Cells(i, 3).Value, "@") = 0 _
           And Trim(ws.Cells(i, 3).Value) <> "" Then

            Print #fileNum, i & _
            ",Invalid,Email," & _
            ws.Cells(i, 3).Value

        End If

    Next i

    Close #fileNum

    MsgBox _
        "CSV Report exported." & vbCrLf & _
        filePath, _
        vbInformation

End Sub
