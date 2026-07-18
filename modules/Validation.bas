Option Explicit

'==================================================
' Excel Data Validator
' Customer data validation and report generation
'
'【JP】
'顧客データの検証、強調表示、重複検出、
'サマリー表示、CSVレポート出力を担当するモジュール。
'
' Expected worksheet columns:
' Column A: Customer ID
' Column B: Customer Name
' Column C: Email Address
'
'【JP】
'対象シートの列構成は次を前提とする。
'A列：Customer ID
'B列：Customer Name
'C列：Email Address
'==================================================


'==================================================
' Cleanup Workflow
' Clear validation highlights and conditional formatting
' Reset the worksheet before running validation again
'
'【JP】
'再検証の前に、前回の背景色と重複チェック用の
'条件付き書式を削除してシートを初期状態へ戻す。
'==================================================
Public Sub ClearValidationHighlights()

    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ActiveSheet
    lastRow = GetLastDataRow(ws, 1)

    ' Stop when the worksheet contains headers only
    '【JP】データ行が存在しない場合は処理を終了する。
    If lastRow < 2 Then Exit Sub

    ' Clear validation highlight colors
    '【JP】A列からC列までの検証用背景色を解除する。
    ws.Range("A2:C" & lastRow).Interior.ColorIndex = xlNone

    ' Remove duplicate detection rules from Customer ID
    '【JP】Customer ID列の条件付き書式を削除する。
    ws.Range("A2:A" & lastRow).FormatConditions.Delete

End Sub


'==================================================
' Customer Validation Workflow
' Detect missing customer names and email addresses
' Highlight invalid records
'
'【JP】
'顧客名とメールアドレスの未入力を検出し、
'対象セルを色分けして表示する。
'
' Red: Missing customer name
' Yellow: Missing email address
'==================================================
Public Sub ValidateCustomerData()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long

    Set ws = ActiveSheet
    lastRow = GetLastDataRow(ws, 1)

    If lastRow < 2 Then Exit Sub

    For rowIndex = 2 To lastRow

        ' Check required customer name
        '【JP】顧客名が空欄または空白のみの場合は赤色で強調する。
        If Trim$(CStr(ws.Cells(rowIndex, 2).Value)) = "" Then
            ws.Cells(rowIndex, 2).Interior.Color = vbRed
        End If

        ' Check email address
        '【JP】メールアドレスが空欄または空白のみの場合は黄色で強調する。
        If Trim$(CStr(ws.Cells(rowIndex, 3).Value)) = "" Then
            ws.Cells(rowIndex, 3).Interior.Color = vbYellow
        End If

    Next rowIndex

End Sub


'==================================================
' Email Validation Workflow
' Detect invalid email addresses
'
'【JP】
'値が入力されているメールアドレスのみを対象に、
'基本的なドメイン形式を検証する。
'
' Missing email addresses are handled separately by
' ValidateCustomerData.
'
'【JP】
'未入力メールはValidateCustomerDataで検出するため、
'この処理では形式不正のみを対象とする。
'==================================================
Public Sub ValidateEmails()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim emailAddress As String

    Set ws = ActiveSheet
    lastRow = GetLastDataRow(ws, 1)

    If lastRow < 2 Then Exit Sub

    For rowIndex = 2 To lastRow

        emailAddress = Trim$(CStr(ws.Cells(rowIndex, 3).Value))

        If emailAddress <> "" Then

            If Not IsValidEmailDomain(emailAddress) Then

                ' Highlight invalid email format
                '【JP】形式が不正なメールアドレスを薄い赤色で強調する。
                ws.Cells(rowIndex, 3).Interior.Color = RGB(255, 199, 206)

            End If

        End If

    Next rowIndex

End Sub


'==================================================
' Email Domain Validation
' Validate the basic structure of an email address
'
' Return value:
' True  = Valid basic structure
' False = Invalid basic structure
'
'【JP】
'メールアドレスの基本構造を検証する関数。
'厳密なRFC準拠ではなく、業務データ検証向けの
'簡易チェックとして使用する。
'==================================================
Public Function IsValidEmailDomain( _
    ByVal emailAddress As String _
) As Boolean

    Dim atPosition As Long
    Dim domainPart As String

    emailAddress = Trim$(emailAddress)

    ' Reject empty values
    '【JP】空欄は無効とする。
    If emailAddress = "" Then Exit Function

    ' The @ symbol must not be the first character
    '【JP】@より前に文字が必要。
    atPosition = InStr(1, emailAddress, "@", vbBinaryCompare)

    If atPosition <= 1 Then Exit Function

    ' Only one @ symbol is allowed
    '【JP】@が複数存在するメールアドレスは無効。
    If atPosition <> InStrRev(emailAddress, "@", -1, vbBinaryCompare) Then
        Exit Function
    End If

    domainPart = Mid$(emailAddress, atPosition + 1)

    ' A domain must exist after @
    '【JP】@より後ろにドメインが必要。
    If domainPart = "" Then Exit Function

    ' A domain must include at least one period
    '【JP】ドメインには少なくとも1つのピリオドが必要。
    If InStr(1, domainPart, ".", vbBinaryCompare) = 0 Then Exit Function

    ' A domain cannot start or end with a period
    '【JP】ドメインの先頭・末尾がピリオドの場合は無効。
    If Left$(domainPart, 1) = "." _
        Or Right$(domainPart, 1) = "." Then

        Exit Function

    End If

    ' Consecutive periods are not allowed
    '【JP】連続するピリオドを含むドメインは無効。
    If InStr(1, domainPart, "..", vbBinaryCompare) > 0 Then Exit Function

    IsValidEmailDomain = True

End Function


'==================================================
' Duplicate Detection Workflow
' Highlight duplicate Customer IDs
' Uses Excel conditional formatting
'
'【JP】
'Customer ID列に条件付き書式を設定し、
'重複しているIDを薄い赤色で強調する。
'
' The range is dynamically determined from the data.
'【JP】
'固定の100行ではなく、実データの最終行までを対象とする。
'==================================================
Public Sub HighlightDuplicates()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim targetRange As Range
    Dim duplicateRule As UniqueValues
    Set ws = ActiveSheet
    lastRow = GetLastDataRow(ws, 1)

    If lastRow < 2 Then Exit Sub

    Set targetRange = ws.Range("A2:A" & lastRow)

    ' Remove old rules before adding a new one
    '【JP】条件付き書式の重複登録を防ぐため、既存ルールを削除する。
    targetRange.FormatConditions.Delete

    targetRange.FormatConditions.AddUniqueValues

    Set duplicateRule = _
        targetRange.FormatConditions(targetRange.FormatConditions.Count)

    duplicateRule.DupeUnique = xlDuplicate
    duplicateRule.Interior.Color = RGB(255, 200, 200)

End Sub


'==================================================
' Report Generation Workflow
' Display a validation summary
'
'【JP】
'顧客名とメールアドレスの未入力件数を集計し、
'メッセージボックスでサマリーを表示する。
'==================================================
Public Sub GenerateValidationReport()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim missingNames As Long
    Dim missingEmails As Long
    Dim invalidEmails As Long
    Dim emailAddress As String

    Set ws = ActiveSheet
    lastRow = GetLastDataRow(ws, 1)

    If lastRow < 2 Then

        MsgBox _
            "No customer records were found.", _
            vbInformation, _
            "Validation Summary"

        Exit Sub

    End If

    For rowIndex = 2 To lastRow

        ' Count missing customer names
        '【JP】顧客名の未入力件数を集計する。
        If Trim$(CStr(ws.Cells(rowIndex, 2).Value)) = "" Then
            missingNames = missingNames + 1
        End If

        emailAddress = Trim$(CStr(ws.Cells(rowIndex, 3).Value))

        ' Count missing and invalid email addresses separately
        '【JP】メール未入力と形式不正を分けて集計する。
        If emailAddress = "" Then

            missingEmails = missingEmails + 1

        ElseIf Not IsValidEmailDomain(emailAddress) Then

            invalidEmails = invalidEmails + 1

        End If

    Next rowIndex

    MsgBox _
        "Validation Summary" & vbCrLf & _
        vbCrLf & _
        "Total Records: " & (lastRow - 1) & vbCrLf & _
        "Missing Names: " & missingNames & vbCrLf & _
        "Missing Emails: " & missingEmails & vbCrLf & _
        "Invalid Emails: " & invalidEmails, _
        vbInformation, _
        "Excel Data Validator"

End Sub


'==================================================
' Last Row Utility
' Return the last populated row in a specified column
'
'【JP】
'指定列の最終データ行を返す共通関数。
'複数の処理で同じ最終行取得コードを繰り返さないために使用する。
'==================================================
Private Function GetLastDataRow( _
    ByVal ws As Worksheet, _
    ByVal columnNumber As Long _
) As Long

    GetLastDataRow = _
        ws.Cells(ws.Rows.Count, columnNumber).End(xlUp).Row

End Function
