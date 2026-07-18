Option Explicit

'==================================================
' Excel Data Validator
' Cross-platform CSV validation report export
'
'【JP】
'顧客データの検証結果をCSVレポートとして出力する
'Windows / macOS 両対応モジュール。
'
' Design policy:
' - Windows: export CSV UTF-8 automatically beside the workbook
' - macOS: use Excel's native Save As dialog and guide users
'            to save the file in the Downloads folder
'
'【JP】
'設計方針:
' - Windows: 元ブックと同じフォルダへCSV UTF-8で自動保存
' - macOS: Excel標準の保存画面を使用し、
'          保存先をDownloadsフォルダーへ統一して案内
'==================================================


'==================================================
' Main Workflow
' Build and export a validation report
'
'【JP】
'検証結果を一時ブックへ作成し、
'OSごとの安定した方法でCSV出力する。
'==================================================
Public Sub GenerateCSVReport()

    Dim sourceWorksheet As Worksheet
    Dim reportWorkbook As Workbook
    Dim reportWorksheet As Worksheet
    Dim exportSucceeded As Boolean
    Dim exportFilePath As String
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo ExportError

    Set sourceWorksheet = ActiveSheet

    If GetCsvLastDataRow(sourceWorksheet, 1) < 2 Then

        MsgBox _
            "No customer records were found.", _
            vbInformation, _
            "CSV Report"

        Exit Sub

    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set reportWorkbook = Workbooks.Add(xlWBATWorksheet)
    Set reportWorksheet = reportWorkbook.Worksheets(1)

    reportWorksheet.Name = "Validation Report"

    BuildValidationReport sourceWorksheet, reportWorksheet

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    If IsMacExcel() Then

        exportSucceeded = ExportCsvForMac(reportWorkbook)

    Else

        exportSucceeded = ExportCsvForWindows( _
            reportWorkbook, _
            exportFilePath _
        )

    End If

    If Not reportWorkbook Is Nothing Then
        reportWorkbook.Close SaveChanges:=False
        Set reportWorkbook = Nothing
    End If

    If exportSucceeded Then

        If IsMacExcel() Then

            MsgBox _
                "CSV report exported successfully." & vbCrLf & vbCrLf & _
                "Please check your Downloads folder.", _
                vbInformation, _
                "CSV Report"

        Else

            MsgBox _
                "CSV report exported successfully." & vbCrLf & vbCrLf & _
                "File:" & vbCrLf & GetFileNameFromPath(exportFilePath) & vbCrLf & vbCrLf & _
                "Location:" & vbCrLf & GetFolderFromPath(exportFilePath), _
                vbInformation, _
                "CSV Report"

        End If

    End If

    Exit Sub

ExportError:

    errorNumber = Err.Number
    errorDescription = Err.Description

    On Error Resume Next

    If Not reportWorkbook Is Nothing Then
        reportWorkbook.Close SaveChanges:=False
    End If

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    On Error GoTo 0

    MsgBox _
        "The CSV report could not be exported." & vbCrLf & _
        "Error " & errorNumber & ": " & errorDescription, _
        vbExclamation, _
        "CSV Report"

End Sub


'==================================================
' Report Builder
' Create the validation issue table
'
'【JP】
'対象シートを検証し、CSV出力用の一覧を作成する。
'
' Output columns:
' Row, IssueType, Column, Value
'==================================================
Private Sub BuildValidationReport( _
    ByVal sourceWorksheet As Worksheet, _
    ByVal reportWorksheet As Worksheet _
)

    Dim lastRow As Long
    Dim sourceRow As Long
    Dim reportRow As Long
    Dim emailAddress As String

    lastRow = GetCsvLastDataRow(sourceWorksheet, 1)

    reportWorksheet.Cells(1, 1).Value = "Row"
    reportWorksheet.Cells(1, 2).Value = "IssueType"
    reportWorksheet.Cells(1, 3).Value = "Column"
    reportWorksheet.Cells(1, 4).Value = "Value"

    reportRow = 2

    For sourceRow = 2 To lastRow

        If Trim$(CStr(sourceWorksheet.Cells(sourceRow, 2).Value)) = "" Then

            WriteCsvIssueRow _
                reportWorksheet, _
                reportRow, _
                sourceRow, _
                "Missing", _
                "Name", _
                sourceWorksheet.Cells(sourceRow, 2).Value

            reportRow = reportRow + 1

        End If

        emailAddress = Trim$(CStr(sourceWorksheet.Cells(sourceRow, 3).Value))

        If emailAddress = "" Then

            WriteCsvIssueRow _
                reportWorksheet, _
                reportRow, _
                sourceRow, _
                "Missing", _
                "Email", _
                emailAddress

            reportRow = reportRow + 1

        ElseIf Not IsCsvEmailValid(emailAddress) Then

            WriteCsvIssueRow _
                reportWorksheet, _
                reportRow, _
                sourceRow, _
                "Invalid", _
                "Email", _
                emailAddress

            reportRow = reportRow + 1

        End If

    Next sourceRow

End Sub


'==================================================
' Windows Export
' Save CSV UTF-8 automatically beside the workbook
'
'【JP】
'Windowsでは元ブックと同じフォルダへ
'CSV UTF-8形式で自動保存する。
'==================================================
Private Function ExportCsvForWindows( _
    ByVal reportWorkbook As Workbook, _
    ByRef exportFilePath As String _
) As Boolean

    Dim exportFolder As String
    Dim exportFileName As String

    exportFolder = ThisWorkbook.Path

    If Len(exportFolder) = 0 Then

        MsgBox _
            "Please save the Excel workbook before exporting the CSV report.", _
            vbExclamation, _
            "CSV Report"

        ExportCsvForWindows = False
        Exit Function

    End If

    exportFileName = _
        "validation_report_" & Format$(Now, "yyyymmdd_hhnnss") & ".csv"

    exportFilePath = CombineCsvPath(exportFolder, exportFileName)

    Application.DisplayAlerts = False

    reportWorkbook.SaveAs _
        FileName:=exportFilePath, _
        FileFormat:=xlCSVUTF8, _
        CreateBackup:=False, _
        Local:=True

    Application.DisplayAlerts = True

    ExportCsvForWindows = True

End Function


'==================================================
' macOS Export
' Use Excel's native Save As dialog
'
'【JP】
'Mac版ExcelではSaveAsによる自動保存が環境依存で
'失敗する場合があるため、標準の保存画面を使用する。
'
' Users are instructed to save the file in Downloads.
' The code also attempts to open the dialog from Downloads,
' but Excel for macOS may restore a previous location.
'
'【JP】
'利用者にはDownloadsフォルダーへ保存するよう案内する。
'コードでもDownloadsを初期位置として設定を試みるが、
'Mac版Excelが前回の保存先を復元する場合がある。
'==================================================
Private Function ExportCsvForMac( _
    ByVal reportWorkbook As Workbook _
) As Boolean

    Dim userConfirmed As VbMsgBoxResult
    Dim saveSucceeded As Boolean
    Dim downloadsFolder As String
    Dim originalFolder As String

    downloadsFolder = GetMacDownloadsFolder()

    userConfirmed = MsgBox( _
        "The Save As dialog will open." & vbCrLf & vbCrLf & _
        "Please save the file in your Downloads folder." & vbCrLf & vbCrLf & _
        "File format:" & vbCrLf & _
        "CSV UTF-8 (Comma delimited) (.csv)" & vbCrLf & vbCrLf & _
        "File name:" & vbCrLf & _
        "validation_report.csv", _
        vbOKCancel + vbInformation, _
        "CSV Report" _
    )

    If userConfirmed <> vbOK Then
        ExportCsvForMac = False
        Exit Function
    End If

    ' Remember the current folder and try to switch to Downloads.
    '【JP】現在のフォルダを保持し、Downloadsへ移動を試みる。
    On Error Resume Next

    originalFolder = CurDir$

    If Len(downloadsFolder) > 0 Then
        ChDir downloadsFolder
    End If

    On Error GoTo 0

    reportWorkbook.Activate

    saveSucceeded = Application.Dialogs(xlDialogSaveAs).Show( _
        "validation_report.csv" _
    )

    ' Restore the previous current folder when possible.
    '【JP】可能であれば、保存画面を閉じた後に元のフォルダへ戻す。
    If Len(originalFolder) > 0 Then

        On Error Resume Next
        ChDir originalFolder
        On Error GoTo 0

    End If

    ExportCsvForMac = saveSucceeded

End Function


'==================================================
' macOS Downloads Folder
' Return the current user's Downloads folder
'
'【JP】
'現在のMacユーザーのDownloadsフォルダーを取得する。
'==================================================
Private Function GetMacDownloadsFolder() As String

    Dim homeFolder As String

    homeFolder = Environ$("HOME")

    If Len(homeFolder) > 0 Then
        GetMacDownloadsFolder = homeFolder & "/Downloads"
    Else
        GetMacDownloadsFolder = ""
    End If

End Function


'==================================================
' Platform Detection
' Return True when Excel is running on macOS
'
'【JP】
'ExcelがmacOS上で動作しているか判定する。
'==================================================
Private Function IsMacExcel() As Boolean

    IsMacExcel = _
        (InStr(1, Application.OperatingSystem, "Mac", vbTextCompare) > 0)

End Function


'==================================================
' Last Data Row
' Return the last used row in the specified column
'
'【JP】
'指定列の最終データ行を取得する。
'==================================================
Private Function GetCsvLastDataRow( _
    ByVal targetWorksheet As Worksheet, _
    ByVal targetColumn As Long _
) As Long

    GetCsvLastDataRow = _
        targetWorksheet.Cells( _
            targetWorksheet.Rows.Count, _
            targetColumn _
        ).End(xlUp).Row

End Function


'==================================================
' Email Validation
' Perform a lightweight email format check
'
'【JP】
'CSV出力時に使用する簡易メール形式チェック。
'==================================================
Private Function IsCsvEmailValid( _
    ByVal emailAddress As String _
) As Boolean

    Dim atPosition As Long
    Dim dotPosition As Long

    emailAddress = Trim$(emailAddress)

    If Len(emailAddress) = 0 Then
        IsCsvEmailValid = False
        Exit Function
    End If

    If InStr(emailAddress, " ") > 0 Then
        IsCsvEmailValid = False
        Exit Function
    End If

    atPosition = InStr(1, emailAddress, "@")

    If atPosition <= 1 Then
        IsCsvEmailValid = False
        Exit Function
    End If

    If InStr(atPosition + 1, emailAddress, "@") > 0 Then
        IsCsvEmailValid = False
        Exit Function
    End If

    dotPosition = InStrRev(emailAddress, ".")

    If dotPosition <= atPosition + 1 Then
        IsCsvEmailValid = False
        Exit Function
    End If

    If dotPosition = Len(emailAddress) Then
        IsCsvEmailValid = False
        Exit Function
    End If

    IsCsvEmailValid = True

End Function


'==================================================
' Issue Row Writer
' Write one validation issue to the report worksheet
'
'【JP】
'検証エラー1件をレポートシートへ書き込む。
'==================================================
Private Sub WriteCsvIssueRow( _
    ByVal reportWorksheet As Worksheet, _
    ByVal reportRow As Long, _
    ByVal sourceRow As Long, _
    ByVal issueType As String, _
    ByVal columnName As String, _
    ByVal issueValue As Variant _
)

    reportWorksheet.Cells(reportRow, 1).Value = sourceRow
    reportWorksheet.Cells(reportRow, 2).Value = issueType
    reportWorksheet.Cells(reportRow, 3).Value = columnName
    reportWorksheet.Cells(reportRow, 4).Value = issueValue

End Sub


'==================================================
' Path Builder
' Combine a folder path and file name safely
'
'【JP】
'フォルダパスとファイル名を安全に連結する。
'==================================================
Private Function CombineCsvPath( _
    ByVal folderPath As String, _
    ByVal fileName As String _
) As String

    Dim separator As String

    separator = Application.PathSeparator

    If Right$(folderPath, 1) = separator Then
        CombineCsvPath = folderPath & fileName
    Else
        CombineCsvPath = folderPath & separator & fileName
    End If

End Function


'==================================================
' File Name Extractor
' Return only the file name from a full path
'
'【JP】
'フルパスからファイル名だけを取得する。
'==================================================
Private Function GetFileNameFromPath( _
    ByVal fullPath As String _
) As String

    Dim separatorPosition As Long

    separatorPosition = InStrRev(fullPath, Application.PathSeparator)

    If separatorPosition > 0 Then
        GetFileNameFromPath = Mid$(fullPath, separatorPosition + 1)
    Else
        GetFileNameFromPath = fullPath
    End If

End Function


'==================================================
' Folder Extractor
' Return only the folder path from a full path
'
'【JP】
'フルパスからフォルダ部分だけを取得する。
'==================================================
Private Function GetFolderFromPath( _
    ByVal fullPath As String _
) As String

    Dim separatorPosition As Long

    separatorPosition = InStrRev(fullPath, Application.PathSeparator)

    If separatorPosition > 0 Then
        GetFolderFromPath = Left$(fullPath, separatorPosition - 1)
    Else
        GetFolderFromPath = ""
    End If

End Function
