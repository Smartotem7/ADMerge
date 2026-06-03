Attribute VB_Name = "ADMergeTool"
Option Explicit

Private Const SHEET_NAME As String = "Sheet1"
Private Const DIFF_SHEET_NAME As String = "diff"
Private Const HEADER_ROW As Long = 1
Private Const FIRST_DATA_ROW As Long = 2
Private Const COLOR_YELLOW As Long = vbYellow

' Entry point for the button on the VBA tool workbook.
' Assign a Form Control or ActiveX button to this macro.
Public Sub RunADMerge()
    Dim sourcePath As Variant
    Dim targetPath As Variant
    Dim sourceWorkbook As Workbook
    Dim targetWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim diffSheet As Worksheet
    Dim sourceRowsByKey As Object
    Dim targetLastRow As Long
    Dim targetRow As Long
    Dim diffRow As Long
    Dim changedRows As Long
    Dim errorRows As Long
    Dim matchedRows As Long
    Dim keyValue As String
    Dim sourceRow As Long
    Dim sourceO As String
    Dim targetO As String
    Dim shouldMerge As Boolean
    Dim hasChange As Boolean
    Dim errorMessage As String
    Dim originalCalculation As XlCalculation
    Dim settingsChanged As Boolean

    On Error GoTo HandleError

    originalCalculation = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    settingsChanged = True

    Set diffSheet = PrepareDiffSheet(ThisWorkbook)
    Set sourceWorkbook = Workbooks("基本設計書マージ対象進行管理表_0602_0930.xlsx")
    Set targetWorkbook = Workbooks("基本設計書マージ対象進行管理表.xlsx")
    Set sourceSheet = GetRequiredSheet(sourceWorkbook, SHEET_NAME)
    Set targetSheet = GetRequiredSheet(targetWorkbook, SHEET_NAME)
    Set sourceRowsByKey = BuildSourceIndex(sourceSheet)

    targetLastRow = LastUsedRow(targetSheet, "E")
    diffRow = 2

    For targetRow = FIRST_DATA_ROW To targetLastRow
        keyValue = NormalizeKey(targetSheet.Cells(targetRow, "E").Value)
        If Len(keyValue) > 0 And sourceRowsByKey.Exists(keyValue) Then
            matchedRows = matchedRows + 1
            sourceRow = CLng(sourceRowsByKey(keyValue))
            sourceO = NormalizeNValue(sourceSheet.Cells(sourceRow, "O").Value)
            targetO = NormalizeNValue(targetSheet.Cells(targetRow, "O").Value)
            shouldMerge = False
            errorMessage = vbNullString

            If sourceO = "BBX" And targetO = "BBX" Then
                shouldMerge = True
            ElseIf sourceO = vbNullString And targetO = "BBX" Then
                shouldMerge = True
            ElseIf sourceO = "BBX" And targetO = vbNullString Then
                errorMessage = "Error: source N is BBX, but target N is blank."
            ElseIf sourceO = vbNullString And targetO = vbNullString Then
                ' Skip this row.
            End If

            If shouldMerge Then
                hasChange = ApplyMergeAndWriteDiff(sourceSheet, sourceRow, targetSheet, targetRow, diffSheet, diffRow)
                If hasChange Then
                    changedRows = changedRows + 1
                    diffRow = diffRow + 1
                End If
            ElseIf Len(errorMessage) > 0 Then
                WriteErrorDiff targetSheet, targetRow, diffSheet, diffRow, errorMessage
                errorRows = errorRows + 1
                diffRow = diffRow + 1
            End If
        End If
    Next targetRow

    FormatDiffSheet diffSheet

    RestoreApplicationSettings originalCalculation, settingsChanged

    MsgBox "AD merge completed." & vbCrLf & _
           "Matched target rows: " & matchedRows & vbCrLf & _
           "Rows with changed values: " & changedRows & vbCrLf & _
           "Rows with errors: " & errorRows & vbCrLf & vbCrLf & _
           "Review the diff sheet in this VBA tool workbook." & vbCrLf & _
           "The target workbook remains open; save it if the result is correct.", _
           vbInformation, "AD Merge"
    Exit Sub

HandleError:
    RestoreApplicationSettings originalCalculation, settingsChanged
    MsgBox "AD merge stopped: " & Err.Description, vbCritical, "AD Merge"
End Sub

Private Sub RestoreApplicationSettings(ByVal originalCalculation As XlCalculation, ByVal settingsChanged As Boolean)
    If settingsChanged Then
        Application.Calculation = originalCalculation
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If
End Sub

Private Function PickExcelFile(ByVal dialogTitle As String) As Variant
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = dialogTitle
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel files", "*.xlsx;*.xlsm;*.xlsb;*.xls"
        If .Show <> -1 Then
            PickExcelFile = False
        Else
            PickExcelFile = .SelectedItems(1)
        End If
    End With
End Function

Private Function GetRequiredSheet(ByVal workbookToCheck As Workbook, ByVal requiredName As String) As Worksheet
    On Error GoTo MissingSheet

    Set GetRequiredSheet = workbookToCheck.Worksheets(requiredName)
    Exit Function

MissingSheet:
    Err.Raise vbObjectError + 1001, "ADMergeTool", _
              "Workbook '" & workbookToCheck.Name & "' must contain only one worksheet named '" & requiredName & "'."
End Function

Private Function PrepareDiffSheet(ByVal toolWorkbook As Workbook) As Worksheet
    Dim diffSheet As Worksheet

    On Error Resume Next
    Set diffSheet = toolWorkbook.Worksheets(DIFF_SHEET_NAME)
    On Error GoTo 0

    If diffSheet Is Nothing Then
        Set diffSheet = toolWorkbook.Worksheets.Add(After:=toolWorkbook.Worksheets(toolWorkbook.Worksheets.Count))
        diffSheet.Name = DIFF_SHEET_NAME
    End If

    diffSheet.Cells.Clear
    diffSheet.Range("A1:F1").Value = Array("Target Row", "Target E", "Target P After", "Target Q After", "Target X After", "Error")
    diffSheet.Range("A1:F1").Font.Bold = True
    Set PrepareDiffSheet = diffSheet
End Function

Private Function BuildSourceIndex(ByVal sourceSheet As Worksheet) As Object
    Dim rowsByKey As Object
    Dim lastRow As Long
    Dim rowNumber As Long
    Dim keyValue As String

    Set rowsByKey = CreateObject("Scripting.Dictionary")
    rowsByKey.CompareMode = vbTextCompare
    lastRow = LastUsedRow(sourceSheet, "E")

    For rowNumber = FIRST_DATA_ROW To lastRow
        keyValue = NormalizeKey(sourceSheet.Cells(rowNumber, "E").Value)
        If Len(keyValue) > 0 And Not rowsByKey.Exists(keyValue) Then
            rowsByKey.Add keyValue, rowNumber
        End If
    Next rowNumber

    Set BuildSourceIndex = rowsByKey
End Function

Private Function ApplyMergeAndWriteDiff( _
    ByVal sourceSheet As Worksheet, _
    ByVal sourceRow As Long, _
    ByVal targetSheet As Worksheet, _
    ByVal targetRow As Long, _
    ByVal diffSheet As Worksheet, _
    ByVal diffRow As Long) As Boolean

    Dim sourceQ As Variant
    Dim sourceR As Variant
    Dim sourceY As Variant
    Dim pChanged As Boolean
    Dim qChanged As Boolean
    Dim xChanged As Boolean

    sourceQ = sourceSheet.Cells(sourceRow, "Q").Value
    sourceR = sourceSheet.Cells(sourceRow, "R").Value
    sourceY = sourceSheet.Cells(sourceRow, "Y").Value

    pChanged = ValuesAreDifferent(targetSheet.Cells(targetRow, "P").Value, sourceQ)
    qChanged = ValuesAreDifferent(targetSheet.Cells(targetRow, "Q").Value, sourceR)
    xChanged = ValuesAreDifferent(targetSheet.Cells(targetRow, "X").Value, sourceY)

    If pChanged Or qChanged Or xChanged Then
        targetSheet.Cells(targetRow, "P").Value = sourceQ
        targetSheet.Cells(targetRow, "Q").Value = sourceR
        targetSheet.Cells(targetRow, "X").Value = sourceY

        diffSheet.Cells(diffRow, "A").Value = targetRow
        diffSheet.Cells(diffRow, "B").Value = targetSheet.Cells(targetRow, "E").Value
        diffSheet.Cells(diffRow, "C").Value = sourceQ
        diffSheet.Cells(diffRow, "D").Value = sourceR
        diffSheet.Cells(diffRow, "E").Value = sourceY

        If pChanged Then diffSheet.Cells(diffRow, "C").Interior.Color = COLOR_YELLOW
        If qChanged Then diffSheet.Cells(diffRow, "D").Interior.Color = COLOR_YELLOW
        If xChanged Then diffSheet.Cells(diffRow, "E").Interior.Color = COLOR_YELLOW
        ApplyMergeAndWriteDiff = True
    End If
End Function

Private Sub WriteErrorDiff( _
    ByVal targetSheet As Worksheet, _
    ByVal targetRow As Long, _
    ByVal diffSheet As Worksheet, _
    ByVal diffRow As Long, _
    ByVal errorMessage As String)

    diffSheet.Cells(diffRow, "A").Value = targetRow
    diffSheet.Cells(diffRow, "B").Value = targetSheet.Cells(targetRow, "E").Value
    diffSheet.Cells(diffRow, "C").Value = targetSheet.Cells(targetRow, "P").Value
    diffSheet.Cells(diffRow, "D").Value = targetSheet.Cells(targetRow, "Q").Value
    diffSheet.Cells(diffRow, "E").Value = targetSheet.Cells(targetRow, "X").Value
    diffSheet.Cells(diffRow, "F").Value = errorMessage
    diffSheet.Cells(diffRow, "F").Interior.Color = COLOR_YELLOW
End Sub

Private Function LastUsedRow(ByVal worksheetToCheck As Worksheet, ByVal columnLetter As String) As Long
    LastUsedRow = worksheetToCheck.Cells(worksheetToCheck.Rows.Count, columnLetter).End(xlUp).Row
    If LastUsedRow < FIRST_DATA_ROW Then LastUsedRow = FIRST_DATA_ROW - 1
End Function

Private Function NormalizeKey(ByVal cellValue As Variant) As String
    NormalizeKey = Trim$(CStr(cellValue))
End Function

Private Function NormalizeNValue(ByVal cellValue As Variant) As String
    NormalizeNValue = UCase$(Trim$(CStr(cellValue)))
End Function

Private Function ValuesAreDifferent(ByVal oldValue As Variant, ByVal newValue As Variant) As Boolean
    ValuesAreDifferent = (CStr(oldValue) <> CStr(newValue))
End Function

Private Sub FormatDiffSheet(ByVal diffSheet As Worksheet)
    With diffSheet
        .Columns("A:F").AutoFit
        .Rows(HEADER_ROW).AutoFilter
        .Activate
        .Range("A1").Select
    End With
End Sub
