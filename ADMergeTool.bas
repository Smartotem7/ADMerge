Attribute VB_Name = "ADMergeTool"

Option Explicit

Private Const HEADER_ROW As Long = 1
Private Const FIRST_DATA_ROW As Long = 2
Private Const COLOR_YELLOW As Long = vbYellow
Private Const CONFIG_KEY_SHEET_NAME As String = "SheetName"
Private Const CONFIG_KEY_SOURCE_WORKBOOK_NAME_CELL As String = "SourceWorkbookNameCell"
Private Const CONFIG_KEY_TARGET_WORKBOOK_NAME_CELL As String = "TargetWorkbookNameCell"
Private Const CONFIG_KEY_KEY_COLUMN As String = "KeyColumn"
Private Const CONFIG_KEY_STATUS_COLUMN As String = "StatusColumn"
Private Const CONFIG_KEY_SOURCE_MERGE_COLUMNS As String = "SourceMergeColumns"
Private Const CONFIG_KEY_TARGET_MERGE_COLUMNS As String = "TargetMergeColumns"
Private Const CONFIG_KEY_DIFF_SHEET_NAME As String = "DiffSheetName"
Private Const CONFIG_KEY_DIFF_COLUMNS As String = "DiffColumns"
Private Const CONFIG_KEY_DIFF_HEADERS As String = "DiffHeaders"
Private Const CONFIG_KEY_DIFF_CHANGED_COLUMNS As String = "DiffChangedColumns"
Private Const CONFIG_KEY_DIFF_ERROR_COLUMN As String = "DiffErrorColumn"

' Entry point for the first button on the VBA tool workbook.
' Assign a Form Control or ActiveX button to this macro.
Public Sub RunAAMerge()
    RunADMergeWithConfig CreateAAMergeConfig(), "AA Merge"
End Sub

' Entry point for the second button on the VBA tool workbook.
' Assign a Form Control or ActiveX button to this macro.
Public Sub RunADMerge()
    RunADMergeWithConfig CreateADMergeConfig(), "AD Merge"
End Sub

Private Sub RunADMergeWithConfig(ByVal mergeConfig As Object, ByVal dialogTitle As String)
    Dim toolConfigSheet As Worksheet
    Dim sourceWorkbookName As String
    Dim targetWorkbookName As String
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
    Dim sourceStatus As String
    Dim targetStatus As String
    Dim shouldMerge As Boolean
    Dim hasChange As Boolean
    Dim errorMessage As String
    Dim originalCalculation As XlCalculation
    Dim settingsChanged As Boolean

    On Error GoTo HandleError

    Set toolConfigSheet = GetToolConfigSheet(ThisWorkbook)
    sourceWorkbookName = ReadRequiredTextCell(toolConfigSheet, mergeConfig(CONFIG_KEY_SOURCE_WORKBOOK_NAME_CELL), "source workbook name")
    targetWorkbookName = ReadRequiredTextCell(toolConfigSheet, mergeConfig(CONFIG_KEY_TARGET_WORKBOOK_NAME_CELL), "target workbook name")

    originalCalculation = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    settingsChanged = True

    Set diffSheet = PrepareDiffSheet(ThisWorkbook, mergeConfig)
    Set sourceWorkbook = GetOpenWorkbook(sourceWorkbookName)
    Set targetWorkbook = GetOpenWorkbook(targetWorkbookName)
    Set sourceSheet = GetRequiredSheet(sourceWorkbook, mergeConfig(CONFIG_KEY_SHEET_NAME))
    Set targetSheet = GetRequiredSheet(targetWorkbook, mergeConfig(CONFIG_KEY_SHEET_NAME))
    Set sourceRowsByKey = BuildSourceIndex(sourceSheet, mergeConfig)

    targetLastRow = LastUsedRow(targetSheet, mergeConfig(CONFIG_KEY_KEY_COLUMN))
    diffRow = 2

    For targetRow = FIRST_DATA_ROW To targetLastRow
        keyValue = NormalizeKey(targetSheet.Cells(targetRow, mergeConfig(CONFIG_KEY_KEY_COLUMN)).Value)
        If Len(keyValue) > 0 And sourceRowsByKey.Exists(keyValue) Then
            matchedRows = matchedRows + 1
            sourceRow = CLng(sourceRowsByKey(keyValue))
            sourceStatus = NormalizeStatusValue(sourceSheet.Cells(sourceRow, mergeConfig(CONFIG_KEY_STATUS_COLUMN)).Value)
            targetStatus = NormalizeStatusValue(targetSheet.Cells(targetRow, mergeConfig(CONFIG_KEY_STATUS_COLUMN)).Value)
            shouldMerge = False
            errorMessage = vbNullString

            If sourceStatus = "BBX" And targetStatus = "BBX" Then
                shouldMerge = True
            ElseIf sourceStatus = vbNullString And targetStatus = "BBX" Then
                shouldMerge = True
            ElseIf sourceStatus = "BBX" And targetStatus = vbNullString Then
                errorMessage = "Error: source " & mergeConfig(CONFIG_KEY_STATUS_COLUMN) & " is BBX, but target " & mergeConfig(CONFIG_KEY_STATUS_COLUMN) & " is blank."
            ElseIf sourceStatus = vbNullString And targetStatus = vbNullString Then
                ' Skip this row.
            End If

            If shouldMerge Then
                hasChange = ApplyMergeAndWriteDiff(sourceSheet, sourceRow, targetSheet, targetRow, diffSheet, diffRow, mergeConfig)
                If hasChange Then
                    changedRows = changedRows + 1
                    diffRow = diffRow + 1
                End If
            ElseIf Len(errorMessage) > 0 Then
                WriteErrorDiff targetSheet, targetRow, diffSheet, diffRow, errorMessage, mergeConfig
                errorRows = errorRows + 1
                diffRow = diffRow + 1
            End If
        End If
    Next targetRow

    FormatDiffSheet diffSheet, mergeConfig

    RestoreApplicationSettings originalCalculation, settingsChanged

    MsgBox dialogTitle & " completed." & vbCrLf & _
           "Matched target rows: " & matchedRows & vbCrLf & _
           "Rows with changed values: " & changedRows & vbCrLf & _
           "Rows with errors: " & errorRows & vbCrLf & vbCrLf & _
           "Review the " & mergeConfig(CONFIG_KEY_DIFF_SHEET_NAME) & " sheet in this VBA tool workbook." & vbCrLf & _
           "The target workbook remains open; save it if the result is correct.", _
           vbInformation, dialogTitle
    Exit Sub

HandleError:
    RestoreApplicationSettings originalCalculation, settingsChanged
    MsgBox "AD merge stopped: " & Err.Description, vbCritical, dialogTitle
End Sub

Private Function CreateAAMergeConfig() As Object
    Dim mergeConfig As Object

    Set mergeConfig = CreateObject("Scripting.Dictionary")
    mergeConfig.CompareMode = vbTextCompare
    mergeConfig.Add CONFIG_KEY_SHEET_NAME, "Sheet1"
    mergeConfig.Add CONFIG_KEY_SOURCE_WORKBOOK_NAME_CELL, "B1"
    mergeConfig.Add CONFIG_KEY_TARGET_WORKBOOK_NAME_CELL, "B2"
    mergeConfig.Add CONFIG_KEY_DIFF_SHEET_NAME, "diffAA"
    mergeConfig.Add CONFIG_KEY_KEY_COLUMN, "E"
    mergeConfig.Add CONFIG_KEY_STATUS_COLUMN, "O"
    mergeConfig.Add CONFIG_KEY_SOURCE_MERGE_COLUMNS, Array("Q", "R", "Y")
    mergeConfig.Add CONFIG_KEY_TARGET_MERGE_COLUMNS, Array("P", "Q", "X")
    mergeConfig.Add CONFIG_KEY_DIFF_COLUMNS, Array("A", "B", "C", "D", "E", "F")
    mergeConfig.Add CONFIG_KEY_DIFF_HEADERS, Array("#", "Target E", "Target P After", "Target Q After", "Target X After", "Error")
    mergeConfig.Add CONFIG_KEY_DIFF_CHANGED_COLUMNS, Array("C", "D", "E")
    mergeConfig.Add CONFIG_KEY_DIFF_ERROR_COLUMN, "F"

    Set CreateAAMergeConfig = mergeConfig
End Function

Private Function CreateADMergeConfig() As Object
    Dim mergeConfig As Object

    Set mergeConfig = CreateObject("Scripting.Dictionary")
    mergeConfig.CompareMode = vbTextCompare
    mergeConfig.Add CONFIG_KEY_SHEET_NAME, "Sheet1"
    mergeConfig.Add CONFIG_KEY_SOURCE_WORKBOOK_NAME_CELL, "B6"
    mergeConfig.Add CONFIG_KEY_TARGET_WORKBOOK_NAME_CELL, "B7"
    mergeConfig.Add CONFIG_KEY_DIFF_SHEET_NAME, "diffAD"
    mergeConfig.Add CONFIG_KEY_KEY_COLUMN, "E"
    mergeConfig.Add CONFIG_KEY_STATUS_COLUMN, "N"
    mergeConfig.Add CONFIG_KEY_SOURCE_MERGE_COLUMNS, Array("P", "Q", "X")
    mergeConfig.Add CONFIG_KEY_TARGET_MERGE_COLUMNS, Array("O", "P", "W")
    mergeConfig.Add CONFIG_KEY_DIFF_COLUMNS, Array("A", "B", "C", "D", "E", "F")
    mergeConfig.Add CONFIG_KEY_DIFF_HEADERS, Array("#", "Target E", "Target O After", "Target P After", "Target W After", "Error")
    mergeConfig.Add CONFIG_KEY_DIFF_CHANGED_COLUMNS, Array("C", "D", "E")
    mergeConfig.Add CONFIG_KEY_DIFF_ERROR_COLUMN, "F"

    Set CreateADMergeConfig = mergeConfig
End Function

Private Sub RestoreApplicationSettings(ByVal originalCalculation As XlCalculation, ByVal settingsChanged As Boolean)
    If settingsChanged Then
        Application.Calculation = originalCalculation
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If
End Sub

Private Function GetToolConfigSheet(ByVal toolWorkbook As Workbook) As Worksheet
    Set GetToolConfigSheet = toolWorkbook.Worksheets(1)
End Function

Private Function ReadRequiredTextCell(ByVal worksheetToRead As Worksheet, ByVal cellAddress As String, ByVal valueDescription As String) As String
    ReadRequiredTextCell = Trim$(CStr(worksheetToRead.Range(cellAddress).Value))
    If Len(ReadRequiredTextCell) = 0 Then
        Err.Raise vbObjectError + 1000, "ADMergeTool", _
                  "Please enter the " & valueDescription & " in tool workbook cell " & cellAddress & "."
    End If
End Function

Private Function GetOpenWorkbook(ByVal workbookName As String) As Workbook
    On Error GoTo MissingWorkbook

    Set GetOpenWorkbook = Workbooks(workbookName)
    Exit Function

MissingWorkbook:
    Err.Raise vbObjectError + 1002, "ADMergeTool", _
              "Workbook '" & workbookName & "' must already be open."
End Function

Private Function GetRequiredSheet(ByVal workbookToCheck As Workbook, ByVal requiredName As String) As Worksheet
    On Error GoTo MissingSheet

    Set GetRequiredSheet = workbookToCheck.Worksheets(requiredName)
    Exit Function

MissingSheet:
    Err.Raise vbObjectError + 1001, "ADMergeTool", _
              "Workbook '" & workbookToCheck.Name & "' must contain only one worksheet named '" & requiredName & "'."
End Function

Private Function PrepareDiffSheet(ByVal toolWorkbook As Workbook, ByVal mergeConfig As Object) As Worksheet
    Dim diffSheet As Worksheet

    On Error Resume Next
    Set diffSheet = toolWorkbook.Worksheets(mergeConfig(CONFIG_KEY_DIFF_SHEET_NAME))
    On Error GoTo 0

    If diffSheet Is Nothing Then
        Set diffSheet = toolWorkbook.Worksheets.Add(After:=toolWorkbook.Worksheets(toolWorkbook.Worksheets.Count))
        diffSheet.Name = mergeConfig(CONFIG_KEY_DIFF_SHEET_NAME)
    End If

    diffSheet.Cells.Clear
    WriteDiffHeaders diffSheet, mergeConfig
    Set PrepareDiffSheet = diffSheet
End Function

Private Sub WriteDiffHeaders(ByVal diffSheet As Worksheet, ByVal mergeConfig As Object)
    Dim diffColumns As Variant
    Dim diffHeaders As Variant
    Dim columnIndex As Long

    diffColumns = mergeConfig(CONFIG_KEY_DIFF_COLUMNS)
    diffHeaders = mergeConfig(CONFIG_KEY_DIFF_HEADERS)

    For columnIndex = LBound(diffColumns) To UBound(diffColumns)
        diffSheet.Cells(HEADER_ROW, diffColumns(columnIndex)).Value = diffHeaders(columnIndex)
    Next columnIndex

    diffSheet.Range(diffColumns(LBound(diffColumns)) & HEADER_ROW & ":" & diffColumns(UBound(diffColumns)) & HEADER_ROW).Font.Bold = True
End Sub

Private Function BuildSourceIndex(ByVal sourceSheet As Worksheet, ByVal mergeConfig As Object) As Object
    Dim rowsByKey As Object
    Dim lastRow As Long
    Dim rowNumber As Long
    Dim keyValue As String

    Set rowsByKey = CreateObject("Scripting.Dictionary")
    rowsByKey.CompareMode = vbTextCompare
    lastRow = LastUsedRow(sourceSheet, mergeConfig(CONFIG_KEY_KEY_COLUMN))

    For rowNumber = FIRST_DATA_ROW To lastRow
        keyValue = NormalizeKey(sourceSheet.Cells(rowNumber, mergeConfig(CONFIG_KEY_KEY_COLUMN)).Value)
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
    ByVal diffRow As Long, _
    ByVal mergeConfig As Object) As Boolean

    Dim sourceMergeColumns As Variant
    Dim targetMergeColumns As Variant
    Dim diffChangedColumns As Variant
    Dim sourceValues() As Variant
    Dim changedColumns() As Boolean
    Dim columnIndex As Long
    Dim hasAnyChange As Boolean

    sourceMergeColumns = mergeConfig(CONFIG_KEY_SOURCE_MERGE_COLUMNS)
    targetMergeColumns = mergeConfig(CONFIG_KEY_TARGET_MERGE_COLUMNS)
    diffChangedColumns = mergeConfig(CONFIG_KEY_DIFF_CHANGED_COLUMNS)
    ReDim sourceValues(LBound(sourceMergeColumns) To UBound(sourceMergeColumns))
    ReDim changedColumns(LBound(sourceMergeColumns) To UBound(sourceMergeColumns))

    For columnIndex = LBound(sourceMergeColumns) To UBound(sourceMergeColumns)
        sourceValues(columnIndex) = sourceSheet.Cells(sourceRow, sourceMergeColumns(columnIndex)).Value
        changedColumns(columnIndex) = ValuesAreDifferent(targetSheet.Cells(targetRow, targetMergeColumns(columnIndex)).Value, sourceValues(columnIndex))
        hasAnyChange = hasAnyChange Or changedColumns(columnIndex)
    Next columnIndex

    If hasAnyChange Then
        For columnIndex = LBound(sourceMergeColumns) To UBound(sourceMergeColumns)
            targetSheet.Cells(targetRow, targetMergeColumns(columnIndex)).Value = sourceValues(columnIndex)
        Next columnIndex

        WriteBaseDiffColumns targetSheet, targetRow, diffSheet, diffRow, mergeConfig
        For columnIndex = LBound(sourceMergeColumns) To UBound(sourceMergeColumns)
            diffSheet.Cells(diffRow, diffChangedColumns(columnIndex)).Value = sourceValues(columnIndex)
            If changedColumns(columnIndex) Then diffSheet.Cells(diffRow, diffChangedColumns(columnIndex)).Interior.Color = COLOR_YELLOW
        Next columnIndex

        ApplyMergeAndWriteDiff = True
    End If
End Function

Private Sub WriteErrorDiff( _
    ByVal targetSheet As Worksheet, _
    ByVal targetRow As Long, _
    ByVal diffSheet As Worksheet, _
    ByVal diffRow As Long, _
    ByVal errorMessage As String, _
    ByVal mergeConfig As Object)

    Dim targetMergeColumns As Variant
    Dim diffChangedColumns As Variant
    Dim columnIndex As Long

    targetMergeColumns = mergeConfig(CONFIG_KEY_TARGET_MERGE_COLUMNS)
    diffChangedColumns = mergeConfig(CONFIG_KEY_DIFF_CHANGED_COLUMNS)

    WriteBaseDiffColumns targetSheet, targetRow, diffSheet, diffRow, mergeConfig
    For columnIndex = LBound(targetMergeColumns) To UBound(targetMergeColumns)
        diffSheet.Cells(diffRow, diffChangedColumns(columnIndex)).Value = targetSheet.Cells(targetRow, targetMergeColumns(columnIndex)).Value
    Next columnIndex

    diffSheet.Cells(diffRow, mergeConfig(CONFIG_KEY_DIFF_ERROR_COLUMN)).Value = errorMessage
    diffSheet.Cells(diffRow, mergeConfig(CONFIG_KEY_DIFF_ERROR_COLUMN)).Interior.Color = COLOR_YELLOW
End Sub

Private Sub WriteBaseDiffColumns( _
    ByVal targetSheet As Worksheet, _
    ByVal targetRow As Long, _
    ByVal diffSheet As Worksheet, _
    ByVal diffRow As Long, _
    ByVal mergeConfig As Object)

    Dim diffColumns As Variant

    diffColumns = mergeConfig(CONFIG_KEY_DIFF_COLUMNS)
    diffSheet.Cells(diffRow, diffColumns(0)).Value = targetRow
    diffSheet.Cells(diffRow, diffColumns(1)).Value = targetSheet.Cells(targetRow, mergeConfig(CONFIG_KEY_KEY_COLUMN)).Value
End Sub

Private Function LastUsedRow(ByVal worksheetToCheck As Worksheet, ByVal columnLetter As String) As Long
    LastUsedRow = worksheetToCheck.Cells(worksheetToCheck.Rows.Count, columnLetter).End(xlUp).Row
    If LastUsedRow < FIRST_DATA_ROW Then LastUsedRow = FIRST_DATA_ROW - 1
End Function

Private Function NormalizeKey(ByVal cellValue As Variant) As String
    NormalizeKey = Trim$(CStr(cellValue))
End Function

Private Function NormalizeStatusValue(ByVal cellValue As Variant) As String
    NormalizeStatusValue = UCase$(Trim$(CStr(cellValue)))
End Function

Private Function ValuesAreDifferent(ByVal oldValue As Variant, ByVal newValue As Variant) As Boolean
    ValuesAreDifferent = (CStr(oldValue) <> CStr(newValue))
End Function

Private Sub FormatDiffSheet(ByVal diffSheet As Worksheet, ByVal mergeConfig As Object)
    Dim diffColumns As Variant

    diffColumns = mergeConfig(CONFIG_KEY_DIFF_COLUMNS)

    With diffSheet
        .Columns(diffColumns(LBound(diffColumns)) & ":" & diffColumns(UBound(diffColumns))).AutoFit
        .Rows(HEADER_ROW).AutoFilter
        .Activate
        .Range(diffColumns(LBound(diffColumns)) & HEADER_ROW).Select
    End With
End Sub
