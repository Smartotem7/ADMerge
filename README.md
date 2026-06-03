# ADMerge VBA Tool

This repository contains a VBA module for an Excel macro-enabled workbook that merges values from a source workbook into a target workbook and writes the processing result to a `diff` sheet in the tool workbook.

## How to use

1. Create or open an Excel macro-enabled workbook (`.xlsm`) that will be the VBA tool workbook.
2. Import `ADMergeTool.bas` into the workbook's VBA project.
3. Add a button to the workbook and assign the `RunADMerge` macro to the button.
4. Open both the source workbook and target workbook in Excel before running the macro.
5. In the first worksheet of the tool workbook, enter the source workbook file name in cell `B1` and the target workbook file name in cell `B2`.
6. Click the button to run the merge.
7. Review the `diff` sheet in the VBA tool workbook.
8. Review and save the target workbook if the merge result is correct.

## Merge configuration

All workbook name cells and column positions are centralized in `CreateADMergeConfig` in `ADMergeTool.bas`. Copy that function and adjust the configured values when adding another merge flow for files whose columns are in different positions.

Current configuration:

| Setting | Value |
| --- | --- |
| Source workbook name cell | `B1` in the first worksheet of the tool workbook |
| Target workbook name cell | `B2` in the first worksheet of the tool workbook |
| Data worksheet name | `Sheet1` |
| Matching key column | `E` in both workbooks |
| Status column | `O` in both workbooks |
| Source merge columns | `Q`, `R`, `Y` |
| Target merge columns | `P`, `Q`, `X` |
| Diff output columns | `A:F` |

## Merge rules

Both the source workbook and target workbook must contain the worksheet used by this tool, named `Sheet1`. The matching key is column `E` in both workbooks.

For each matching key:

- If source status column `O` is `BBX` and target status column `O` is `BBX`, the tool overwrites target values.
- If source status column `O` is blank and target status column `O` is `BBX`, the tool overwrites target values.
- If source status column `O` is `BBX` and target status column `O` is blank, the tool writes an error message to the `diff` sheet.
- If source status column `O` is blank and target status column `O` is blank, the tool skips the row.

When overwriting, the tool maps these columns:

| Source column | Target column |
| --- | --- |
| `Q` | `P` |
| `R` | `Q` |
| `Y` | `X` |

## Diff sheet output

The tool creates or clears a `diff` sheet in the VBA tool workbook with these columns:

| Diff column | Value |
| --- | --- |
| `A` | Target row number |
| `B` | Target column `E` value |
| `C` | Target column `P` value after overwrite |
| `D` | Target column `Q` value after overwrite |
| `E` | Target column `X` value after overwrite |
| `F` | Error message |

Rows are written to `diff` only when at least one overwritten cell changes value, or when an error is detected. Changed output cells in columns `C:E` are highlighted yellow. Error cells in column `F` are highlighted yellow.
