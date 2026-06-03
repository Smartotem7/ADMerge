# ADMerge VBA Tool

This repository contains a VBA module for an Excel macro-enabled workbook that merges values from a source workbook into a target workbook and writes the processing result to a `diff` sheet in the tool workbook.

## How to use

1. Create or open an Excel macro-enabled workbook (`.xlsm`) that will be the VBA tool workbook.
2. Import `ADMergeTool.bas` into the workbook's VBA project.
3. Add a button to the workbook and assign the `RunADMerge` macro to the button.
4. Click the button, then select the source file and target file when prompted.
5. Review the `diff` sheet in the VBA tool workbook.
6. Review and save the target workbook if the merge result is correct.

## Merge rules

Both the source workbook and target workbook must contain exactly the worksheet used by this tool, named `Sheet1`. The matching key is column `E` in both workbooks.

For each matching key:

- If source `N` is `BBX` and target `N` is `BBX`, the tool overwrites target values.
- If source `N` is blank and target `N` is `BBX`, the tool overwrites target values.
- If source `N` is `BBX` and target `N` is blank, the tool writes an error message to the `diff` sheet.
- If source `N` is blank and target `N` is blank, the tool skips the row.

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
