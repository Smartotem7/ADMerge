# ADMerge VBA Tool

This repository contains a VBA module for an Excel macro-enabled workbook that merges values from a source workbook into a target workbook and writes the processing result to a merge-specific diff sheet in the tool workbook.

## How to use

1. Create or open an Excel macro-enabled workbook (`.xlsm`) that will be the VBA tool workbook.
2. Import `ADMergeTool.bas` into the workbook's VBA project.
3. Add a button to the workbook and assign the `RunAAMerge` macro to the button. Optionally add a second button and assign the `RunADMerge` macro to run the AD merge configuration.
4. Open both the source workbook and target workbook in Excel before running the macro.
5. In the first worksheet of the tool workbook, enter the AA source workbook file name in cell `B1` and the AA target workbook file name in cell `B2`. For AD merge, enter the AD source workbook file name in cell `B6` and the AD target workbook file name in cell `B7`.
6. Click the desired button to run that merge.
7. Review the generated diff sheet in the VBA tool workbook: `diffAA` for `RunAAMerge`, or `diffAD` for `RunADMerge`.
8. Review and save the target workbook if the merge result is correct.

## Merge configuration

Workbook name cells, diff sheet names, and column positions are centralized in `CreateAAMergeConfig` and `CreateADMergeConfig` in `ADMergeTool.bas`. The `RunAAMerge` macro uses `CreateAAMergeConfig`, and the `RunADMerge` macro uses `CreateADMergeConfig`. Adjust the configured values when adding another merge flow for files whose columns are in different positions.

Current AA merge configuration:

| Setting | Value |
| --- | --- |
| Source workbook name cell | `B1` in the first worksheet of the tool workbook |
| Target workbook name cell | `B2` in the first worksheet of the tool workbook |
| Diff sheet | `diffAA` |
| Data worksheet name | `Sheet1` |
| Matching key column | `E` in both workbooks |
| Status column | `O` in both workbooks |
| Source merge columns | `Q`, `R`, `Y` |
| Target merge columns | `P`, `Q`, `X` |
| Diff output columns | `A:F` |

Current AD merge configuration:

| Setting | Value |
| --- | --- |
| Source workbook name cell | `B6` in the first worksheet of the tool workbook |
| Target workbook name cell | `B7` in the first worksheet of the tool workbook |
| Diff sheet | `diffAD` |
| Data worksheet name | `Sheet1` |
| Matching key column | `E` in both workbooks |
| Status column | `N` in both workbooks |
| Source merge columns | `P`, `Q`, `X` |
| Target merge columns | `O`, `P`, `W` |
| Diff output columns | `A:F` |

## Merge rules

Both the source workbook and target workbook must contain the worksheet used by this tool, named `Sheet1`. The matching key is column `E` in both workbooks.

For each matching key:

- If the configured source status column is `BBX` and the configured target status column is `BBX`, the tool overwrites target values.
- If the configured source status column is blank and the configured target status column is `BBX`, the tool overwrites target values.
- If the configured source status column is `BBX` and the configured target status column is blank, the tool writes an error message to the configured diff sheet.
- If the configured source status column is blank and the configured target status column is blank, the tool skips the row.

When AA merge overwrites, the tool maps these columns:

| Source column | Target column |
| --- | --- |
| `Q` | `P` |
| `R` | `Q` |
| `Y` | `X` |

When AD merge overwrites, the tool maps these columns:

| Source column | Target column |
| --- | --- |
| `P` | `O` |
| `Q` | `P` |
| `X` | `W` |

## Diff sheet output

The tool creates or clears the configured diff sheet in the VBA tool workbook (`diffAA` for `RunAAMerge`, `diffAD` for `RunADMerge`) with these columns:

| Diff column | AA merge value | AD merge value |
| --- | --- | --- |
| `A` | Target row number | Target row number |
| `B` | Target column `E` value | Target column `E` value |
| `C` | Target column `P` value after overwrite | Target column `O` value after overwrite |
| `D` | Target column `Q` value after overwrite | Target column `P` value after overwrite |
| `E` | Target column `X` value after overwrite | Target column `W` value after overwrite |
| `F` | Error message | Error message |

Rows are written to the configured diff sheet only when at least one overwritten cell changes value, or when an error is detected. Changed output cells in columns `C:E` are highlighted yellow. Error cells in column `F` are highlighted yellow.
