<#
.SYNOPSIS
This script converts an Excel file from .xls to .xlsx format if necessary, detects the header row dynamically, and transforms the data based on specific criteria.

.DESCRIPTION
The script performs the following steps:
1. Converts an .xls file to .xlsx format if necessary, adding the suffix '_ConvertedToXLSX'.
2. Automatically detects if the input file is already in .xlsx format and skips the conversion step.
3. Dynamically detects the header row and column in the Excel file.
4. Imports the data starting from the detected header row.
5. Transforms the data by remapping columns and updating specific fields if certain conditions are met.
6. Exports the transformed data to a new Excel file with the suffix '_Transformed'. If the file was converted, the suffix '_ConvertedToXLSX_Transformed' is used.

.PARAMETER inputFilePath
The path to the input Excel file (.xls or .xlsx).

.NOTES
Ensure the ImportExcel module is installed before running this script.
#>

param (
    [string]$inputFilePath
)

Clear-Host

# Function to validate parameters
function Validate-Parameters {
    param (
        [string]$inputFilePath
    )

    if (-not $inputFilePath) {
        throw "Please provide the input file path as a parameter."
    }

    if (-not (Test-Path $inputFilePath)) {
        throw "The file path '$inputFilePath' does not exist."
    }
}

# Function to convert .xls to .xlsx
function Convert-XlsToXlsx {
    param (
        [string]$inputFilePath
    )

    $baseFileName = [System.IO.Path]::GetFileNameWithoutExtension($inputFilePath)
    $directoryPath = [System.IO.Path]::GetDirectoryName($inputFilePath)
    $xlsxFilePath = "$directoryPath\$baseFileName`_ConvertedFromXls.xlsx"

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    try {
        $workbook = $excel.Workbooks.Open($inputFilePath)
        $workbook.SaveAs($xlsxFilePath, 51)  # 51 is the Excel constant for .xlsx
        $workbook.Close()
        Write-Host "Conversion complete: $inputFilePath -> $xlsxFilePath"
    } catch {
        Write-Host "Error converting file '$inputFilePath' to .xlsx: $_"
    } finally {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
    }

    return $xlsxFilePath
}

# Function to detect header row and column
function Detect-Header {
    param (
        [string]$excelFilePath
    )

    $headerColumns = @('S No.', 'Value Date', 'Transaction Date', 'Cheque Number', 'Transaction Remarks', 'Withdrawal Amount (INR )', 'Deposit Amount (INR )', 'Balance (INR )')
    $startRow = 0
    $startColumn = 0

    try {
        $excelPackage = Open-ExcelPackage -Path $excelFilePath
        foreach ($worksheet in $excelPackage.Workbook.Worksheets) {
            for ($rowIndex = 1; $rowIndex -le $worksheet.Dimension.End.Row; $rowIndex++) {
                $isRowBlank = $true
                for ($colIndex = 1; $colIndex -le $worksheet.Dimension.End.Column; $colIndex++) {
                    if ($worksheet.Cells[$rowIndex, $colIndex].Value -ne $null) {
                        $isRowBlank = $false
                        break
                    }
                }

                if ($isRowBlank) {
                    continue
                }

                for ($startColIndex = 1; $startColIndex -le 4; $startColIndex++) {
                    $headerMatch = $true
                    for ($colIndex = $startColIndex; $colIndex -le $headerColumns.Length + $startColIndex - 1; $colIndex++) {
                        $cellValue = $worksheet.Cells[$rowIndex, $colIndex].Value
                        if ($cellValue -ne $headerColumns[$colIndex - $startColIndex]) {
                            $headerMatch = $false
                            break
                        }
                    }

                    if ($headerMatch) {
                        $startRow = $rowIndex + 1
                        $startColumn = $startColIndex
                        break
                    }
                }

                if ($startRow -ne 0) {
                    break
                }
            }

            if ($startRow -ne 0) {
                break
            }
        }
    } catch {
        Write-Host "Error detecting header row in file '$excelFilePath': $_"
    } finally {
        $excelPackage.Dispose()
    }

    if ($startRow -eq 0) {
        Write-Host "Header row not found in file '$excelFilePath'. Please check the file format."
        return $null
    }

    return @{ StartRow = $startRow; StartColumn = $startColumn }
}

# Function to transform data and handle concatenation of split 'Transaction Remarks'
function Transform-Data {
    param (
        [string]$inputFilePath,
        [int]$startRow,
        [int]$startColumn,
        [string]$MOP,
        [string]$transformedFilePath
    )

    # Correct header names as per the file
    $customHeaders = @('S No.', 'Value Date', 'Transaction Date', 'Cheque Number', 'Transaction Remarks', 'Withdrawal Amount (INR )', 'Deposit Amount (INR )', 'Balance (INR )')

    # Import data from the Excel file
    $oldData = Import-Excel -Path $inputFilePath -StartRow $startRow -StartColumn $startColumn -HeaderName $customHeaders
    $newData = @()

    $previousRow = $null  # Keep track of the previous row for concatenation

    foreach ($row in $oldData) {
        # Check if the current row has only 'Transaction Remarks' and all other columns are blank
        $isBlankRow = -not $row.'S No.' -and -not $row.'Value Date' -and -not $row.'Transaction Date' -and -not $row.'Cheque Number' `
                      -and -not $row.'Withdrawal Amount (INR )' -and -not $row.'Deposit Amount (INR )' -and -not $row.'Balance (INR )'

        if ($isBlankRow -and $row.'Transaction Remarks') {
            # If this row contains only 'Transaction Remarks', append it to the previous row
            if ($previousRow -ne $null) {
                #$previousRow.'Transaction Remarks' += " " + $row.'Transaction Remarks'
                #$previousRow.'Narration' += " " + $row.'Transaction Remarks'
                $previousRow.'Narration' += $row.'Transaction Remarks'
            }
        } else {
            # If the row has meaningful data, process it as a new row

            # Create a new custom object to represent the transformed row
            $newRow = [PSCustomObject]@{
                Date = $row.'Value Date'
                Narration = $row.'Transaction Remarks'
                Item = ""  # Empty by default
                Category = ""  # Empty by default
                Place = ""  # Empty by default
                Freq = ""  # Empty by default
                For = ""  # Empty by default
                MOP = $MOP  # Retrieved from filename
                'Amt (Dr)' = $row.'Withdrawal Amount (INR )'
                'Chq./Ref.No.' = $row.'Cheque Number'
                'Value Dt' = $row.'Transaction Date'
                'Amt (Cr)' = $row.'Deposit Amount (INR )'
            }

            # Handle the case where 'Transaction Remarks' contain 'RESET'
            if ($row.'Transaction Remarks' -match 'RESET') {
                $newRow.Item = "RESET"
                $newRow.Category = "RESET"
                $newRow.Place = "RESET"
                $newRow.Freq = "RESET"
                $newRow.For = "RESET"
            }

            # Add the processed row to the new data set
            $newData += $newRow
            $previousRow = $newRow  # Set this row as the previous row for concatenation purposes
        }
    }

    # Export the new data to the transformed Excel file
    $newData | Export-Excel -Path $transformedFilePath -WorksheetName "FormattedData" -AutoSize
    Write-Host "Transformation complete: $inputFilePath -> $transformedFilePath"
}

# Main script execution
try {
    # Process all XLS files in the current directory
    $xlsFiles = Get-ChildItem -Path . -Filter ICICI_DC_*.xls
    foreach ($xlsFile in $xlsFiles) {
        $xlsFilePath = $xlsFile.FullName
        $baseFileName = [System.IO.Path]::GetFileNameWithoutExtension($xlsFilePath)
        $convertedFilePath = "$($xlsFile.DirectoryName)\$baseFileName`_ConvertedFromXls.xlsx"
        $transformedFilePath = "$($xlsFile.DirectoryName)\$baseFileName`_ConvertedFromXls_Transformed.xlsx"

        try {
            # Convert .xls to .xlsx
            $xlsxFilePath = Convert-XlsToXlsx -inputFilePath $xlsFilePath

            # Process the newly created XLSX file
            $filenameParts = $baseFileName -split '_'
            $bank = $filenameParts[0]
            $accountType = $filenameParts[1]
            $name = $filenameParts[2]
            $MOP = "$bank`_$accountType`_$name"

            $headerInfo = Detect-Header -excelFilePath $xlsxFilePath
            if ($headerInfo -ne $null) {
                Transform-Data -inputFilePath $xlsxFilePath -startRow $headerInfo.StartRow -startColumn $headerInfo.StartColumn -MOP $MOP -transformedFilePath $transformedFilePath
            } else {
                Write-Host "Skipping transformation for file '$xlsxFilePath' due to header detection issues."
            }
        } catch {
            Write-Host "An error occurred while processing file '$xlsFilePath': $_"
        }
    }
} catch {
    Write-Error "An error occurred during script execution: $_"
}
