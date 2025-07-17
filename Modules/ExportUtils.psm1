function Format-InboxRuleXlsx {
    param(
        [Parameter(Mandatory=$true)]
        [string]$CsvPath,
        [Parameter(Mandatory=$true)]
        [string]$XlsxPath,
        [Parameter(Mandatory=$false)]
        [int]$RowHighlightColor = 6, # Excel ColorIndex for Yellow
        [Parameter(Mandatory=$false)]
        [string]$RowHighlightColumnHeader = "IsHidden", 
        [Parameter(Mandatory=$false)]
        [object]$RowHighlightValue = $true,
        [Parameter(Mandatory=$false)]
        [int]$CellHighlightColor = 38, # Excel ColorIndex for Light Red (Rose)
        [Parameter(Mandatory=$false)]
        [object]$CellHighlightValue = $true
    )

    $excel = $null; $workbook = $null; $worksheet = $null; $usedRange = $null; $columns = $null; $rows = $null; $headerRange = $null; $targetColumnIndex = $null
    $xlOpenXMLWorkbook = 51
    $xlExpression = 2 
    $xlCellValue = 1  
    $xlEqual = 3      
    $missing = [System.Reflection.Missing]::Value 

    try { $excel = New-Object -ComObject Excel.Application -ErrorAction Stop } 
    catch { 
        $ex = $_.Exception
        Write-Error ("Excel COM object creation failed: {0}" -f $ex.Message)
        return $false 
    }

    try {
        $excel.Visible = $false; $excel.DisplayAlerts = $false    
        Write-Host "Converting '$CsvPath' to '$XlsxPath'..."
        $workbook = $excel.Workbooks.Open($CsvPath); $workbook.SaveAs($XlsxPath, $xlOpenXMLWorkbook); $workbook.Close($false) 
        Write-Host "Initial conversion successful. Formatting..."
        $workbook = $excel.Workbooks.Open($XlsxPath); $worksheet = $workbook.Worksheets.Item(1); $usedRange = $worksheet.UsedRange; $columns = $usedRange.Columns; $rows = $usedRange.Rows

        if ($usedRange.Rows.Count -gt 0) {
            Write-Host " - Autofitting columns..."; $columns.AutoFit() | Out-Null
            Write-Host " - Autofitting rows..."; $rows.AutoFit() | Out-Null
            Write-Host " - Bolding header row..."; $headerRange = $worksheet.Rows.Item(1); $headerRange.Font.Bold = $true

            if ($usedRange.Rows.Count -gt 1) {
                $dataRange = $usedRange.Offset(1,0).Resize($usedRange.Rows.Count -1) 
                Write-Host " - Clearing existing conditional formats from data range..."
                $dataRange.FormatConditions.Delete() | Out-Null

                Write-Host "   - Applying Rule 1: Light Red for any TRUE cell..."
                $formatCondition1 = $dataRange.FormatConditions.Add($xlCellValue, $xlEqual, "TRUE") 
                $formatCondition1.Interior.ColorIndex = $CellHighlightColor
                
                Write-Host "   - Applying Rule 2: Manually searching for '$RowHighlightColumnHeader' column for row highlighting..."
                for ($colIdx = 1; $colIdx -le $headerRange.Columns.Count; $colIdx++) {
                    $cell = $null
                    try {
                        $cell = $headerRange.Cells.Item(1, $colIdx)
                        if ($cell.Value2 -is [string] -and $cell.Value2.Equals($RowHighlightColumnHeader, [System.StringComparison]::OrdinalIgnoreCase)) {
                            $targetColumnIndex = $colIdx
                            Write-Host "     - Found '$RowHighlightColumnHeader' at column index $targetColumnIndex."
                            break
                        }
                    } finally {
                        if ($cell) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($cell) | Out-Null }
                    }
                }

                if ($targetColumnIndex) {
                    Write-Host "     - Highlighting rows where '$RowHighlightColumnHeader' is '$($RowHighlightValue)'..."
                    $columnLetter = $worksheet.Columns.Item($targetColumnIndex).Address($false, $false) -replace '\d','' 
                    $formulaForRowHighlight = "`$${columnLetter}$($dataRange.Row)=$($RowHighlightValue.ToString().ToUpper())" 
                    Write-Host "     - Using formula for row highlight: $formulaForRowHighlight"
                    $formatCondition2 = $dataRange.FormatConditions.Add($xlExpression, $missing, $formulaForRowHighlight) 
                    $formatCondition2.Interior.ColorIndex = $RowHighlightColor
                    Write-Host "     - Row highlighting rule for '$RowHighlightColumnHeader' applied."
                } else { Write-Warning "   - '$RowHighlightColumnHeader' column not found. Skipping row highlighting." }
            } else { Write-Host " - Only header row found, skipping conditional formatting." }
        } else { Write-Host " - Worksheet appears empty, skipping formatting." }
        
        Write-Host "Saving formatted XLSX file..."; $workbook.Save(); $workbook.Close()
        Write-Host "XLSX formatting complete."
    } catch {
        $ex = $_.Exception
        Write-Error ("Excel formatting/conversion error: {0}`n{1}" -f $ex.Message, $ex.ScriptStackTrace)
        try { if ($workbook -ne $null) { $workbook.Close($false) } } catch {}
        return $false 
    } finally {
        if ($formatCondition1) {[System.Runtime.InteropServices.Marshal]::ReleaseComObject($formatCondition1) | Out-Null}
        if ($formatCondition2) {[System.Runtime.InteropServices.Marshal]::ReleaseComObject($formatCondition2) | Out-Null}
        if ($dataRange) {[System.Runtime.InteropServices.Marshal]::ReleaseComObject($dataRange) | Out-Null}
        if ($headerRange) {[System.Runtime.InteropServices.Marshal]::ReleaseComObject($headerRange) | Out-Null}
        if ($columns) {[System.Runtime.InteropServices.Marshal]::ReleaseComObject($columns) | Out-Null}
        if ($rows) {[System.Runtime.InteropServices.Marshal]::ReleaseComObject($rows) | Out-Null}
        if ($usedRange) {[System.Runtime.InteropServices.Marshal]::ReleaseComObject($usedRange) | Out-Null}
        if ($worksheet) {[System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null}
        if ($workbook) {[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null}
        if ($excel) {$excel.Quit();[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null}
        [gc]::Collect(); [gc]::WaitForPendingFinalizers(); Write-Host "COM cleanup finished."
    }
    return $true 
}

Export-ModuleMember -Function Format-InboxRuleXlsx 