# Set-StrictMode -Version 3.0 # For development purposes only

if (-not (Get-command -module psexcel)) {
    Install-module PSExcel -Scope CurrentUser
    Get-command -module psexcel
    import-module psexcel
}

$host.PrivateData.ProgressBackgroundColor = $host.UI.RawUI.BackgroundColor
$host.privatedata.ProgressForegroundColor = "green";

$currentFolder = (Get-Location | Select-Object -ExpandProperty Path)

$file = $args[0]
$excel = new-object -com Excel.Application -Property @{ Visible = $false }
$excel.DisplayAlerts = $false

try {
    # Test-Path $file
    $workbook = $excel.Workbooks.Open($file)
}
catch {
    Write-Output "Invalid path: '$file'"
    exit
}

$myCustomObject = New-Object -TypeName psobject

function getLastRowWithValue {
    param (
        [int]$startColumn,
        [int]$endColumn,
        [int]$tableHeadRow,
        [int]$touchedRowsEnd
    )

    # Get last row with value by going from bottom to top
    for ($row = $touchedRowsEnd; $row -ge $tableHeadRow + 1; $row--) {
        for ($column = $startColumn; $column -le $endColumn; $column++) {
            $value = $sheet.Cells.Item($row, $column).Text
            if (([string]$value).Replace(" ", "") -ne "") {
                return $row
            }
        }
    }
}

function Stopwatch {
    param ( [ScriptBlock]$ScriptBlock )
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    & $ScriptBlock
    $stopwatch.Stop()
    Set-Variable Elapsed -value $stopwatch.Elapsed -scope 1
}

function detectStart {
    param (
        $sheet
    )
    $touchedRowsCount = $sheet.UsedRange.Rows.Count
    $touchedColumnsCount = $sheet.UsedRange.Columns.Count

    $touchedRowsStart = $sheet.UsedRange.Row
    $touchedColumnsStart = $sheet.UsedRange.Column
    $touchedRowsEnd = $touchedRowsStart + $touchedRowsCount - 1
    $touchedColumnsEnd = $touchedColumnsStart + $touchedColumnsCount - 1

    Write-Output "tesst"
    Write-Output "Used Range: $touchedColumnsStart,$touchedRowsStart - $touchedColumnsEnd,$touchedRowsEnd"
    Write-Output "tesst"

    $presumableTableHead = @()
    $potentialTableHead = @()

    # Find table head and and first respectively last column
    for ($row = $touchedRowsStart; $row -le $touchedRowsEnd; $row++) {
        for ($column = $touchedColumnsStart; $column -le $touchedColumnsEnd; $column++) {
            $value = $sheet.Cells.Item($row, $column).Text
            Write-Output $value
            if (([string]$value).Replace(" ", "") -ne "") {
                $potentialTableHead += @{column = $column; row = $row }
            }
            else {
                if ($potentialTableHead.length -gt $presumableTableHead.length) {
                    $presumableTableHead = $potentialTableHead
                }
                $potentialTableHead = @()
            }
        }
    }
    break

    # [PSCustomObject] -> Otherwise it won't add anything
    $boundaries = [PSCustomObject]@{
        tableHeadRow = $presumableTableHead[0].row;
        startColumn  = $presumableTableHead[0].column;
        endColumn    = $presumableTableHead[-1].column
    }

    $lastRowWithValue = getLastRowWithValue `
        -startColumn $boundaries.startColumn `
        -endColumn $boundaries.endColumn `
        -tableHeadRow $boundaries.tableHeadRow `
        -touchedRowsEnd $touchedRowsEnd

    $boundaries | Add-Member -MemberType NoteProperty -Name lastRow -Value $lastRowWithValue

    return $boundaries
}

function detectStartCSV {
    param (
        $sheet
    )
    $Path = (Get-Location | Select-Object -ExpandProperty Path) + "\csvs\*.csv"
    (Get-Content -Path $Path) | Set-Content -Path $Path -Encoding UTF8 

    $csv = Import-Csv  ((Get-Location | Select-Object -ExpandProperty Path) + ".\csvs\Tabelle1.csv") 

    $presumableTableHead = @()
    $potentialTableHead = @()

    for ($row = 0; $row -lt $csv.length; $row++) {
        $columns = $csv[$row].PsObject.Properties.Value
        for ($column = 0; $column -lt $columns.length; $column++) {
            $value = $columns[$column]
            if (([string]$value).Replace(" ", "") -ne "") {
                $potentialTableHead += @{column = $column; row = $row }
            }
            else {
                if ($potentialTableHead.length -gt $presumableTableHead.length) {
                    $presumableTableHead = $potentialTableHead
                }
                $potentialTableHead = @()
            }
        }
    }

    # Write-Output "$presumableTableHead"
      
    $boundaries = [PSCustomObject]@{
        tableHeadRow = $presumableTableHead[0].row;
        startColumn  = $presumableTableHead[0].column;
        endColumn    = $presumableTableHead[-1].column;
        lastRow = ($csv.length - 1)
    }

    # [PSCustomObject] -> Otherwise it won't add anything
    # $lastRowWithValue = getLastRowWithValue `
    #     -startColumn $boundaries.startColumn `
    #     -endColumn $boundaries.endColumn `
    #     -tableHeadRow $boundaries.tableHeadRow `
    #     -touchedRowsEnd $touchedRowsEnd


    # $boundaries | Add-Member -MemberType NoteProperty -Name "lastRow" -Value ($csv.length - 1)

    return $boundaries
}


# Loop though all the available worksheets
foreach ($sheet in $workbook.WorkSheets) {
    Write-Output "Looking for table in worksheet $($sheet.Name) ..."

    # $startTime = (Get-Date)
    $boundaries1 = detectStart $sheet
    break
    # $elapsedTime = (Get-Date) - $startTime

    # Write-Output ("Excel: " + $elapsedTime.PSObject.Properties["TotalSeconds"].Value + " seconds")
    # $startTime = (Get-Date)
    # $boundaries2 = detectStartCSV $sheet
    # $elapsedTime = (Get-Date) - $startTime
    # Write-Output ("CSV: " + $elapsedTime.PSObject.Properties["TotalSeconds"].Value + " seconds")
    
    Write-Output $boundaries1
    Write-Output $boundaries2

    Write-Output "Table found!"
    Write-Output ""

    $tableHeaderRow = $boundaries.tableHeadRow
    $startRow = $boundaries.tableHeadRow + 1
    $startColumn = $boundaries.startColumn

    $rows = @()

    $rowEnd = $boundaries.lastRow
    $columnEnd = $boundaries.endColumn


    # Which row is reperesents the table head
    $tableHeaderRowPrompt = Read-Host -Prompt "Table header row=$($tableHeaderRow)? [Overwrite]"

    if (-not [string]::IsNullOrWhiteSpace($tableHeaderRowPrompt)) {
        $tableHeaderRow = [int32]$tableHeaderRowPrompt
    }

    # Which row to start from
    $startRowPrompt = Read-Host -Prompt "Start row=$($startRow)? [Overwrite]"

    if (-not [string]::IsNullOrWhiteSpace($startRowPrompt)) {
        $startRow = [int32]$startRowPrompt
    }

    # Which column to start from
    $startColumnPrompt = Read-Host -Prompt "Start column=$($startColumn)? [Overwrite]"

    if (-not [string]::IsNullOrWhiteSpace($startColumnPrompt)) {
        $startColumn = [int32]$startColumnPrompt
    }

    # Set max row
    $rowEndPrompt = Read-Host -Prompt "Last row=$rowEnd [Overwrite]"

    if (-not [string]::IsNullOrWhiteSpace($rowEndPrompt)) {
        $rowEnd = [int32]$rowEndPrompt
    }

    #Set max column
    $columnEndPrompt = Read-Host -Prompt "Last column=$columnEnd [Overwrite]"

    if (-not [string]::IsNullOrWhiteSpace($columnEndPrompt)) {
        $columnEnd = [int32]$columnEndPrompt
    }

    Write-Output "Dimensions: Rows $($startRow)-$($rowEnd), Columns $($startColumn)-$($columnEnd)"

    $rowIterationCount = $rowEnd - $startRow + 1

    # Loop through rows and columns
    for ($row = $startRow; $row -le $rowEnd; $row++) {
        Write-Progress -Activity "Reading rows ..." -Status "Row $($row - $startRow + 1) of $rowIterationCount" -PercentComplete $($row / $rowEnd * 100)

        [PSCustomObject]$columns = New-Object -TypeName psobject

        for ($column = $startColumn; $column -le $columnEnd; $column++) {
            $columnHead = $sheet.Cells.Item($tableHeaderRow, $column).Text.Replace("`n", " ")
            $value = $sheet.Cells.Item($row, $column).Text

            if ($column -eq $startColumn -or -not ($columns.psobject.Properties.name -contains $columnHead)) {
                # Check if key already exists or object is empty
                $columns | Add-Member -MemberType NoteProperty -Name $columnHead -Value $value # Add all values in row to $columns
            }
        }
        $rows += , $columns # Add array of values to $rows
    }
    $myCustomObject | Add-Member -MemberType NoteProperty -Name $sheet.Name -Value $rows # Create new Key for each Worksheet with corresponding rows

    Write-Output "--------"
}

Write-Output $myCustomObject

# Save result to JSON file
[PSCustomObject]$myCustomObject | ConvertTo-Json -Depth 3 -Compress | Set-Content -Path "$($currentFolder)\$($ws.name).json" -Encoding UTF8

Write-Output "Quitting Excel ..."

$excel.Workbooks.Close()
$excel.Quit()