Function ExcelToCsv () {
    param (
        $file,
        $csvFolder
    )
    $myDir = (Get-Location | Select-Object -ExpandProperty Path)
    $excel = new-object -com Excel.Application -Property @{ Visible = $false }
    $excel.DisplayAlerts = $false
    $wb = $Excel.Workbooks.Open($file)
    Write-Output $file
    $xlCSV = 6
    try {
        foreach ($ws in $wb.Worksheets) {
            $x = $ws
            $x.SaveAs("$myDir\$csvFolder\$($ws.name).csv", $xlCSV) 
        }
    } 
    catch {
        Write-Output $_
    }
    $Excel.Workbooks.Close()
    $Excel.Quit()
}


function detectStartCSV {
    param (
        $csv
    )
    Write-Output "length" $csv.length
    $presumableTableHead = @()
    $potentialTableHead = @()
        
    for ($row = 0; $row -lt $csv.length; $row++) {
        $columns = $csv[$row].PsObject.Properties.Value
        for ($column = 0; $column -lt $columns.length; $column++) {
            $value = $columns[$column]
            # Write-Output ($columns.GetType() -Eq [string])
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
    
    
    $boundaries = [PSCustomObject]@{
        tableHeadRow = $presumableTableHead[0].row;
        startColumn  = $presumableTableHead[0].column;
        endColumn    = $presumableTableHead[-1].column;
        lastRow      = ($csv.length - 1);
    }
    # Write-Output ($csv.length - 1)
    # Write-Output $boundaries
    
    
    # [PSCustomObject] -> Otherwise it won't add anything
    # $lastRowWithValue = getLastRowWithValue `
    #     -startColumn $boundaries.startColumn `
    #     -endColumn $boundaries.endColumn `
    #     -tableHeadRow $boundaries.tableHeadRow `
    #     -touchedRowsEnd $touchedRowsEnd


    # $boundaries | Add-Member -MemberType NoteProperty -Name "lastRow" -Value ($csv.length - 1)

    return $boundaries
}

$file = $args[0]
$csvFolderName = ".csv-files"
$currentFolder = (Get-Location | Select-Object -ExpandProperty Path)
$excel = new-object -com Excel.Application -Property @{ Visible = $false }

$pathToCSVs = $currentFolder + "\$csvFolderName"

if (-not(Test-Path $pathToCSVs)) {
    mkdir $csvFolderName
}
# Test-Path $file
ExcelToCsv $file $csvFolderName

$myCustomObject = New-Object -TypeName psobject

Get-ChildItem "$currentFolder\$csvFolderName" | ForEach-Object {  
    $content = Get-Content -Path "$pathToCSVs\$($_.Name)"
    $content | Set-Content -Path "$pathToCSVs\$($_.Name)" -Encoding UTF8
}


Get-ChildItem "$currentFolder\$csvFolderName" | ForEach-Object {
    $sheetName = $_.Name
    $pathToSheet = "$pathToCSVs\$sheetName"

    $csv = Import-Csv $pathToSheet
    Write-Output "Columns" $csv.length
    $boundaries = detectStartCSV $csv
    $tableHeaderRow = $boundaries.tableHeadRow
    $startRow = $boundaries.tableHeadRow + 1
    $startColumn = $boundaries.startColumn
    $rowEnd = $boundaries.lastRow
    $columnEnd = $boundaries.endColumn
    
    $rows = @()

    # if($sheetName -ne "Allgemeines.csv") {
    #     break
    # }
        
    # if ($tableHeaderRow -eq $null) {
    #     break
    # }
    for ($row = $startRow; $row -le $rowEnd; $row++) {
        # Write-Progress -Activity "Reading rows ..." -Status "Row $($row - $startRow + 1) of $rowIterationCount" -PercentComplete $($row / $rowEnd * 100)
        
        [PSCustomObject]$columns = New-Object -TypeName psobject

        for ($column = $startColumn; $column -le $columnEnd; $column++) {
            $columnHead = $csv[$tableHeaderRow].PsObject.Properties.Value[$column].Replace("`n", " ").Replace("`r", " ")
            $value = $csv[$row].PsObject.Properties.Value[$column]

            if ($column -eq $startColumn -or -not ($columns.psobject.Properties.name -contains $columnHead)) {
                # Check if key already exists or object is empty
                $columns | Add-Member -MemberType NoteProperty -Name $columnHead -Value $value # Add all values in row to $columns
            }
        }
        $rows += , $columns # Add array of values to $rows
    }
    $myCustomObject | Add-Member -MemberType NoteProperty -Name $sheetName.Replace(".csv", "") -Value $rows # Create new Key for each Worksheet with corresponding rows

}


[PSCustomObject]$myCustomObject `
| ConvertTo-Json -Depth 3 -Compress `
| Set-Content -Path "$($currentFolder)\$($file.Split("\")[-1].Replace('.xlsx', '')).json" -Encoding UTF8

Remove-Item -R $csvFolderName