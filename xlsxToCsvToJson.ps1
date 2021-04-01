if (-not (Get-command -module psexcel)) {
    Install-module PSExcel -Scope CurrentUser
    Get-command -module psexcel
    import-module psexcel
}

$host.PrivateData.ProgressBackgroundColor = $host.UI.RawUI.BackgroundColor
$host.privatedata.ProgressForegroundColor = "green";

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

    # Write-Output "-$($csv[3])-$($csv.length)"
    # break
    for ($row = 0; $row -lt $csv.length; $row++) {
        $columns = $csv[$row].PsObject.Properties.Value
        $isString = ($csv[$row].PsObject.Properties.Value.GetType() -Eq [string])
        # break
        for ($column = 0; $column -lt $columns.length; $column++) {
            
            $value = $columns[$column]
            # Write-Host $columns
            # if($isString) {
            #     $value = $columns
            # }
            # Write-Host "-$($csv[$row])-" 
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
    # break
    
    
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
    $boundaries = detectStartCSV $csv

    $tableHeaderRow = $boundaries.tableHeadRow
    $startRow = $boundaries.tableHeadRow + 1
    $startColumn = $boundaries.startColumn
    $rowEnd = $boundaries.lastRow
    $columnEnd = $boundaries.endColumn
    
    $rows = @()

    try {
        for ($row = $startRow; $row -le $rowEnd; $row++) {
            Write-Progress -Activity "Reading rows ..." -Status "Row $($row - $startRow + 1) of $($csv.length)" -PercentComplete $($row / $rowEnd * 100)
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
    catch {
        Write-Host $_
        break
        Write-Output "Error reading $($sheetName.Replace('.csv', ''))"
    }

}

$jsonFileName = "$($file.Split("\")[-1].Replace('.xlsx', '')).json"
Write-Output "Saving results to $jsonFileName ..."

[PSCustomObject]$myCustomObject `
| ConvertTo-Json -Depth 3 -Compress `
| Set-Content -Path "$($currentFolder)\$jsonFileName" -Encoding UTF8

Write-Output "Done!"
Remove-Item -R $csvFolderName