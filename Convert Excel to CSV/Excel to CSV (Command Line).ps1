#Command Line Input --Filepath, --Sheetname

if($args[0] -eq $null -or $args[1] -eq $null){
    Write-Output "Parameters not provided (--Filepath --Sheetname). Exiting."
}
else {
    Write-Output "Welcome!"

    #Set directory-level values
    $excelFile = $args[0]
    $splitFilename = $excelFile.split("\")
    $excelDirectory = ($splitFilename[0..($splitFilename.Count - 2)] -join "\") + "\"
    $excelFilename = $splitFilename[($splitFilename.Count-1)].Replace(".xlsx", "")

    # Load module
    Write-Output "Opening Excel file. This may take some time."
    $excelApp = New-Object -ComObject Excel.Application
    $excelApp.Visible = $false
    $excelApp.DisplayAlerts = $false

    #Load Workbook
    $wb = $excelApp.Workbooks.Open($excelFile)

    #Set the sheet name and some final variables
    $sheetName = $args[1]

    $n = $excelFileName + "_" + $sheetName
    $saveAsFilename = $excelDirectory + $n + ".csv"

    #Get the worksheet we are going to save
    Write-Output "Getting worksheet"
    $ws = $wb.Worksheets | where Name -eq $sheetName

    #Save worksheet as CSV and quit module
    Write-Output "Saving worksheet as CSV"
    $ws.SaveAs($saveAsFilename, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlCSV)

    Write-Output "Cleaning up"

    $excelApp.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null

    Write-Output "Complete!"
}
