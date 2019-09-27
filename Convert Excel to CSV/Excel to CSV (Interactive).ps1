Write-Output "Welcome!"

#Set necessary assemblies
Add-Type -AssemblyName System.Windows.Forms

#Show the file browser
Write-Output "Please select an Excel file"
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
    Filter = 'SpreadSheet (*.xlsx)|*.xlsx'
}

$null = $FileBrowser.ShowDialog()

#Set directory-level values
$excelFile = $FileBrowser.FileName
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
Write-Output "Please select a sheet"
$availableSheets = foreach ($ws in $wb.Worksheets)
    {
        $ws.Name
    }

$sheetName = $availableSheets | Out-GridView -Title 'Select a sheet to convert' -OutputMode Single

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