
## Microsoft Azure Active Directory Module for Windows PowerShell
#Install-Module MSOnline
#Use -Force to override if you are updating

# Set function to get path later on
Function Get-Folder($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select a folder"
    $foldername.rootfolder = "MyComputer"

    if($foldername.ShowDialog() -eq "OK")
    {
        $folder += $foldername.SelectedPath
    }
    return $folder
}

#NOTE: Do not need to run in admin.
Connect-MsolService

# All users with a pro license (will take some time to run)
$proUsers = Get-MsolUser -All | 
            Where-Object {$_.Licenses.AccountSkuId -eq "discoverycomm:POWER_BI_PRO"} | 
            Select-Object -Property UserPrincipalName, DisplayName, FirstName, LastName, Title, Department, Office, Country, BlockCredential

#No Disconnect for MsolService

## Get JSON data on Power BI Usage
$path = Get-Folder

$proUsers | Export-Csv -Path $path -NoTypeInformation