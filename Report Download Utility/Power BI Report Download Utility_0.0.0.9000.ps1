Param(
    $OutputDirectory=[Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop),
    [Alias("ReportType")]
    [ValidateSet("PDF", "PPTX")]
    $OutputType = "PDF",
    [Alias("GroupName")] 
    $WorkspaceName, 
    $ReportName,
    [Alias("Pages")]
    $PagesToInclude, 
    $CredentialPath
    )


##############################################################
# Power BI Report Download Utility
# Classification: Public Safe
# Version:        0.0.0.9000
# Author:         jay.sumners@gmail.com
# Last Update :   2020-04-13
##############################################################

############################## Installation Preprequisites ##############################

#This script will use Power BI Cmdlets
## Power BI API
#Install-Module -Name MicrosoftPowerBIMgmt
#Use -Force to override if you are updating

if (Get-Module -Name MicrosoftPowerBIMgmt -ListAvailable) {
    [Console]::WriteLine("MicrosoftPowerBIMgmt Module Already Installed...")
} 
else {
    [Console]::WriteLine('Installing SqlServer Module...')
    Install-Module -Name MicrosoftPowerBIMgmt -Force
}

################################# Configuration Settings ################################
Add-Type -AssemblyName System.Windows.Forms

####################################### Functions #######################################
function Get-FileUrl {
    Param($group, $report, $export, [int]$waitIterations=60)

    $i=0
    do{
        $status = Invoke-PowerBIRestMethod -Method Get -Url "https://api.powerbi.com/v1.0/myorg/groups/$group/reports/$report/exports/$export" | ConvertFrom-Json
        ++$i
        [Console]::Write("..$i")
        Start-Sleep -Seconds 1
    } while ((@("Failed", "Succeeded") -notcontains $status.status) -and ($i -lt $waitIterations))

    if($status.status -eq "Failed"){
        [Console]::WriteLine("..Request Failed.")
        return $null
    } else {
        [Console]::WriteLine("..Request Succeeded.")
        return $status.resourceLocation
    }
}

function Get-Settings {
    Param($groupName, $reportName, $pagesToInclude)

    #Local Functions
    function Get-Group {
        Param($groupName)

        [Console]::WriteLine("Getting group...")
        $groups = (Invoke-PowerBIRestMethod -Method Get -Url "https://api.powerbi.com/v1.0/myorg/groups" | ConvertFrom-Json).value

        if($groupName){
            $groupId = ($groups.Where({$_.name -eq $groupName})).Id
            $message = "Workspace not found."
        } else {
            $groupId = ($groups | Select-Object -Property @{n="Workspace Name";e={$_.name}}, Id |Sort-Object -Property 'Workspace Name' | Out-GridView -Title 'Select a workspace:' -OutputMode Single).Id
            $message = "A Workspace is required."
        }

        if($null -eq $groupId){
            $result = [System.Windows.Forms.MessageBox]::Show($message, "Selection Error",[System.Windows.Forms.MessageBoxButtons]::RetryCancel, [System.Windows.Forms.MessageBoxIcon]::Error)
            if($result -eq "Retry"){
                ParamsNull
            } else {
                Disconnect-PowerBIServiceAccount
                EXIT
            }
        }

        return $groupId
    }

    function Get-Report {
        Param($reportName, [Parameter(Mandatory)]$groupId)

        [Console]::WriteLine("Getting report...")
        $reports = (Invoke-PowerBIRestMethod -Method Get -Url "https://api.powerbi.com/v1.0/myorg/groups/$groupId/reports" | ConvertFrom-Json).value

        if($reportName){
            $reportId = ($reports.Where({$_.name -eq $reportName})).Id
            $message = "Report not found."
        } else {
            $reportData = $reports | Select-Object -Property @{n="Report Name";e={$_.name}}, Id | Sort-Object -Property 'Report Name' | Out-GridView -Title 'Select a report:' -OutputMode Single
            $reportName = $reportData.'Report Name'
            $reportId = $reportData.Id
            $message = "A Report is required."
        }

        if($null -eq $reportId){
            $result = [System.Windows.Forms.MessageBox]::Show($message, "Selection Error",[System.Windows.Forms.MessageBoxButtons]::RetryCancel, [System.Windows.Forms.MessageBoxIcon]::Error)
            if($result -eq "Retry"){
                ParamsNull
            } else {
                Disconnect-PowerBIServiceAccount
                EXIT
            }
        }

        return @{
            name = $reportName
            id = $reportId
        }
    }

    function Get-Pages {
        Param($pagesToInclude, $groupId, $reportId)

        [Console]::WriteLine("Getting pages...")
        $pages = ((Invoke-PowerBIRestMethod -Method Get -Url "https://api.powerbi.com/v1.0/myorg/groups/$groupId/reports/$reportId/pages") | ConvertFrom-Json).value

        if($pagesToInclude){
            $pageNames = ($pages.Where({$pagesToInclude -contains $_.displayName})).Name
        } 
        elseif($groupName -and $reportName){
            $pageNames = $null
        }
        else {
            $pageNames = ($pages | Sort-Object -Property order | Select-Object -Property @{n="Page Name";e={$_.displayName}}, @{n="Id";e={$_.Name}} | Out-GridView -Title 'Select included pages (multi-select or Cancel for all):' -OutputMode Multiple).Id
        }

        return $pageNames
    }

    function Get-CallBody {
        
    }

    function ParamsNull {
        $groupId = Get-Group
        $reportHash = Get-Report -groupId $groupId
        $pageNames = Get-Pages -groupId $groupId -reportId $reportHash["id"]

        return @{
            groupId = $groupId
            reportName = $reportHash["name"]
            reportId = $reportHash["id"]
            pageNames = $pageNames
        }
    }

    #Run Settings selection loop
    if(($null -eq $groupName) -or ($null -eq $reportName)){
        $result = ParamsNull
    } 
    else {
        $groupId = Get-Group -groupName $groupName
        $reportHash = Get-Report -groupId $groupId -reportName $reportName
        $pageNames = Get-Pages -pagesToInclude $pagesToInclude -groupId $groupId -reportId $reportHash["id"]

        $result = @{
            groupId = $groupId
            reportName = $reportHash["name"]
            reportId = $reportHash["id"]
            pageNames = $pageNames
        }
    }

    return $result
}

###################################### Main Program ######################################

######## LOGIN / PARTIAL PARAMETER INSURANCE ########
[Console]::WriteLine("Getting credentials and missing parameters...")
# Check for missing parameters
$requiresInteractive = (@("", $null) -contains $WorkspaceName) -or (@("", $null) -contains $ReportName) -or (@("", $null) -contains $CredentialPath) 

# Stored Credentials Function.
function Get-StoredCredentials {
    Param($Path)
    if ( Test-Path $Path ) {
        #crendetial is stored, load it 
        $localCred = Import-CliXml -Path $Path
    } else {
        # no stored credential: create store, get credential and save it
        $localParent = Split-Path $Path -parent
        if ( -not (Test-Path $localParent)) {
            New-Item -ItemType Directory -Force -Path $localParent
        }
        $localCred = Get-Credential
        $localCred | Export-CliXml -Path $Path
    }

    return $localCred
}

# Setup a variable to return
Set-Variable -Name cred -Option AllScope -Force
Set-Variable -Name outputDir -Value $OutputDirectory -Option AllScope -Force
Set-Variable -Name outputExt -Value $OutputType -Option AllScope -Force

if($requiresInteractive){
    # Set assemblies etc.
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")  
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void][System.Windows.Forms.Application]::EnableVisualStyles() 

    # Get Monitor Size and conversion function
    $monitor = [System.Windows.Forms.Screen]::PrimaryScreen
    $widthFactor  = 1366 / $monitor.WorkingArea.Width
    $heightFactor = 728  / $monitor.WorkingArea.Height

    function Get-X {
        Param($x)

        return [Math]::Round($x * $widthFactor)
    }

    function Get-Y {
        Param($y)

        return [Math]::Round($y * $heightFactor)
    }

    # General Functions
    function Get-FileBrowser {
        Param($filter='Power BI (*.pbix)|*.pbix')
        $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
            Title = 'Please select a file'
            InitialDirectory = [Environment]::GetFolderPath('Desktop') 
            Filter = $filter
        }

        [void]$FileBrowser.ShowDialog()
        return $FileBrowser.FileName
    }

    function Get-FolderBrowser {
        $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
            Description = 'Select an output directory...'
        }

        [void]$FolderBrowser.ShowDialog()
        return $FolderBrowser.SelectedPath
    }

    function PopulateTextBox {
        Param($object, $text)

        $object.text = $text
    }

    # Create form object
    $credForm = New-Object System.Windows.Forms.Form -Property @{
        Width = (Get-X 400)
        Height = (Get-Y 400)
        AutoSize = $true
        MaximizeBox = $false
        StartPosition = "CenterScreen"
        FormBorderStyle = "Fixed3D"
        ShowIcon = $false
        Text = "Power BI Report Download Utility: Settings"
    }

    ##### Set standard form elements #####
    # Fonts
    # These can be done in 2 ways, depending on options. If not options, a string is OK.
    #$label_font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Bold) 
    $font_title = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Bold) 
    $font_label = "Arial, 10"
    $font_box = "Arial, 10"
    $font_button = "Arial, 10"
    $font_button_admin = "Arial, 8"

    # Spacing
    $marginX = Get-X 5
    $marginY = Get-Y 5
    $h_space = Get-X 5
    $v_space_close = Get-Y 5
    $v_space_wide = Get-Y 10
    $centerX = ($credForm.Width / 2)

    $headerSize = Get-Y 70
    $centerY = (($credForm.Height - $headerSize) / 2)

    # Standard widths
    $width_cred_button = 230
    $adj_width_cred_button = Get-X $width_cred_button
    $adj_width_cred_button_half = Get-X ($width_cred_button/2)
    $adj_width_admin_button = Get-X 50
    $adj_width_continue_button = Get-X 85

    ##### General Variables #####
    $availableTypes = [String[]]@("Portable Document Format (PDF)", "Microsoft PowerPoint (PPTX)" )
    $availableTypeParse = @{
        "Portable Document Format (PDF)" = "PDF"
        "Microsoft PowerPoint (PPTX)" = "PPTX"
    }

    ##### Build form #####
    # Add Objects
    $label_title = New-Object System.Windows.Forms.Label -Property @{
        Font = $font_title
        Text = "Sign in to the Power BI Service"
        AutoSize = $false
        TextAlign="MiddleCenter"
        Width = (Get-X 250)
        Height = (Get-Y 22)
        Location = New-Object System.Drawing.Point(($centerX - (Get-X (250/2))), $marginY)
    }

    # First
    $button_usr_once = New-Object System.Windows.Forms.Button -Property @{
        Font = $font_button
        Text = "Single Sign-In"
        AutoSize = $false
        Width = $adj_width_cred_button
        Height = (Get-Y 27)
        Location = New-Object System.Drawing.Point (($centerX - ($adj_width_cred_button_half)), ($label_title.Location.Y + $label_title.Height + $v_space_wide))
    }

    $button_usr_new = New-Object System.Windows.Forms.Button -Property @{
        Font = $font_button
        Text = "New Credentials"
        AutoSize = $false
        Width = $adj_width_cred_button
        Height = (Get-Y 27)
        Location = New-Object System.Drawing.Point (($centerX - ($adj_width_cred_button_half)), ($button_usr_once.Location.Y + $button_usr_once.Height + $v_space_close))
    }

    $button_usr_saved = New-Object System.Windows.Forms.Button -Property @{
        Font = $font_button
        Text = "Saved Credentials"
        AutoSize = $false
        Width = $adj_width_cred_button
        Height = (Get-Y 27)
        Location = New-Object System.Drawing.Point (($centerX - ($adj_width_cred_button_half)), ($button_usr_new.Location.Y + $button_usr_new.Height + $v_space_close))
    }

    $divider_main = New-Object System.Windows.Forms.Label -Property @{
        Font = $font_label
        Text = $null
        BorderStyle = "Fixed3D"
        AutoSize = $false
        Width = (Get-X 300)
        Height = (Get-Y 2)
        Location = New-Object System.Drawing.Point( ($centerX - (Get-X 150)), ($button_usr_saved.Location.Y + $button_usr_saved.Height + $v_space_wide + $v_space_close) )
        
    }

    # Second
    $label_reporType = New-Object System.Windows.Forms.Label -Property @{
        Font = $font_label
        Text = "Select the type of output document:"
        AutoSize = $false
        TextAlign="MiddleCenter"
        Width = (Get-X 300)
        Height = (Get-Y 22)
        Location = New-Object System.Drawing.Point( ($centerX - (Get-X 150)), ($divider_main.Location.Y + $divider_main.Height + $v_space_wide) )
        Enabled = $false
    }

    $comboBox_reportType = New-Object System.Windows.Forms.ComboBox -Property @{
        Font = $font_button
        AutoSize = $false
        Width = $adj_width_cred_button
        Height = (Get-Y 27)
        DropDownStyle = "DropDownList"
        DropDownWidth = $adj_width_cred_button
        DropDownHeight = (Get-Y (27*$availableTypes.Count))
        Location = New-Object System.Drawing.Point( ($centerX - ($adj_width_cred_button_half)), ($label_reporType.Location.Y + $label_reporType.Height + $v_space_close) )
        Enabled = $false
    }

    $comboBox_reportType.Items.AddRange($availableTypes)
    $comboBox_reportType.SelectedIndex = 0

    $label_outputDirectory = New-Object System.Windows.Forms.Label -Property @{
        Font = $font_label
        Text = "Select an output folder for the document:"
        AutoSize = $false
        TextAlign="MiddleCenter"
        Width = (Get-X 300)
        Height = (Get-Y 22)
        Location = New-Object System.Drawing.Point( ($centerX - (Get-X 150)), ($comboBox_reportType.Location.Y + $comboBox_reportType.Height + $v_space_wide) )
        Enabled = $false
    }

    $textbox_outputPath = New-Object System.Windows.Forms.TextBox -Property @{
        Font = $font_button
        Text = $OutputDirectory
        AutoSize = $false
        Width = (Get-X ($width_cred_button - 75)) - $h_space
        Height = (Get-Y 22)
        Location = New-Object System.Drawing.Point( ($centerX - ($adj_width_cred_button_half)), ($label_outputDirectory.Location.Y + $label_outputDirectory.Height + $v_space_close) )
        Enabled = $false
    }

    $button_folderBrowse = New-Object System.Windows.Forms.Button -Property @{
        Font = $font_button
        Text = "Browse"
        AutoSize = $false
        Width = (Get-X 75)
        Height = (Get-Y 22)
        Location = New-Object System.Drawing.Point ( ($textbox_outputPath.Location.X + $textbox_outputPath.Width + $h_space), ($textbox_outputPath.Location.Y) )
        Enabled = $false
    }

    $button_continue = New-Object System.Windows.Forms.Button -Property @{
        Font = $font_button
        Text = "Continue>>"
        AutoSize = $false
        Width = $adj_width_continue_button
        Height = (Get-Y 22)
        Location = New-Object System.Drawing.Point(($credForm.Width - $marginX - $adj_width_continue_button - $h_space), ($credForm.Height - $marginX - (Get-Y 22) - $v_space_close) )
        Enabled = $false
    }

    # Helper Functions
    function Switch_Enable_First {
        Param($passingButton)

        $t = @($button_usr_once, $button_usr_new, $button_usr_saved).Where({$_ -ne $passingButton})
        $t.ForEach({$_.Enabled = $true})
    }

    function Enable_Second {
        $label_reporType.Enabled = $true
        $comboBox_reportType.Enabled = $true
        $label_outputDirectory.Enabled = $true
        $textbox_outputPath.Enabled = $true
        $button_folderBrowse.Enabled = $true
        $button_continue.Enabled = $true
    }

    # Event Functions
    function Click_User_Once {
        # Ensure they do not want to save
        $continue = [System.Windows.Forms.MessageBox]::Show(("Credentials will not be saved." + [Environment]::NewLine + "To save credentials locally, select 'New Credentials'."), "", [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
  
        if($continue -eq 'OK'){
            # Disable the button
            $button_usr_once.Enabled = $false

            # Main Operations
            $cred = "once"

            # Adjust form
            $null = Switch_Enable_First -passingButton $button_usr_once
            $null = Enable_Second
        } else {
            #user hit cancel
            #bring them back to the main screen
        }
    }

    function Click_User_New {
        $prompt = New-Object System.Windows.Forms.SaveFileDialog -Property @{
            Filter = 'XML (*.xml)|*.xml'
        }

        $prompt.ShowDialog()

        if($prompt.FileName){
            #user gave a valid path
            # Disable the button
            $button_usr_new.Enabled = $false

            # Main Operations
            $cred = Get-StoredCredentials -Path $prompt.FileName

            # Adjust the form
            $null = Switch_Enable_First -passingButton $button_usr_new
            $null = Enable_Second

        } else {
            #user hit cancel
            #bring them back to the main screen
        }
    }

    function Click_User_Saved {
        $credPath = Get-FileBrowser -filter 'XML (*.xml)|*.xml'
        if($credPath){
            #user gave a valid path
            # Disable the button
            $button_usr_saved.Enabled = $false

            # Main Operations
            $cred = Get-StoredCredentials -Path $credPath

            # Adjust the form
            $null = Switch_Enable_First -passingButton $button_usr_new
            $null = Enable_Second

        } else {
            #user hit cancel. 
            #send back to main screen.
        }
    }

    function Click_Continue {
        $errorMessages = @{
            noOutput = 'An output directory is required.'
            invalidOutput = 'A valid output directory is required.'
        }

        $essentials = @(
            [PSCustomObject]@{name='noOtput'; value=$textbox_outputPath.Text} 
            [PSCustomObject]@{name='invalidOutput'; value=(Test-Path -LiteralPath $textbox_outputPath.Text -PathType Container -ErrorAction SilentlyContinue)}
        )

        $err = ($essentials.Where({@("", $null, $false, 0) -contains $_.value})).ForEach({$errorMessages[$_.name]})

        if($err){
            #If there is an error
            [System.Windows.Forms.MessageBox]::Show(($err -join [Environment]::NewLine))
        } 
        else {
            #If we are good-to-go
            $outputDir = $textbox_outputPath.Text
            $outputExt = $availableTypeParse[$comboBox_reportType.SelectedItem]
            $credForm.Close()
            $credForm.Dispose()
        }
    }

    # Add Events
    $button_usr_once.Add_Click({Click_User_Once})
    $button_usr_new.Add_Click({Click_User_New})
    $button_usr_saved.Add_Click({Click_User_Saved})
    $button_folderBrowse.Add_Click({PopulateTextBox -object $textbox_outputPath -text (Get-FolderBrowser)})
    $button_continue.Add_Click({Click_Continue})

    ##### Centering & Adding #####
    # Centering and adding
    $shift = [Math]::Round($centerY - $divider_main.Location.Y, 0)
    $corrected_controls = 
    @(
        $label_title, 
        $button_usr_once, 
        $button_usr_new, 
        $button_usr_saved, 
        $divider_main, 
        $label_reporType, 
        $comboBox_reportType, 
        $label_outputDirectory, 
        $textbox_outputPath, 
        $button_folderBrowse
    ).ForEach({
        $_.Location = New-Object System.Drawing.Point($_.Location.X, ($_.Location.Y + $shift))
        $credForm.Controls.Add($_)
    })

    #Add Controls out of tab order
    $credForm.Controls.Add($button_continue)

    ##### Show Form #####
    [void]$credForm.ShowDialog()
} 
else {
    $cred = Get-StoredCredentials -Path $CredentialPath
}

# Login to Service
try{
    if($cred -ieq "once"){
        $null = Connect-PowerBIServiceAccount
    } elseif ($cred) {
        $null = Connect-PowerBIServiceAccount -Credential $cred
    } else {
        EXIT
    }
}
catch
{
    $null = [System.Windows.Forms.MessageBox]::Show(($_.Exception.Message.ToString()), "Login Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    $credForm.Close()
    $credForm.Dispose()
    EXIT
}

######## Get Ids ########
# Deal with main parameters
$settings = Get-Settings -groupName $WorkspaceName -reportName $ReportName -pagesToInclude $PagesToInclude

$groupId = $settings["groupId"]
$report = $settings["reportName"]
$reportId = $settings["reportId"]
$pageNames = $settings["pageNames"]


######## Set Outputs ########
# Set outputFile
[Console]::WriteLine("Setting output file name...")
if(@("", $null) -contains $OutputDirectory){
    [Console]::WriteLine("  Setting default directory (Desktop)...")
    $outputDir = [Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
}

# Select the correct extension
$ext = "." + $outputExt.ToLower()

# outFile with date appended
#$outFile = $outputDir + "\" + $report + "_" + "$(Get-Date -Format "yyyy-MM-dd")" + $ext

# outFile with PREVIEW appended
$outFile = $outputDir + "\" + $report + "_PREVIEW" + $ext

######## Generate & Export PDF ########
[Console]::WriteLine("Generating body of call...")
if($pageNames){
    $joinedPages = $pageNames.ForEach({'{"pageName":"' + $_ + '"}'}) -join ","
    $body = '{"format":"' + $outputExt + '", "powerBIReportConfiguration":{"pages":[' + $joinedPages + ']}}'
}
else {
    $body = '{"format":"' + $outputExt + '"}'
}


[Console]::WriteLine("Calling Service to start report...")
$exportId = ((Invoke-PowerBIRestMethod -Method Post -Url "https://api.powerbi.com/v1.0/myorg/groups/$groupId/reports/$reportId/ExportTo" -Body $body) | ConvertFrom-Json).Id

# Check on export file (will not export if already exists)
[Console]::WriteLine("Testing output directory...")
if(Test-Path -LiteralPath $outFile) {
    Remove-Item -LiteralPath $outFile -Force
}

# Get PDF and export
[Console]::WriteLine("Waiting on report from service...")
$getUrl = Get-FileUrl -group $groupId -report $reportId -export $exportId -waitIterations 180

if($getUrl){
    [Console]::WriteLine("Getting and saving PDF..")
    $null = Invoke-PowerBIRestMethod -Method Get -Url $getUrl -OutFile $outFile
}
else {
    [Console]::WriteLine("Failed to get document.")
}

######## DISCONNECT ########
[Console]::WriteLine("Disconnecting from Service...")
Disconnect-PowerBIServiceAccount

[Console]::WriteLine("Goodbye!")

