<#
.SYNOPSIS

Downloads PBIX from Power BI Service as PPTX or PDF.

.DESCRIPTION

Uses the Power BI Rest API to generate a PPTX or PDF of a Power BI report,
on the server and saves it locally.

.PARAMETER pbixFilePath

Specifies the name and path for the PBIX file you are documenting.

Type: String
Optional: Absense will trigger the WinForms GUI.

.PARAMETER outputFolder

Specifies the directory where the output documents will be stored.

Type: String
Optional: Absense will trigger the WinForms GUI.

.PARAMETER outputDocuments

Types of reports to generate. Limited options.

Type: String or Array of Strings
Optional: Absense will trigger the WinForms GUI.
Options: DataDescriptions, VisualContainers, UnusedObjects

.INPUTS

None. You cannot pipe objects to Update-Month.ps1.

.OUTPUTS

No pipleline outputs. The script will only output documents to the save
location indicated in the parameters.
#>

Param(
    $pbixFilePath,
    $outputFolder,
    $outputDocuments
)

##############################################################
# Power BI Unused Objects Retreival
# NOTE: Very similar to Power BI Data Definition Retreival
# Version:        1.0.1.0
# Author:         Jay Sumners
# Last Update :   2021-09-04
##############################################################

############################## Installation Preprequisites ##############################

# Requires remote signed
# This only has to be done once on your machine, not every time the script runs
# Set-ExecutionPolicy RemoteSigned

# Will also need to have the OLAP drivers installed.
# These come with Power BI for Excel or can be downloaded directly.

################################# Parameter Insurance ################################
# Function for Testing Text Parameters
function Test-Param {
    Param([String]$Param)
    return (@($null, "") -contains $Param)
}

# User Interface if parameters not set on CLI
if((Test-Param $pbixFilePath) -or (Test-Param $outputFolder)) {
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

    # Event Functions
    function Get-FileBrowser {
        $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
            Title = 'Please select a Power BI Desktop file (PBIX)'
            InitialDirectory = [Environment]::GetFolderPath('Desktop') 
            Filter = 'Power BI (*.pbix)|*.pbix'
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

    function Get-ListItems {
        Param($list_object)

        $checked_items =
            for($i=0; $i -lt ($list_object.Items).Count; ++$i) {
                if($list_object.GetItemChecked($i)) {
                    $list_object.Items[$i]
                }
            }
        return $checked_items
    }

    function GoNoCheck {
        $errorMessages = @{
            pbix = 'A Power BI Desktop file must be selected.'
            folder = 'An ouput folder must be selected.'
            output_choice = 'At least one (1) output document must be selected.'
        }

        $essentials = @(
            [PSCustomObject]@{name='pbix'; value=$box_pbix.Text} 
            [PSCustomObject]@{name='folder'; value=$box_folder.Text}
            [PSCustomObject]@{name='output_choice'; value=(Get-ListItems -list_object $check_choice).Count}
        )

        $err = ($essentials.Where({@("", $null, 0) -contains $_.value})).ForEach({$errorMessages[$_.name]})

        return $err
    }

    function OKButton {
        $err = GoNoCheck

        if($err){
            #If there is an error
            [System.Windows.Forms.MessageBox]::Show(($err -join [Environment]::NewLine))
        } else {
            #If we are good-to-go
            $form.Close()
        }
    }

    # Create form object
    $form = New-Object System.Windows.Forms.Form -Property @{
        Width = (Get-X 400)
        Height = (Get-Y 400)
        AutoSize = $true
        MaximizeBox = $false
        StartPosition = "CenterScreen"
        FormBorderStyle = "Fixed3D"
        ShowIcon = $false
        Text = "Power BI WinForms Utility"
    }

    ##### Set standard form elements #####
    # Fonts
    # These can be done in 2 ways, depending on options. If not options, a string is OK.
    #$label_font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Bold) 
    $font_label = "Arial, 10"
    $font_box = "Arial, 10"
    $font_button = "Arial, 10"

    # Spacing
    $marginX = Get-X 5
    $marginY = Get-Y 5
    $h_space = Get-X 5
    $v_space_close = Get-Y 5
    $v_space_wide = Get-Y 10

    ##### Build form #####
    ##### Get PBIX file path ######
    # Add Objects
    $label_pbix = New-Object System.Windows.Forms.Label -Property @{
        Font = $font_label
        Text = "Select a Power BI Desktop (.pbix) file"
        AutoSize = $false
        Width = (Get-X 300)
        Height = (Get-Y 15)
        Location = New-Object System.Drawing.Point($marginX,$marginY)
    }

    $box_pbix = New-Object System.Windows.Forms.TextBox -Property @{
        Font = $font_box
        MultiLine = $false
        AutoSize = $false
        Width = (Get-X 295)
        Height = (Get-Y 20)
        Location = New-Object System.Drawing.Point(($label_pbix.Location.X + $h_space),($label_pbix.Location.Y + $label_pbix.Height + $v_space_close))
    }

    $button_file_browse = New-Object System.Windows.Forms.Button -Property @{
        Font = $font_button
        Text = "Browse"
        AutoSize = $false
        Width = (Get-X 75)
        Height = (Get-Y 20)
        Location = New-Object System.Drawing.Point (($box_pbix.Location.X + $box_pbix.Width + $h_space), ($box_pbix.Location.Y))
    }

    # Add Events
    $box_pbix.add_Click({PopulateTextBox -object $box_pbix -text (Get-FileBrowser)})
    $button_file_browse.add_Click({PopulateTextBox -object $box_pbix -text (Get-FileBrowser)})

    # Add to Form
    $form.Controls.Add($label_pbix)
    $form.Controls.Add($box_pbix)
    $form.Controls.Add($button_file_browse)

    ##### Get output folder path #####
    # Add Objects
    $label_folder = New-Object System.Windows.Forms.Label -Property @{
        Font = $font_label
        Text = "Select an folder for the output CSVs"
        AutoSize = $false
        Width = (Get-X 300)
        Height = (Get-Y 15)
        Location = New-Object System.Drawing.Point($marginX, ($box_pbix.Location.Y + $box_pbix.Height + $v_space_wide))
    }

    $box_folder = New-Object System.Windows.Forms.TextBox -Property @{
        Font = $font_box
        MultiLine = $false
        AutoSize = $false
        Width = (Get-X 295)
        Height = (Get-Y 20)
        Location = New-Object System.Drawing.Point(($label_folder.Location.X +$h_space),($label_folder.Location.Y + $label_folder.Height + $v_space_close))
    }

    $button_folder_browse = New-Object System.Windows.Forms.Button -Property @{
        Font = $font_button
        Text = "Browse"
        AutoSize = $false
        Width = (Get-X 75)
        Height = (Get-Y 20)
        Location = New-Object System.Drawing.Point (($box_folder.Location.X + $box_folder.Width + $h_space), ($box_folder.Location.Y))
    }

    # Add Events
    $box_folder.add_Click({PopulateTextBox -object $box_folder -text (Get-FolderBrowser)})
    $button_folder_browse.add_Click({PopulateTextBox -object $box_folder -text (Get-FolderBrowser)})

    # Add to Form
    $form.Controls.Add($label_folder)
    $form.Controls.Add($box_folder)
    $form.Controls.Add($button_folder_browse)

    ##### Output Report Selection #####
    # Add Objects
    $label_choice = New-Object System.Windows.Forms.Label -Property @{
        Font = $label_font
        Text = "Select output documents"
        AutoSize = $false
        Width = (Get-X 300)
        Height = (Get-Y 15)
        Location = New-Object System.Drawing.Point($marginX, ($box_folder.Location.Y + $box_folder.Height + $v_space_wide))
    }

    $check_choice = New-Object System.Windows.Forms.CheckedListBox -Property @{
        Font = $box_font
        Text = "Something"
        AutoSize = $false
        Width = (Get-X 295)
        Height = (Get-Y 200)
        Location = New-Object System.Drawing.Point(($label_choice.Location.X + $h_space), ($label_choice.Location.Y + $label_choice.Height + $h_space))
    }

    # Add Items to CheckBoxList
    @("Data Definitions", "Visual Containers", "Unused Objects").ForEach({[void]$check_choice.Items.Add($_)})

    # Add Events

    # Add to Form
    $form.Controls.Add($label_choice)
    $form.Controls.Add($check_choice)

    ##### Go / No-Go #####
    # Add Objects
    $button_ok = New-Object System.Windows.Forms.Button -Property @{
        Text = "OK"
        AutoSize = $false
        Width = (Get-X 75)
        Height = (Get-Y 20)
        Location = New-Object System.Drawing.Point ((($form.Width - 10) - (2*$h_space) - (2*75)), ($form.Height - 60))
        Font = $button_font
    }

    $button_cancel = New-Object System.Windows.Forms.Button -Property @{
        Text = "Cancel"
        AutoSize = $false
        Width = (Get-X 75)
        Height = (Get-Y 20)
        Location = New-Object System.Drawing.Point (($button_ok.Location.X + $button_ok.Width + $h_space), ($button_ok.Location.Y))
        Font = $button_font
    }

    # Add Events
    $button_ok.Add_Click({OKButton})
    $button_cancel.Add_Click({$form.Close()})

    # Add to Form
    $form.Controls.Add($button_ok)
    $form.Controls.Add($button_cancel)

    ##### Show Form #####
    [void]$form.ShowDialog()

    ##### Set Global Variables #####
    $pbixFilePath = $box_pbix.Text
    $outputFolder = $box_folder.Text

    $essentials = Get-ListItems -list_object $check_choice
    $outputDocuments = @{
        DataDefinitions = ($essentials -contains "Data Definitions")
        VisualContainers = ($essentials -contains "Visual Containers")
        UnusedObjects = ($essentials -contains "Unused Objects")
    }
}

# Check if parameters actually set. If not, exit script.
if((Test-Param $pbixFilePath) -or (Test-Param $outputFolder)) {
    [Console]::WriteLine("Required parameters not set. Script will exit in 5 seconds.")
    Start-Sleep -Seconds 5
    Exit
}

# Set $outputDocuments to all if never included
if(Test-Param $outputDocuments) {
    $outputDocuments = @{
        DataDescriptions = $true
        VisualContainers = $true
        UnusedObjects = $true
    }
}

###################################### Module Check #####################################
# If using the Invoke-ASCmd, need SQl Server
# Install-Module -Name SqlServer
if (Get-Module -Name SqlServer -ListAvailable) {
    [Console]::WriteLine("SqlServer Module Already Installed...")
} else {
    [Console]::WriteLine('Installing SqlServer Module...')
    Install-Module -Name SqlServer -Force
}

################################# Configuration Settings ################################
[Console]::WriteLine('Setting Variables...')

# Set Variables (Internal)
$path = (Get-Item $pbixFilePath).Directory.FullName + "\"
$tmpPath = $path + 'tmp_read_pbix'
$tmp = $tmpPath + "\"

####################################### Functions #######################################
[Console]::WriteLine('Setting Functions...')

function Lookup-Left {
    Param($LTbl, $LJoin, $RTbl, $RJoin, $RColName)

    $return_val = 
        foreach($ob in $LTbl) {
            $lookup = $ob.$LJoin

            $ob | Select-Object -Property *, @{n=$RColName; e={($RTbl | Where-Object {$_.$RJoin -eq $lookup}).$RColName}}
        }

    return $return_val
}

function Parse-PrototypeQuery {
    Param($Query)

    $report_page = $Query.displayName
    $container_name = $Query.name
    $local_list = $Query.singleVisual.prototypeQuery
    $local_list_prop = $Query.singleVisual.columnProperties

    if($null -eq $local_list) {
        $local_objects = $null
    } else {
        $local_tables = $local_list.From | Select-Object -Property @{n='Table_ID';e={$_.Name}}, @{n='Table_Name';e={$_.Entity}}

        $local_measures = $local_list.Select | Select-Object -Property Name, @{n='Table_ID';e={$_.Measure.Expression.SourceRef.Source}}, @{n='Object_ID';e={$_.Measure.Property}}
        $local_columns = $local_list.Select | Select-Object -Property Name, @{n='Table_ID';e={$_.Column.Expression.SourceRef.Source}}, @{n='Object_ID';e={$_.Column.Property}}
        $local_aggregations = $local_list.Select.Aggregation | Select-Object -Property Name, @{n='Table_ID';e={$_.Aggregation.Expression.SourceRef.Source}}, @{n='Object_ID';e={$_.Aggregation.Property}}

        #Combine and build out local_objects
        $local_objects = @($local_measures, $local_columns, $local_aggregations).ForEach({$_ | Where-Object {$_.Object_ID -ne $null}})
        $local_objects = Lookup-Left -LTbl $local_objects -LJoin Table_ID -RTbl $local_tables -RJoin Table_ID -RColName Table_Name

        if($null -eq $local_list_prop){
            $local_objects = $local_objects | Select-Object -Property Table_Name, Object_ID, @{n="Display_Name"; e={$null}}
        } else {
            $local_objects = $local_objects | Select-Object -Property Table_Name, Object_ID, @{n="Display_Name"; e={$local_list_prop.($_.Name).displayName}}
        }
    }

    return ($local_objects | Select-Object -Property *, @{n="Container_Name"; e={$container_name}}, @{n="Report_Page"; e={$report_page}})
}

function Parse-VisualContainer {
    Param($Query)
    $Query.layouts.position | Select-Object -Property *,
                        @{n='Visual Container';e={$Query.name}},
                        @{n='Visual Type' ;e={$Query.singleVisual.visualType}},
                        @{n='Visual Title';e={$Query.singleVisual.vcObjects.title.properties.text.expr.Literal.Value}},
                        @{n='Report Page' ;e={$Query.displayName}}
}

function Get-DAXQueryXMLA {
    Param($Server, $Query)

    [xml]$db = Invoke-ASCmd -Server $Server -Query $Query
    return $db.return.root.row
}

function Invoke-ReductionPass {
    Param($local_dd, $local_page_reduced)

    #mr stands for model referenced
    #nr stands for not referenced
    $local_mr_pass = $local_dd | Select-Object -Property @{n='Table';e={$_.Tgt_Table}}, @{n='Object';e={$_.Tgt_Object}} -Unique
    [console]::WriteLine('Length of mr:' + $local_mr_pass.Count)

    $local_nr_pass = (Compare-Object -ReferenceObject $local_page_reduced -DifferenceObject $local_mr_pass -Property Table, Object).Where({$_.SideIndicator -eq '<='}) |
                     Select-Object -Property Table, Object

    $local_dd = (Compare-Object -ReferenceObject $local_dd -DifferenceObject $local_nr_pass -Property Table, Object).Where({$_.SideIndicator -eq '<='}) |
                Select-Object -Property Table, Object

    return $local_dd
}

###################################### Main Program ######################################
[Console]::WriteLine('Starting Main Program...')

##### Data from PBIX (Zip) File #####
[Console]::WriteLine('  Creating & getting data from zip file...')

# Create Path if needed
if(!(Test-Path -LiteralPath $tmpPath)){
    New-Item -ItemType Directory -Force -Path $tmpPath | Out-Null
}

#Copy and Expand
Copy-Item -LiteralPath $pbixFilePath -Destination ($tmpPath + "\tmp.zip") -Force
Expand-Archive -LiteralPath ($tmpPath + "\tmp.zip") -DestinationPath $tmpPath -Force

$content = Get-Content -LiteralPath ($tmpPath + "\Report\Layout") -Encoding Unknown | ConvertFrom-Json
$report_names = $content.sections.displayName | Get-Unique

#Expand Properties & Perform Conversions
#Note: this will also affect parent objects (thus the $nulls). Potential fix coming for Powershell Core, but not Windows Powershell.
$null = $content.sections | Select-Object -Property displayName -ExpandProperty visualContainers
$visContainer = $content.sections.visualContainers | Select-Object -Property displayName, x, y, @{n="config"; e={$_.config | ConvertFrom-Json}}

$null = $visContainer | Select-Object -Property displayName -ExpandProperty config
$config_values = $visContainer.config

#Building Prototype query function
$objects = $config_values.ForEach({Parse-PrototypeQuery -Query $_}) | Where-Object {$_ -ne $null}
$visualContainer = $config_values.ForEach({Parse-VisualContainer -Query $_})

#Remove tmp directory
Remove-Item -LiteralPath $tmpPath -Force -Recurse


##### Data from OLAP Queries #####
##### PBIX Start #####
[Console]::WriteLine('  Opening PBIX File at: ' + $pbixFilePath)
# Start PBIX to save time later
$pbix = Start-Process -FilePath $pbixFilePath -PassThru -WindowStyle Hidden
$pid_val = $pbix.Id

[Console]::WriteLine('  Getting OLAP data...')

# Check that PBIX file is actually open and ready for query
[bool]$cont_switch = 1
[int]$cont_cnt = 0

while (($cont_switch) -and ($cont_cnt -lt 30)){
    $remote_port = (Get-NetTCPConnection -OwningProcess $pid_val -State Established -ErrorAction SilentlyContinue).Where({$_.RemoteAddress -eq $_.LocalAddress})
    $cont_switch = @($null, "") -contains $remote_port
    [Console]::WriteLine('    ' + $cont_cnt)
    Start-Sleep -Seconds 2
    ++$cont_cnt
}

# Get Port and Connection information
$remote_port = ($remote_port | Measure-Object -Property RemotePort -Maximum).Maximum

$connection = "localhost:$remote_port"

# Open/Close connection and perform queries
$tables_data = Get-DAXQueryXMLA -Server $connection -Query 'SELECT * FROM $SYSTEM.MDSCHEMA_DIMENSIONS' | 
                Select-Object -Property @{n='Table_ID'; e={$_.DIMENSION_UNIQUE_NAME}}, @{n='Table'; e={$_.DIMENSION_NAME}}, @{n='Description'; e={$_.DESCRIPTION}}

$columns_data = Get-DAXQueryXMLA -Server $connection -Query 'SELECT * FROM $SYSTEM.MDSCHEMA_HIERARCHIES' | 
                Select-Object -Property @{n='Table_ID';e={$_.DIMENSION_UNIQUE_NAME}}, @{n='Object';e={$_.HIERARCHY_NAME}}, @{n='Description';e={$_.DESCRIPTION}}

$columns_data = Lookup-Left -LTbl $columns_data -LJoin Table_ID -RTbl $tables_data -RJoin Table_ID -RColName Table |
                Select-Object -Property Table, Object, Description

$measures_data = Get-DAXQueryXMLA -Server $connection -Query 'SELECT * FROM $SYSTEM.MDSCHEMA_MEASURES' | 
                 Select-Object -Property @{n='Table';e={$_.MEASUREGROUP_NAME}}, @{n='Object';e={$_.MEASURE_NAME}}, @{n='Description';e={$_.DESCRIPTION}}

# Only needed for unused objects, but easy to save the check.
$dependencies_data = Get-DAXQueryXMLA -Server $connection -Query 'SELECT * from $SYSTEM.DISCOVER_CALC_DEPENDENCY' |
                     Select-Object -Property @{n='Object';e={$_.TABLE +".[||SplitHere||]." + $_.OBJECT}}, @{n='Reference';e={$_.REFERENCED_TABLE +".[||SplitHere||]." + $_.REFERENCED_OBJECT}} -Unique

# Close PBIX File
Stop-Process -Id $pid_val -Force


##### Combining OLAP and Zip Data #####
[Console]::WriteLine('  Combining data and working through dependencies...')

$combined = @($measures_data, $columns_data).ForEach({$_ | Where-Object {$_.Table -ne $null}})

if($outputDocuments["VisualContainers"]){
    # Visual Container Section
    [Console]::WriteLine('  Putting together Visual Containers...')
    [Console]::WriteLine('    Output Visual Containers to directory: ' + $outputFolder)
    $visualContainer | Export-Csv -LiteralPath ($outputFolder + '\Visual Containers.csv') -NoTypeInformation
}

if($outputDocuments["DataDefinitions"]){
    # Data Definitions Section
    [Console]::WriteLine('  Putting together Data Definitions...')
    $objects_dd =
        foreach($ob in $objects){
            $ob_tbl = $ob.Table_Name
            $ob_object = $ob.Object_ID
            $ob | Select-Object *, @{n='Description';e={($combined | Where-Object {($_.Table -eq $ob_tbl) -and ($_.Object -eq $ob_object)}).Description}}
        } 
    

    $objects_dd = $objects_dd | Select-Object -Property @{n='Report Page';e={$_.Report_Page}},
                                                        @{n='Visual Container';e={$_.Container_Name}},
                                                        @{n='Column/Measure';e={$_.Object_ID}},
                                                        @{n='Display Names';e={$_.Display_Name}},
                                                        @{n='Description';e={$_.Description}}

    ##### Export / Output #####
    [Console]::WriteLine('    Output data to directory: ' + $outputFolder)
    $objects_dd | Export-Csv -LiteralPath ($outputFolder + '\Data Definitions.csv') -NoTypeInformation
}

if($outputDocuments["UnusedObjects"]){
    #Unused Objects Section
    [Console]::WriteLine('  Putting together Unused Objects...')

    $all_objects = ($combined | Select-Object @{n='Object';e={$_.Table + ".[||SplitHere||]." + $_.Object}} -Unique).Object
    $page_referenced = ($objects | Select-Object -Property @{n='Object';e={$_.Table_Name + ".[||SplitHere||]." + $_.Object_ID}} -Unique).Object
    $page_reduced = $all_objects.Where({$page_referenced -notcontains $_})

    #Run loop (max 100 runs) to deal break down dependencies
    [int]$i = 0
    [int]$intersect = 99

    while(($i -lt 100) -and ($intersect -gt 0)){
        # Main Script
        #mr stands for model referenced
        #nr stands for not referenced

        #Unique Reference from dd
        $mr_pass = ($dependencies_data | Select-Object -Property Reference -Unique).Reference

        #Page Reduce that are not in MR Pass (Reference)
        $nr_pass = $page_reduced.Where({$mr_pass -notcontains $_})
    
        # Continuation Check
        $intersect = ($dependencies_data | Select-Object Object -Unique | Where-Object {$nr_pass -contains $_.Object}).Count 
        ++$i

        #dd Update
        $dependencies_data = $dependencies_data.Where({$nr_pass -notcontains $_.Object})

        # Reporting
        [console]::WriteLine("      Pass $i : $intersect additional removals needed.")  
    }

    # Finalize output
    if($intersect -gt 0){
        $nr_pass = 'Maximum Iteration Reached (100). Contact script administrator.'
    } else {
        $not_referenced = $nr_pass.ForEach({
            $splitString = $_ -split "\.\[\|\|SplitHere\|\|\]\.", 2, [System.Management.Automation.SplitOptions]::RegexMatch
            $tbl = $splitString[0]
            $obj = $splitString[1]
            $description = ($combined.Where({($_.Table -eq $tbl) -and ($_.Object -eq $obj)})).Description

            return [PSCustomObject]@{Table = $tbl; Object=$obj; Description=$description }   
        })
    }

    # Output CSV
    [Console]::WriteLine('    Output Unused Objects to directory: ' + $outputFolder)
    $pbixDir = Get-Item $pbixFilePath
    $csv_name = $outputFolder + '\Unused Objects (' + ($pbixDir.Name -replace $pbixDir.Extension, "") + ").csv"
    $not_referenced | Export-Csv -LiteralPath $csv_name -NoTypeInformation

}

[Console]::WriteLine('Goodbye!')