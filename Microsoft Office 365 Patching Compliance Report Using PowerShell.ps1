# < Office 365 Patch compliance reprot using  PowerShell >
#.DESCRIPTION
 # <Office 365 Patch compliance reprot using  PowerShell update Windows Autopilot groupTags using  PowerShell and CSV>
#.Demo
#<YouTube video link-->https://www.youtube.com/@ChanderManiPandey
#.INPUTS
 # <Provide all required inforamtion in User Input Section >
#.OUTPUTS
 # <Office 365 Patch compliance reprot in csv >

#.NOTES
 
 <#
  Version:         1.0
  Author:          Chander Mani Pandey
  Creation Date:   09 Oct 2024
  
  Find Author on 
  Youtube:-        https://www.youtube.com/@chandermanipandey8763
  Twitter:-        https://twitter.com/Mani_CMPandey
  LinkedIn:-       https://www.linkedin.com/in/chandermanipandey
 #>

#=================================================================================================================================================
#------------------------------------------------------ User Input Section Start------------------------------------------------------------------
#=================================================================================================================================================
$tenantNameOrID  = "xxxx"       # Tenant Name or ID
$clientAppId     = "xxxx"       # Client Application ID
$clientAppSecret = "xxxx"   # Client Application Secret

$WorkingFolder   = "C:\TEMP\M365_Patching_Compliance_Status"     # Application reporting folder Location
$AppName         = "Microsoft 365 Apps for enterprise"           # Enter the application name e.g- "Microsoft 365 Apps for enterprise - en-us"
$Platform        = "Windows"                                     # Enter the $Platform e.g- Windows,AndroidFullyManagedDedicated,Other,AndroidWorkProfile,IOS,AndroidDeviceAdministrator
$FilterOperator  = "Like"                                        # Choose filter operator: 'like' or 'eq'


#=================================================================================================================================================
#------------------------------------------------------ User Input Section End--------------------------------------------------------------------
#=================================================================================================================================================
CLS
Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force -ErrorAction 'Stop'
$error.Clear() # Clear error history
$ErrorActionPreference = 'SilentlyContinue'

#=================================================================================================================================================
                                                        #PART-1
#=================================================================================================================================================

$overallStartTime = Get-Date
Write-Host "=================================== Creating $AppName Update Compliance Report =====================================" -ForegroundColor White
# Validate the filter operator input
Write-Host ""
if ($FilterOperator -notin @('like', 'eq')) {
    Write-Host "Invalid filter operator. Please enter 'like' or 'eq'." -ForegroundColor Red
    exit
}

# Define folder path based on $AppName and $Platform
$AppFolderPath = "$WorkingFolder\$AppName`_$Platform"
$ReportName = "AppInvRawData"
$Path = "$AppFolderPath\$ReportName\"

# Start overall time tracking
# Check if the $AppFolderPath folder exists, and delete it if it does

if (Test-Path -Path $AppFolderPath) {
    #Write-Host " "
        #Write-Host " "
    Write-Host "$AppName Application Folder present. Removing existing folder: $AppFolderPath" -ForegroundColor Cyan
    Remove-Item -Path $AppFolderPath -Recurse -Force | Out-Null
    Write-Host "" -ForegroundColor Yellow
}

# Create the new working folder and subfolder
#Write-Host " "
Write-Host "Creating working folder and subfolder at $AppFolderPath" -ForegroundColor Yellow
New-Item -ItemType Directory -Path $AppFolderPath -Force -InformationAction SilentlyContinue | Out-Null
New-Item -ItemType Directory -Path $Path -Force -InformationAction SilentlyContinue | Out-Null
Write-Host "Created working folder and subfolder" -ForegroundColor Green

#=================================================================================================================================================
# Check and install the Microsoft.Graph.Intune module
Write-Host ""
Write-Host "Checking for Microsoft.Graph.Intune module" -ForegroundColor Cyan
if (-not (Get-Module -Name "Microsoft.Graph.Intune" -ListAvailable)) {
    Write-Host "Installing Microsoft.Graph.Intune module" -ForegroundColor Cyan
    Install-Module -Name Microsoft.Graph.Intune -Force
}
Write-Host "Importing Microsoft.Graph.Intune module" -ForegroundColor Cyan
Import-Module Microsoft.Graph.Intune -Force -InformationAction SilentlyContinue

# Authentication
Write-Host "Setting up authentication for MS Graph" -ForegroundColor Cyan

$tenant = $tenantNameOrID
$authority = "https://login.windows.net/$tenant"
$clientId = $clientAppId
$clientSecret = $clientAppSecret

Update-MSGraphEnvironment -AppId $clientId -AuthUrl $authority -SchemaVersion "Beta" -Quiet -InformationAction SilentlyContinue
Connect-MSGraph -ClientSecret $clientSecret -InformationAction SilentlyContinue -Quiet

#=================================================================================================================================================
# Create request body and initiate export job
#Write-Host "" 
Write-Host "Initiating export job for '$AppName' application and for '$Platform' Platform" -ForegroundColor Yellow
$exportJobStartTime = Get-Date
$postBody = @{ 
    'reportName' = $ReportName 
    'search' = $AppName
}
$exportJob = Invoke-MSGraphRequest -HttpMethod POST -Url "DeviceManagement/reports/exportJobs" -Content $postBody
$exportJobEndTime = Get-Date
$exportJobDuration = $exportJobEndTime - $exportJobStartTime
Write-Host "Export Job initiated. Monitoring Downloading status..." -ForegroundColor Cyan
#Write-Host ""

# Polling for export job status
$dotCount = 0
$pollingStartTime = Get-Date
do {
    Start-Sleep -Seconds 2
    $exportJob = Invoke-MSGraphRequest -HttpMethod Get -Url "DeviceManagement/reports/exportJobs('$($exportJob.id)')" -InformationAction SilentlyContinue
    Write-Host -NoNewline '.'
    $dotCount++
    if ($dotCount -eq 100) {
        Write-Host ""
        $dotCount = 0
    }
} while ($exportJob.status -eq 'inprogress')

$pollingEndTime = Get-Date
$pollingDuration = $pollingEndTime - $pollingStartTime
Write-Host ""

if ($exportJob.status -eq 'completed') {
    $fileName = (Split-Path -Path $exportJob.url -Leaf).Split('?')[0]
    #Write-Host ""
    Write-Host "Export Job completed. Writing File $fileName to Disk..." -ForegroundColor Cyan
    $downloadStartTime = Get-Date
    Invoke-WebRequest -Uri $exportJob.url -Method Get -OutFile "$Path$fileName"
    $downloadEndTime = Get-Date
    $downloadDuration = $downloadEndTime - $pollingStartTime
    #Write-Host "Time taken to Initiate and download $AppName file: $([math]::Round($downloadDuration.TotalMinutes, 2)) minutes" -ForegroundColor White

    Remove-Item -Path "$Path*" -Include *.csv -Force
    Expand-Archive -Path "$Path$fileName" -DestinationPath $Path
    #=============================================================================================================================================
    # Processing CSV file
    #Write-Host ""
    #Write-Host "Processing CSV file. Please wait..." -ForegroundColor Cyan

    # Get the path of the CSV file
    $csvPath = Get-ChildItem -Path $Path -Filter *.csv | Where-Object {! $_.PSIsContainer} | Select-Object -ExpandProperty FullName

    # Check if the CSV file exists
    if (-not (Test-Path $csvPath)) {
        Write-Host "CSV file not found at $csvPath" -ForegroundColor Red
        exit
    }

    $overallCSVImportStartTime = Get-Date
    # Read the CSV file
    $csvData = Import-Csv -Path $csvPath

    $overallCSVImportEndTime = Get-Date
    $overallImportDuration = $overallCSVImportEndTime - $overallCSVImportStartTime
    #Write-Host "Time taken to Import CSV: $([math]::Round($overallImportDuration.TotalMinutes, 2)) minutes" -ForegroundColor White
    #Write-Host ""

    # Filter data based on user choice and Platform
    if ($FilterOperator -eq 'like') {
        $filteredData = $csvData | Where-Object {
            $_.ApplicationName -like "*$AppName*" -and $_.Platform -eq $Platform
        }
    } elseif ($FilterOperator -eq 'eq') {
        $filteredData = $csvData | Where-Object {
            $_.ApplicationName -eq $AppName -and $_.Platform -eq $Platform
        }
    }

    $processingStartTime = Get-Date
    # Write filtered data to CSV
    $filteredOutputPath = "$AppFolderPath\Filtered_$($AppName)_$($Platform).csv"
    $filteredData | Export-Csv -Path $filteredOutputPath -NoTypeInformation -Encoding utf8

    # End time for processing CSV
    $processingEndTime = Get-Date
    $processingDuration = $processingEndTime - $processingStartTime
    Write-Host "Processing complete. Filtered data saved to $filteredOutputPath" -ForegroundColor Cyan
    #Write-Host "Time taken to process CSV: $([math]::Round($processingDuration.TotalMinutes, 2)) minutes" -ForegroundColor White
    #Write-Host ""

    # Import and summarize the final report
    Write-Host "Summarizing final report for $AppName..." -ForegroundColor Cyan
    #Write-Host ""
    $finalizeStartTime = Get-Date
    $FinalReport = Import-Csv -Path $filteredOutputPath
    $TotalDevices = ($FinalReport | Measure-Object | Select-Object -ExpandProperty Count)
    $TotalApplicationVersions = ($FinalReport | Select-Object -ExpandProperty ApplicationVersion -Unique | Measure-Object | Select-Object -ExpandProperty Count)

    #Write-Host "Total Number of $Platform Devices where $AppName is Installed:          $TotalDevices" -ForegroundColor Yellow
    #Write-Host "Total Number of Detected Application Versions on $Platform Platform:         $TotalApplicationVersions" -ForegroundColor Yellow
    #Write-Host ""

    # Format and export the final report with dynamic filename
    Write-Host "Formatting and exporting final report..." -ForegroundColor Cyan
    $formattedReportStartTime = Get-Date
    $formattedReportPath = "$AppFolderPath\$($AppName)_Report.csv"
    $FormattedReport = $FinalReport | Select-Object DeviceName, UserName, EmailAddress, OSDescription, OSVersion, Platform, ApplicationName, ApplicationVersion, ApplicationPublisher
    $FormattedReport | Export-Csv -Path $formattedReportPath -NoTypeInformation

    # End time for finalizing report
    $formattedReportEndTime = Get-Date
    $formattedReportDuration = $formattedReportEndTime - $formattedReportStartTime
    #Write-Host "Time taken to finalize report: $([math]::Round($formattedReportDuration.TotalMinutes, 2)) minutes" -ForegroundColor White
    #Write-Host ""

    # End overall time tracking
    $overallEndTime = Get-Date
    $overallDuration = $overallEndTime - $overallStartTime
    #Write-Host "Overall time to complete the script: $([math]::Round($overallDuration.TotalMinutes, 2)) minutes" -ForegroundColor Green
}

#=================================================================================================================================================
# Define the list of paths to clean up

$pathsToRemove = @($filteredOutputPath, $Path)
#Write-Host "Performing Cleanup work" -ForegroundColor Cyan
foreach ($path in $pathsToRemove) {
    if (Test-Path -Path $path) {
        Remove-Item -Path $path -Recurse -Force
    }
}
#Write-Host ""
#Write-Host "Final report saved location: $formattedReportPath" -ForegroundColor Yellow
#Write-Host "Overall time to complete the script: $([math]::Round($overallDuration.TotalMinutes, 2)) minutes" -ForegroundColor White
#Write-Host ""
#Write-Host "============================= Successfully created $AppName Application Report for $Platform Platform ============================" -ForegroundColor Magenta
# Optionally output the formatted report to the console
#$FormattedReport | Out-GridView
#=================================================================================================================================================
                                                        #PART-2
#===================================================================================================================================================
#Write-host "============================= Windwos Device ===============" -ForegroundColor Magenta
#Write-host ""
$Starttime = get-date
$Path = $WorkingFolder
# Check if the path exists, and create it if not
if (-not (Test-Path -Path $Path -PathType Container)) {
    New-Item -Path $Path -ItemType Directory -Force
}
$R1 = "DevicesWithInventory"
# Post body for the export job
$postBody = @{
    'reportName' = $R1
    'filter'     = "(DeviceType  eq '1')"  # 1 = Windows
    'select'     = "DeviceId", "DeviceName","UserName","UserEmail", "LastContact", "OSVersion","StorageTotal","StorageFree","ManagementAgent","OwnerType","JoinType"
}
# Initiate export job
$exportJob = Invoke-MSGraphRequest -HttpMethod POST -Url "DeviceManagement/reports/exportJobs" -Content $postBody
#Write-Host ""
Write-Host "Export Job initiated for $R1 Report " -ForegroundColor Cyan
# Check export job status
do {
    $exportJob = Invoke-MSGraphRequest -HttpMethod Get -Url "DeviceManagement/reports/exportJobs('$($exportJob.id)')" -InformationAction SilentlyContinue
    Start-Sleep -Seconds 2
    Write-Host -NoNewline '.'
} while ($exportJob.status -eq 'inprogress')
Write-Host 'Report is in Ready(Completed) status for Downloading' -ForegroundColor Yellow
If ($exportJob.status -eq 'completed') {
    $fileName = (Split-Path -Path $exportJob.url -Leaf).split('?')[0]
    Write-host "Export Job completed.......  Writing File $fileName to Disk........" -ForegroundColor Yellow
    # Download the file
    Invoke-WebRequest -Uri $exportJob.url -Method Get -OutFile "$Path\$fileName"
    # Remove previous CSV files in the destination folder
    $removePath1 = Join-Path -Path $Path -ChildPath "$R1"
    # Ensure $removePath1 has a trailing backslash
    $removePath1 = if ($removePath1.EndsWith("\")) { $removePath1 } else { "$removePath1\" }
   if ($removePath1 -eq $null) {
    # Remove items if the path is not null
 #   Write-Host "Path is null. Nothing to remove."
       } else {
    # Do nothing if the path is null
    if (Test-Path -Path $removePath1 -PathType Container) 
    { Remove-Item -Path "$removePath1\*" -Include *.csv -Force }
    #Write-Host "Items removed from $removePath1*.csv"
    }
    # Extract the downloaded file to a specific folder
    $null =  Expand-Archive -Path "$Path\$fileName" -DestinationPath "$Path\$R1" -Force 
    # Construct the full path to the extracted file
    $extractedFilePath = Join-Path -Path "$Path\$R1" -ChildPath "$fileName"
    }
    Remove-Item -Path "$Path\*" -Include *.zip -Force

# Specify the folder where the CSV file is located
$folderPath = "$Path\$R1"

# Get the CSV file(s) in the specified folder
$csvFiles1 = Get-ChildItem -Path $folderPath -Filter *.csv

# Check if any CSV files were found
if ($csvFiles1.Count -eq 0) {
    Write-Host "No CSV files found in $folderPath" -ForegroundColor Red
} else {
    # Assuming you want to import the first CSV file found
    $csvFile1 = $csvFiles1[0]
    $FINALPATH = $csvFile1.FullName

    $IntuneDevice = Import-Csv -Path $FINALPATH
}
#$IntuneDevice

#==============================================================================================================================================================
                                                        #PART-3
#==============================================================================================================================================================

# Initialize variables
$CollectedData = @()

# Download Office Master Patch List
Write-Host "Downloading Office 365 Patch List from Microsoft" -ForegroundColor yellow
$URI = "https://learn.microsoft.com/en-us/officeupdates/update-history-microsoft365-apps-by-date"
$response = Invoke-WebRequest -Uri $URI -UseBasicParsing -ErrorAction Continue
$CollectedData = $response.Links

# Extract outerHTML content from elements where outerHTML contains "channel"
$CollectedDataAll = $CollectedData | Where-Object { $_.outerHTML -match "channel" -and $_.outerHTML -cnotmatch "DeployOffice/overview" } | ForEach-Object {
    # Extract href and split into channel and version
    $hrefParts = $_.href -split '#'
    $channel = $hrefParts[0]
    $version = if ($hrefParts.Count -gt 1) { $hrefParts[1] } else { $null }

    # Extract ReleaseDate from version
    $ReleaseDate = if ($version) { $version -replace "version-", "" } else { $null }

    # Extract BuildVersion from outerHTML
    $BuildVersion = $null
    if ($_.outerHTML -match "\((Build\s+\d+\.\d+)\)") {
        $BuildVersion = [regex]::Match($_.outerHTML, "\((Build\s+\d+\.\d+)\)").Groups[1].Value -replace "Build ", ""
    }

    # Initialize year
    $year = $null

    # Determine the year based on ReleaseDate prefix
    if ($ReleaseDate) {
        $prefix = ($ReleaseDate -split '-')[0]
        if ($prefix.Length -eq 4) {
            $prefixYear = [int]$prefix.Substring(0, 2)
            $year = switch ($prefixYear) { 30 { "2030" } 29 { "2029" } 28 { "2028" } 27 { "2027" } 26 { "2026" } 25 { "2025" } 24 { "2024" } 23 { "2023" } 22 { "2022" } 21 { "2021" } 20 { "2020" } 19 { "2019" } 18 { "2018" } 17 { "2017" } 16 { "2016" } 15 { "2015" } default { "Unknown" } }
        } else { $year = "Unknown" }
    }

    # Extract the month and day from ReleaseDate
    $month = $null
    $day = $null
    if ($ReleaseDate -match "(\d{4})-(\w+)-(\d{1,2})") {
        $monthMap = @{
            "January" = "01"; "February" = "02"; "March" = "03"; "April" = "04"; "May" = "05"; "June" = "06"; "July" = "07"; "August" = "08"; "September" = "09"; "October" = "10"; "November" = "11"; "December" = "12"
        }
        
        $monthName = $matches[2]
        $month = if ($monthMap.ContainsKey($monthName)) { $monthMap[$monthName] } else { "Unknown" }
        $day = $matches[3].PadLeft(2, '0')
    }

    # Create R_Date by combining month, day, and year
    $R_DateString = if ($month -ne "Unknown" -and $day -ne $null -and $year -ne "Unknown") {
        "$year-$month-$day"
    } else {
        $null
    }

    # Convert R_DateString to DateTime object
    $R_Date = if ($R_DateString) {
        [DateTime]::ParseExact($R_DateString, "yyyy-MM-dd", $null)
    } else {
        $null
    }

    # Create PatchReleaseDate
    $PatchReleaseDate = if ($month -ne "Unknown" -and $day -ne $null -and $year -ne "Unknown") {
        [DateTime]::ParseExact("$month-$day-$year", "MM-dd-yyyy", $null)
    } else {
        $null
    }

    # Clean up Channel value
    if ($channel -match "-\d{4}$") {
        $channel = $channel -replace "-\d{4}$", ""
    }
    
    # Ensure no trailing dashes remain
    $channel = $channel.TrimEnd('-')

    # Create OfficeBuildVersion string
    $OfficeBuildVersion = if ($BuildVersion) { "16.0.$BuildVersion" } else { $null }

    # Extract Latest_Version from Version
    $Latest_Version = if ($version -match "version-(\d+)") { $matches[1] } else { $null }

    # Create a custom object with the desired properties
    [PSCustomObject]@{
        Channel             = $channel
        Version             = $version
        Latest_Version      = $Latest_Version
        Year                = $year
        ReleaseDate         = $ReleaseDate
        Day                 = $day
        BuildVersion        = $BuildVersion
        OfficeBuildVersion  = $OfficeBuildVersion
        R_Date              = $R_Date
        Month               = $month
        PatchReleaseDate    = if ($PatchReleaseDate) { $PatchReleaseDate.ToString("MM-dd-yyyy") } else { $null }
    }
}

# Group by Channel, Version, and ReleaseDate and select the entry with the highest BuildVersion
$GroupedData = $CollectedDataAll | Group-Object Channel, Version, ReleaseDate | ForEach-Object {
    $_.Group | Sort-Object -Property BuildVersion -Descending | Select-Object -First 1
}

# Initialize a variable to track the previous day and year
$previousDay = $null
$previousYear = $null

# Iterate through grouped data to apply year correction logic
$ProcessedData = $GroupedData | ForEach-Object {
    if ($_.Day -eq $previousDay) {
        $_.Year = $previousYear
    } else {
        $previousDay = $_.Day
        $previousYear = $_.Year
    }

    $_.R_Date = if ($_.Month -ne "Unknown" -and $_.Day -ne $null -and $_.Year -ne "Unknown") {
        $R_DateString = "$($_.Year)-$($_.Month)-$($_.Day)"
        [DateTime]::ParseExact($R_DateString, "yyyy-MM-dd", $null)
    } else {
        $null
    }

    $_.PatchReleaseDate = if ($_.Month -ne "Unknown" -and $_.Day -ne $null -and $_.Year -ne "Unknown") {
        [DateTime]::ParseExact("$($_.Month)-$($_.Day)-$($_.Year)", "MM-dd-yyyy", $null).ToString("MM-dd-yyyy")
    } else {
        $null
    }

    $_.OfficeBuildVersion = if ($_.BuildVersion) { "16.0.$($_.BuildVersion)" } else { $null }

    $_
}

# Determine compliance for each channel based on R_Date
$ChannelList = $ProcessedData | Select-Object -ExpandProperty Channel -Unique 

# Define the list of channels you want to keep
$DesiredChannels = @(    'current-channel',    'monthly-enterprise-channel',      'semi-annual-enterprise-channel',    'semi-annual-enterprise-channel-preview')

# Filter $ChannelList to only include the desired channels
$FinalResults = @()

foreach ($channel in $ChannelList) {
    $ChannelData = $ProcessedData | Where-Object { $_.Channel -eq $channel }
    if ($ChannelData) {
        # Determine the most recent R_Date for the current channel
        $LatestR_Date = ($ChannelData | Sort-Object -Property R_Date -Descending | Select-Object -First 1).R_Date

        # Mark all entries with the most recent R_Date as Compliant
        $ChannelDataWithCompliance = $ChannelData | ForEach-Object {
            [PSCustomObject]@{
                Channel             = $_.Channel
                Version             = $_.Version
                Latest_Version      = $_.Latest_Version
                Year                = $_.Year
                ReleaseDate         = $_.ReleaseDate
                Day                 = $_.Day
                BuildVersion        = $_.BuildVersion
                OfficeBuildVersion  = $_.OfficeBuildVersion
                Month               = $_.Month
                LatestReleaseDate   = $_.R_Date
                PatchReleaseDate    = $_.PatchReleaseDate
                Compliance          = if ($_.R_Date -eq $LatestR_Date) { "Compliant" } else { "Non-Compliant" }
            }
        }

        $FinalResults += $ChannelDataWithCompliance
    }
}

# Sort FinalResults based on the custom order before exporting to CSV
$customOrder = @(    'monthly-enterprise-channel',    'semi-annual-enterprise-channel',    'current-channel',    'semi-annual-enterprise-channel-preview',    'semi-annual-channel',
    'monthly-channel',    'semi-annual-channel-targeted',    'monthly-enterprise-channel-archived',    'monthly-channel-archived',    'semi-annual-enterprise-channel-archived',    'semi-annual-enterprise-channel-preview-archived'
)

$FinalResults = $FinalResults | Sort-Object -Property @{Expression={ [array]::IndexOf($customOrder, $_.Channel) }; Ascending=$true}

# Define the path where the CSV will be saved
$M365 = "M365_Release_Patch_List"
$M365PList = "$WorkingFolder\$M365\M365_PatchList.csv"
$M365FileRemovePath = "$WorkingFolder\$M365"

# Check if the M365 folder exists, and create it if it doesn't
if (-not (Test-Path "$WorkingFolder\$M365")) {
    New-Item -Path "$WorkingFolder\$M365" -ItemType Directory | Out-Null
}

# Export the sorted data to a CSV file
$FinalResults | Select-Object Channel, Version, Latest_Version, Year, ReleaseDate, Day, BuildVersion, OfficeBuildVersion, Month, LatestReleaseDate, PatchReleaseDate, Compliance | Export-Csv -Path $M365PList -NoTypeInformation

# Output to GridView for review
#$FinalResults | Out-GridView


#=====================================================================================================================================
                                                 #PART-4  Final Report
#=======================================================================================================================================

# Initialize the report collection
$FinalIntuneReport = @()

# Create a hash table for faster lookups of formatted entries
$formattedLookup = @{}
foreach ($entry in $FormattedReport) {
    $formattedLookup[$entry.DeviceName] = $entry
}

# Get the total number of devices for the progress bar
$totalDevices = $IntuneDevice.Count

# Initialize counters for compliance status
$complianceCounts = @{
    Compliant = 0
    NonCompliant = 0
    ManuallyCheck = 0
    M365AppsNotInstalled = 0
}

foreach ($ID in $IntuneDevice) {
    # Update the progress bar
    $currentIndex = $IntuneDevice.IndexOf($ID) + 1
    Write-Progress -Activity "Generating Final $AppName Compliance Report" -Status "Processing device $currentIndex of $totalDevices" -PercentComplete (($currentIndex / $totalDevices) * 100)

    # Lookup the formatted entry directly
    $formattedEntry = $formattedLookup[$ID."Device name"]
    
    # Default values
    $appChannel = 'ManuallyCheck'
    $compliance = 'ManuallyCheck'
    $LatestReleaseDate = 'ManuallyCheck'  # Default value

    # Check the Last_Check_In value
    $lastCheckInRaw = $ID."Last check-in"
    $lastCheckIn = 'Not Available'  # Default value

    if ($lastCheckInRaw) {
        # Attempt to parse the datetime directly
        try {
            $lastCheckIn = [datetime]::Parse($lastCheckInRaw).ToString("dd MMMM yyyy")
        } catch {
            Write-Host "Invalid Last_Check_In format for Device: $($ID."Device name") - Raw Value: $lastCheckInRaw"
        }
    }

    if ($formattedEntry) {
        $applicationVersion = $formattedEntry.ApplicationVersion

        # Check if ApplicationVersion indicates that apps are not installed
        if ($applicationVersion -eq 'M365 Apps Not Installed') {
            $appChannel = 'M365 Apps Not Installed'
            $compliance = 'M365 Apps Not Installed'
            $LatestReleaseDate = 'M365 Apps Not Installed'
            $complianceCounts.M365AppsNotInstalled++
        } else {
            # Get the relevant data from FinalResults once
            $finalResult = $FinalResults | Where-Object { $_.OfficeBuildVersion -eq $applicationVersion } | Select-Object -First 1

            if ($finalResult) {
                $appChannel = $finalResult.Channel
                $compliance = $finalResult.Compliance
                $LatestReleaseDate = if ($formattedEntry.ApplicationName -ne 'Not Found') {
                    $finalResult.LatestReleaseDate.ToString("dd MMMM yyyy")
                } else {
                    'M365_Apps_Not_Installed'
                }
            }

            if ($formattedEntry.ApplicationName -eq 'Not Found') {
                $compliance = 'M365 Apps Not Installed'
                $complianceCounts.M365AppsNotInstalled++
            }
        }
        
        # Adjust counters for compliance based on the value of compliance
        switch ($compliance) {
            'Compliant' {
                $complianceCounts.Compliant++
            }
            'Non-Compliant' {
                $complianceCounts.NonCompliant++
            }
            'Manually Check' {
                $complianceCounts.ManuallyCheck++
            }
            'M365 Apps Not Installed' {
                $complianceCounts.M365AppsNotInstalled++
            }
            default {
                $complianceCounts.ManuallyCheck++
            }
        }
    } else {
        # Handle cases where there is no formatted entry
        $appChannel = 'M365 Apps Not Installed'
        $compliance = 'M365 Apps Not Installed'
        $LatestReleaseDate = 'M365 Apps Not Installed'
        $complianceCounts.M365AppsNotInstalled++
    }

    # Create the report object
    $FinalIntuneReport += [PSCustomObject]@{
        Device_name        = $ID."Device name"
        Primary_user_name  = $ID."Primary user display name"
        User_Mail          = $ID."Primary user email address"
        Last_Check_In      = $lastCheckIn
        OS_Version         = $ID."OS version"
        Ownership          = $ID.Ownership 
        JoinType           = $ID.JoinType
        ApplicationName    = if ($formattedEntry) { $formattedEntry.ApplicationName } else { 'M365 Apps Not Installed' }
        ApplicationVersion = if ($formattedEntry) { $formattedEntry.ApplicationVersion } else { 'M365 Apps Not Installed' }
        AppChannel         = $appChannel
        PatchReleaseDate   = $LatestReleaseDate
        Compliance         = $compliance
    }
}

# Complete the progress bar
Write-Progress -Activity "Generating $AppName Final Compliance Status" -Status "Completed" -Completed

# Calculate the number of scoped devices
$scopedDevices = $totalDevices - $complianceCounts.M365AppsNotInstalled - $complianceCounts.ManuallyCheck

# Calculate compliance percentage
$compliancePercentage = if ($scopedDevices -gt 0) {
    ($complianceCounts.Compliant / $scopedDevices) * 100
} else {
    0
}

# Define the list of paths to clean up
$pathsToRemove = @($AppFolderPath,$folderPath,$M365FileRemovePath)
Write-Host "Performing Cleanup work" -ForegroundColor Cyan
foreach ($path in $pathsToRemove) {
    if (Test-Path -Path $path) {
        Remove-Item -Path $path -Recurse -Force
    }
}
# Output results
Write-Host ""
Write-Host "------ $AppName Compliance Status ------" -ForegroundColor White
Write-Host ""
Write-Host "Total Devices:               $totalDevices" -ForegroundColor Yellow
Write-Host "Compliant Devices:           $($complianceCounts.Compliant)" -ForegroundColor Green
Write-Host "Non-Compliant Devices:       $($complianceCounts.NonCompliant)" -ForegroundColor Red
Write-Host "Manually Check Devices:      $($complianceCounts.ManuallyCheck)" -ForegroundColor Gray
Write-Host "M365 Apps Not Installed:     $($complianceCounts.M365AppsNotInstalled)" -ForegroundColor Gray
Write-Host "Compliance Percentage:       $([math]::Round($compliancePercentage, 2))%" -ForegroundColor Cyan

# Output the final report
#$FinalIntuneReport | Out-GridView
$FinalIntuneReport | Export-Csv -Path "$WorkingFolder\Microsoft_Office_365_Update_Compliance_Report.csv" -NoTypeInformation 
Write-Host""
Write-Host "Report Location : - "$WorkingFolder\Microsoft_Office_365_Update_Compliance_Report.csv"" -ForegroundColor Green

Write-Host ""
Write-Host "=================================== Created $AppName Update Compliance Report =====================================" -ForegroundColor White




