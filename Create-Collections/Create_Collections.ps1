<#
    .DESCRIPTION
        This script can be used to bulk create collections in Configuration Manager
	
    .PARAMETER SiteCode
        Site Code of the target Configuration Manager envrionment
	
    .PARAMETER SiteServer
        Name of the Configuration Manager primary site server
	
    .PARAMETER CsvPath
        Path to the CSV file that contains the collections to be created
	
    .EXAMPLE
        Create_Collections.ps1 -SiteCode PS1 -SiteServer cm.example.local -CsvPath "C:\Temp\Collections.csv"
	
    .NOTES
        Created by: Jon Anderson
        Reference: https://www.configjon.com/configuration-manager-collection-creation-script/
        Modified: 02/17/2020
#>

#Parameters ===================================================================================================================
param(
    [Parameter(Mandatory=$true)][ValidateLength(3,3)][string]$SiteCode,
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$SiteServer,
    [ValidateScript({
        if (!($_ | Test-Path))
        {
            throw "The specified file does not exist"
        }
        if (!($_ | Test-Path -PathType Leaf))
        {
            throw "The Path argument must be a file. Folder paths are not allowed."
        }
        if ($_ -notmatch "(\.csv)")
        {
            throw "The specified file must be a .csv file"
        }
        return $true 
    })]
    [System.IO.FileInfo]$CsvPath
)

#Main program =================================================================================================================

#Get the current working directory
$OriginalLocation = Get-Location | Select-Object -ExpandProperty Path

#Check if the Configuration Manager console is installed
$DriveLetter = Get-Volume | Select-Object -ExpandProperty DriveLetter
ForEach($Letter in $DriveLetter){
    $ModulePath = "$($Letter):\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"
    If(Test-Path $ModulePath){Break}
    $ModulePath = "$($Letter):\Program Files\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"
    If(Test-Path $ModulePath){Break}
    Clear-Variable -Name "ModulePath"
}
If($NULL -eq $ModulePath)
{
    Throw "Could not find \Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1. This script should be run from a computer with the Configuration Manager console installed."
}

#Connect to the Configuration Manager site
If (!(Test-Path $SiteCode":"))
{
    Import-Module $ModulePath
    New-PSDrive -Name $SiteCode -PSProvider "AdminUI.PS.Provider\CMSite" -Root $SiteServer -ErrorAction SilentlyContinue
}
Set-Location $SiteCode":"

#Import the CSV file
$csv = Import-Csv $CsvPath

#Create the collections
$Counter = 0
while($Counter -lt $csv.CollectionName.Count)
{
    if($csv.FolderPath[$Counter])
    {
        #Create the folder path if it does not exist
        $FolderSplit = $csv.FolderPath[$Counter].Split('\')

        $FolderCounter = 0
        while($FolderCounter -lt $FolderSplit.Count)
        {
            if($TempPath)
            {
                if(!(Test-Path "$($SiteCode):\$($csv.CollectionType[$Counter])Collection$($TempPath)\$($FolderSplit[$FolderCounter])"))
                {
                    Write-Output "Creating new folder: $($SiteCode):\$($csv.CollectionType[$Counter])Collection$($TempPath)\$($FolderSplit[$FolderCounter])"
                    New-Item -Name $FolderSplit[$FolderCounter] -Path "$($SiteCode):\$($csv.CollectionType[$Counter])Collection$($TempPath)"
                }
            }
            else
            {
                if(!(Test-Path "$($SiteCode):\$($csv.CollectionType[$Counter])Collection\$($FolderSplit[$FolderCounter])"))
                {
                    Write-Output "Creating new folder: $($SiteCode):\$($csv.CollectionType[$Counter])Collection\$($FolderSplit[$FolderCounter])"
                    New-Item -Name $FolderSplit[$FolderCounter] -Path "$($SiteCode):\$($csv.CollectionType[$Counter])Collection\"
                }
            }

            $TempPath += "\$($FolderSplit[$FolderCounter])"
            $FolderCounter++
        }
    }

    #Create a refersh schedule
    $Schedule = New-CMSchedule -RecurInterval $csv.ScheduleInterval[$Counter] -RecurCount $csv.IntervalCount[$Counter]

    if($csv.IncrementalUpdate[$Counter] -eq "Yes")
    {
        #Create a collection with incremental updates enabled
        if($csv.FolderPath[$Counter])
        {
            New-CMCollection -Name $csv.CollectionName[$Counter] -CollectionType $csv.CollectionType[$Counter] -LimitingCollectionName $csv.LimitingCollectionName[$Counter] -Comment $csv.Comment[$Counter] -RefreshSchedule $Schedule -RefreshType Both | Move-CMObject -FolderPath "$($SiteCode):\$($csv.CollectionType[$Counter])Collection\$($csv.FolderPath[$Counter])"
        }
        else
        {
            New-CMCollection -Name $csv.CollectionName[$Counter] -CollectionType $csv.CollectionType[$Counter] -LimitingCollectionName $csv.LimitingCollectionName[$Counter] -Comment $csv.Comment[$Counter] -RefreshSchedule $Schedule -RefreshType Both | Out-Null
        }
    }
    else
    {
        #Create a collection with incremental updates disabled
        if($csv.FolderPath[$Counter])
        {
            New-CMCollection -Name $csv.CollectionName[$Counter] -CollectionType $csv.CollectionType[$Counter] -LimitingCollectionName $csv.LimitingCollectionName[$Counter] -Comment $csv.Comment[$Counter] -RefreshSchedule $Schedule | Move-CMObject -FolderPath "$($SiteCode):\$($csv.CollectionType[$Counter])Collection\$($csv.FolderPath[$Counter])"
        }
        else
        {
            New-CMCollection -Name $csv.CollectionName[$Counter] -CollectionType $csv.CollectionType[$Counter] -LimitingCollectionName $csv.LimitingCollectionName[$Counter] -Comment $csv.Comment[$Counter] -RefreshSchedule $Schedule | Out-Null
        }
    }
    
    #Create a query rule
    if($csv.QueryRule[$Counter])
    {
        if($csv.CollectionType[$Counter] -eq "Device")
        {
            Add-CMDeviceCollectionQueryMembershipRule -CollectionName $csv.CollectionName[$Counter] -RuleName $csv.QueryName[$Counter] -QueryExpression $csv.QueryRule[$Counter]
        }
        elseif($csv.CollectionType[$Counter] -eq "User")
        {
            Add-CMUserCollectionQueryMembershipRule -CollectionName $csv.CollectionName[$Counter] -RuleName $csv.QueryName[$Counter] -QueryExpression $csv.QueryRule[$Counter]
        }
    }

    #Create a include rule
    if($csv.IncludeCollection[$Counter])
    {
        $IncludeSplit = ($csv.IncludeCollection[$Counter]).Split(';')
        $IncludeCount = 0
        while($IncludeCount -lt $IncludeSplit.Count)
        {
            if($csv.CollectionType[$Counter] -eq "Device")
            {
                Add-CMDeviceCollectionIncludeMembershipRule -CollectionName $csv.CollectionName[$Counter] -IncludeCollectionName $IncludeSplit[$IncludeCount].Trim()
            }
            elseif($csv.CollectionType[$Counter] -eq "User")
            {
                Add-CMUserCollectionIncludeMembershipRule -CollectionName $csv.CollectionName[$Counter] -IncludeCollectionName $IncludeSplit[$IncludeCount].Trim()
            }
            $IncludeCount++
        }
    }

    #Create a exclude rule
    if($csv.ExcludeCollection[$Counter])
    {
        $ExcludeSplit = ($csv.ExcludeCollection[$Counter]).Split(';')
        $ExcludeCount = 0
        while($ExcludeCount -lt $ExcludeSplit.Count)
        {
            if($csv.CollectionType[$Counter] -eq "Device")
            {
                Add-CMDeviceCollectionExcludeMembershipRule -CollectionName $csv.CollectionName[$Counter] -ExcludeCollectionName $ExcludeSplit[$ExcludeCount].Trim()
            }
            elseif($csv.CollectionType[$Counter] -eq "User")
            {
                Add-CMUserCollectionExcludeMembershipRule -CollectionName $csv.CollectionName[$Counter] -ExcludeCollectionName $ExcludeSplit[$ExcludeCount].Trim()
            }
            $ExcludeCount++
        }
    }

    #Report if the collection was successfully created
    $CollectionCheck = Get-CMCollection -CollectionType $csv.CollectionType[$Counter] -Name $csv.CollectionName[$Counter]

    if($CollectionCheck)
    {
        Write-Output "Successfully created collection: $($csv.CollectionName[$Counter])"
    }
    else
    {
        Write-Host "Failed to create collection: $($csv.CollectionName[$Counter])"
    }

    Clear-Variable TempPath -ErrorAction SilentlyContinue
    $Counter++
}

#Set the working directory back to the original location
Set-Location $OriginalLocation