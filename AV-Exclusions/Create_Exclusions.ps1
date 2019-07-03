<#
    .DESCRIPTION
        This script can be used to automatically add exclusion settings
        to a new or existing antimalware policy in Configuration Manager.
	
    .PARAMETER SiteCode
        Site Code of the target Configuration Manager envrionment
	
    .PARAMETER SiteServer
        Name of the Configuration Manager primary site server
	
    .PARAMETER CsvPath
        Path to the CSV file that contains the antimalware exclusions
    
    .PARAMETER PolicyName
        Name of the anitmalware policy to create or modify
	
    .EXAMPLE
        Create_Exclusions.ps1 -SiteCode PS1 -SiteServer cm.example.local -CsvPath "C:\Temp\Exchange.csv" -PolicyName "Exchange AV Exclusions"
	
    .NOTES
        Created by: Jon Anderson
        Reference: https://www.configjon.com/create-configuration-manager-antimalware-policies-with-powershell/
#>

#Parameters
param(
    [Parameter(Mandatory=$true)][ValidateLength(3,3)][string]$SiteCode,
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$SiteServer,
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$CsvPath,
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$PolicyName
)

#Get the current working directory
$OriginalLocation = Get-Location | Select-Object -ExpandProperty Path

#Check if the Configuration Manager console is installed
$DriveLetter = Get-Volume | Select-Object -ExpandProperty DriveLetter
ForEach ($Letter in $DriveLetter){
    $ModulePath = "$($Letter):\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"
    If (Test-Path $ModulePath){Break}
    $ModulePath = "$($Letter):\Program Files\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"
    If (Test-Path $ModulePath){Break}
    Clear-Variable -Name "ModulePath"
}
If ($NULL -eq $ModulePath) {
    Throw "Could not find \Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1. This script should be run from a computer with the Configuration Manager console installed."
}

#Connect to the Configuration Manager site
If (!(Test-Path $SiteCode":")) {
    Import-Module $ModulePath
    New-PSDrive -Name $SiteCode -PSProvider "AdminUI.PS.Provider\CMSite" -Root $SiteServer
}
Set-Location $SiteCode":"

#Create the specified policy object if it does not already exist
If (!(Get-CMAntimalwarePolicy -Name $PolicyName)) {
    New-CMAntimalwarePolicy -Name $PolicyName -Policy ExclusionSettings
}

#Declare the array objects
$FilePath = @()
$FileType = @()
$Process = @()

#Import CSV data into the arrays
Import-Csv $CsvPath | ForEach-Object {
    $FilePath += $_.FilePath
    $FileType += $_.FileType
    $Process += $_.Process
}

#Remove blank lines from the arrays
$FilePath = $filePath | Where-Object {$_}
$FileType = $fileType | Where-Object {$_}
$Process = $process | Where-Object {$_}

#If the array is not empty, add the file paths to the specified antimalware policy
If ($FilePath.count -gt 0) {
    Set-CMAntimalwarePolicy -Name $PolicyName -ExcludeFilePath $FilePath
}

#If the array is not empty, add the file types to the specified antimalware policy
If ($FileType.count -gt 0) {
    Set-CMAntimalwarePolicy -Name $PolicyName -ExcludeFileType $FileType
}

#If the array is not empty, add the processes to the specified antimalware policy
If ($Process.count -gt 0) {
    Set-CMAntimalwarePolicy -Name $PolicyName -ExcludeProcess $Process
}

#Set the working directory back to the original location
Set-Location $OriginalLocation