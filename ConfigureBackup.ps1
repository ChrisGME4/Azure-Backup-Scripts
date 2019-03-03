#Requires -Version 3.0

$TaskName = "Azure Backup Result"

[Security.Principal.WindowsPrincipal]$Identity = [Security.Principal.WindowsIdentity]::GetCurrent()            
if (!$Identity.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw "In order for this script to configure a scheduled task, it needs to be run from an elevated command prompt, please run as Admininstrator."
}

function Install-PowerShellPackageManagement {
<#

   This Function will Check and see if PowerShellGet (aka PowerShellPackageManagement) is installed on your system

   This uses System.Net.WebRequest & System.Net.WebClient to download the specific version of PowerShellPackageManager for your OS version (x64/x86) and then uses
   msiexec to install it.
#>

#Requires -Version 3.0
[Cmdletbinding()]
param(
[Switch]$Latest

)

if (!(Get-command -Module PowerShellGet).count -gt 0)
    {
    if ($Latest) {
    $x86 = 'https://download.microsoft.com/download/C/4/1/C41378D4-7F41-4BBE-9D0D-0E4F98585C61/PackageManagement_x86.msi'
    $x64 = 'https://download.microsoft.com/download/C/4/1/C41378D4-7F41-4BBE-9D0D-0E4F98585C61/PackageManagement_x64.msi'
        Write-Verbose "Using the March 2016 Version of the MSI installer"
    }
    Else{ 
        $x86 = 'https://download.microsoft.com/download/4/1/A/41A369FA-AA36-4EE9-845B-20BCC1691FC5/PackageManagement_x86.msi'
        $x64 = 'https://download.microsoft.com/download/4/1/A/41A369FA-AA36-4EE9-845B-20BCC1691FC5/PackageManagement_x64.msi'
        Write-Verbose "Using the Pre-March 2016 Version of the MSI installer"
        }
    switch ($env:PROCESSOR_ARCHITECTURE)
    {
        'x86' {$version = $x86}
        'AMD64' {$version = $x64}
    }
    Write-Verbose "You are on a $version based OS and we are starting the Download of the MSI for your OS Version"
    $Request = [System.Net.WebRequest]::Create($version)
    $Request.Timeout = "100000000"
    $URL = $Request.GetResponse()
    $Filename = $URL.ResponseUri.OriginalString.Split("/")[-1]
    $url.close()
    $WC = New-Object System.Net.WebClient
    $WC.DownloadFile($version,"$env:TEMP\$Filename")
    $WC.Dispose()
    Write-Verbose "MSI Downloaded - Now executing the MSI to add the PackageManagement functionality to your Machine"
    msiexec.exe /package "$env:TEMP\$Filename" /q
    
    Start-Sleep 80
    Write-Verbose "MSI installed now removing the temporary file from your machine"
    Remove-Item "$env:TEMP\$Filename"
    }
}

$AzureModule = Get-Module Azure -ListAvailable
if (!$AzureModule) {
    Write-Host "AzureRM module not installed, attempting to install..."
    $PoshGet = Get-Module PowerShellGet -ListAvailable
    if (!($PoshGet.Version.Major)) {
        Write-Host "Installing PowerShell Package Management components..."
        Install-PowerShellPackageManagement -Latest -Verbose
        $env:PSModulePath += ";C:\Program Files\WindowsPowerShell\Modules"
    }

    Install-PackageProvider Nuget -Force

    Install-Module Azure -Force
    Import-Module Azure
}

$ScriptFolder = Split-Path $MyInvocation.InvocationName -Parent
if ($ScriptFolder -eq ".") { $ScriptFolder = Get-Location }
$ScriptFullPath = Join-Path $ScriptFolder "AzureBackup.ps1"
if (!(Test-Path $ScriptFullPath) -or (!$ScriptFullPath)) {
    throw "Unable to locate required script file $ScriptFullPath."
}

$TASKTEMPLATE = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>2017-06-21T18:27:29.9506422</Date>
    <Author>BackupRadar</Author>
  </RegistrationInfo>
  <Triggers>
    <EventTrigger>
      <Enabled>true</Enabled>
      <Subscription>&lt;QueryList&gt;&lt;Query Id="0" Path="CloudBackup"&gt;&lt;Select Path="CloudBackup"&gt;*&lt;/Select&gt;&lt;Suppress Path="CloudBackup"&gt;*[System[(EventID=17 or EventID=7 or EventID=1 or EventID=15 or EventID=6)]]&lt;/Suppress&gt;&lt;/Query&gt;&lt;/QueryList&gt;</Subscription>
    </EventTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>%USERLOGONNAME%</UserId>
      <LogonType>S4U</LogonType>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>Parallel</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>false</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>P3D</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>powershell.exe</Command>
      <Arguments>-ExecutionPolicy Bypass -File "%SCRIPTPATH%"</Arguments>
    </Exec>
  </Actions>
</Task>
"@

$TaskFullPath = Join-Path $ScriptFolder "AzureBackupResultTask.xml"
$TASKFILE = $TASKTEMPLATE -replace "%USERLOGONNAME%","$($env:USERDOMAIN)\$($env:USERNAME)" -replace "%SCRIPTPATH%",$ScriptFullPath
$TASKFILE | Out-File $TaskFullPath -Force

& schtasks /Create /TN $TaskName /XML $TaskFullPath /F

Remove-Item $TaskFullPath -Force