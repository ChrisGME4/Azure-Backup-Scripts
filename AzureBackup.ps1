<#
Azure Online Backup Report via Mail
This script will compile a report of the Azure Backup jobs from the last 24 hours.
#>

# Sets Company name
$Company = "My IT"
 
# Sets the recipient/sender email-address
$MailTo = "noreply@test.com"
$MailFrom = "noreply@test.com"

# User account and password for logging into Exchange
$MailUser = "noreply@test.com"
$MailPassword = "password" 

# Points to Exchange server
$MailServer = "smtp.office365.com"
 
# Sets SMTP port
$MailPort = 587
 
# If server uses SSL = true if not = $false
$UseSSL = $true


$Computer = Hostname


Try {

Function FormatBytes {
	Param(
		[System.Int64]$Bytes
	)
	[string]$BigBytes = ""
	#Convert to TB
	If ($Bytes -ge 1TB) {$BigBytes = [math]::round($Bytes / 1TB, 2); $BigBytes += " TB"}
	#Convert to GB
	ElseIf ($Bytes -ge 1GB) {$BigBytes = [math]::round($Bytes / 1GB, 2); $BigBytes += " GB"}
	#Convert to MB
	ElseIf ($Bytes -ge 1MB) {$BigBytes = [math]::round($Bytes / 1MB, 2); $BigBytes += " MB"}
	#Convert to KB
	ElseIf ($Bytes -ge 1KB) {$BigBytes = [math]::round($Bytes / 1KB, 2); $BigBytes += " KB"}
	#If smaller than 1KB, leave at bytes.
	Else {$BigBytes = $Bytes; $BigBytes += " Bytes"}
	Return $BigBytes
}

Function Log-BackupItems {
    Param(
        [System.String]$Name,
        [System.String]$Status,
        [System.String]$Start,
        [System.String]$End,
        [System.Int64]$Upload,
        [System.Int64]$Size
    )
    $Item = New-Object System.Object;
    $Item | Add-Member -Type NoteProperty -Name "Name" -Value $Name;
    $Item | Add-Member -Type NoteProperty -Name "Status" -Value $Status;
    $Item | Add-Member -Type NoteProperty -Name "Start Time" -Value $Start;
    $Item | Add-Member -Type NoteProperty -Name "End Time" -Value $End;
    $Item | Add-Member -Type NoteProperty -Name "Uploaded" -Value (FormatBytes -Bytes $Upload);
    $Item | Add-Member -Type NoteProperty -Name "Total Size" -Value (FormatBytes -Bytes $Size);
    Return $Item;
}

Import-module Azure

$Password = ConvertTo-SecureString $MailPassword -AsPlainText -Force
$Credentials = New-Object System.Management.Automation.PSCredential ($MailUser, $Password)

$CurrentTime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
$OBScope = (Get-Date).AddDays(-1)
$OBJob = (Get-OBJob -Previous 1).JobStatus | Sort StartTime | Where { $_.StartTime -gt $OBScope }
$env:computername
If ($OBJob.JobState -contains "Failed") { $OBResult = "Failed" }
ElseIf ($OBJob.JobState -contains "Aborted") { $OBResult = "Failed" }
ElseIf ($OBJob.JobState -contains "Completed") { $OBResult = "Normal" }
ElseIf ($OBJob.JobState -eq $null) { $OBResult = "Failed" }
Else { $OBResult = "Failed" }

$results=@()
If ($OBJob.JobState -ne $null) {
$OBJob | % {
$count = 0
foreach($obj in $_.DatasourceStatus.Datasource)
{
$BackupItem = $null
$OBStartTime = $_.StartTime.AddHours(-6).ToString("yyyy-MM-dd HH:mm:ss")
$OBEndTime = $_.EndTime.AddHours(-6).ToString("yyyy-MM-dd HH:mm:ss")
$BackupItem = Log-BackupItems -Start $OBStartTime -End $OBEndTime -Name $obj.DataSourceName -Status $_.Jobstate -Changed $_.DatasourceStatus.ByteProgress[$count].Changed -Bytes $_.DatasourceStatus.ByteProgress[$count].Total
$BackupItem
$results += $BackupItem
$count += 1
        }
    }
}
Else { $results = Log-BackupItems -Start "N/A" -End "N/A" -Name "N/A" -Status "N/A" -Upload 0 -Size 0 }

# Assemble the HTML Report
$HTMLMessage = @"
Device: $computer
Status: $OBResult

<title>$Company Azure Backup Report</title>
    <style>
    body { font-family: Verdana, Arial, sans-serif; font-size: 12px }
    h3{ clear: both; font-size: 1.5em; margin-left: 50px;margin-top: 30px; text-align:center }
    table { padding: 15px 0 20px; text-align: left; width:800px }
    td, th { padding: 0 20px 0 0; margin 0; text-align: left; }
    th { margin-top: 15px }
    tr { margin-top: 5px }
    a, a:visited { color: #2ea3f2; text-decoration: none; }
    .completed { color: green }
    .aborted, .missing { color: orange }
    .failed { color: red }
    </style>
    <div align="center">
    <table><tbody>
    <tr><td><h3><a>$computer Azure Backup Report</a></h3></td></tr>
    <tr><td>Backup Result: <b class="$OBResult">$OBResult</b></td></tr>
    </tbody></table>
    $(
	    $html = $results | ConvertTo-HTML -Fragment
	    $xml=[xml]$html
	    $attr=$xml.CreateAttribute('id')
	    $attr.Value='items'
	    $xml.table.Attributes.Append($attr) | out-null
	    $html=$xml.OuterXml | out-string
	    $html
    )
    </div>
    
 

"@

$email = @{
    SMTPServer = $MailServer
    UseSSL = $UseSSL
    BodyAsHtml = $true
    Port = $MailPort
    Credential = $Credentials
    Encoding = ([System.Text.Encoding]::UTF8)
    To = $MailTo
    From = $MailFrom
    Subject = "[Microsoft Azure Backup Report]"
    Body = $HTMLMessage
    }
    Send-MailMessage @email
}

Catch {
    $email = @{
    SMTPServer = $MailServer
    UseSSL = $UseSSL
    BodyAsHtml = $true
    Port = $MailPort
    Credential = $Credentials
    Encoding = ([System.Text.Encoding]::UTF8)
    To = $MailTo
    From = $MailFrom
    Subject = "Azure Backup Failed."
    Body = "The backup script failed to run!"
    }
    Send-MailMessage @email
}
