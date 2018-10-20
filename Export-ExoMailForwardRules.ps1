<#	
	.NOTES
	===========================================================================
	 Created on:   	19-October-2018
	 Created by:   	Tito Castillote
					june.castillote@gmail.com
	 Filename:     	Export-ExoMailForwardRules.ps1
	 Version:		1.0 (19-October-2018)
	===========================================================================

	.LINK
        https://www.lazyexchangeadmin.com/2018/10/office-365-mailbox-forwarding-rules.html
        https://github.com/junecastillote/Export-ExoMailForwardRules

	.SYNOPSIS
        This script can be used to export the mailbox forwarding rules to a csv file
        for security audit and reporting.

	.DESCRIPTION
		For more details and usage instruction, please visit the link:
        https://www.lazyexchangeadmin.com/2018/10/office-365-mailbox-forwarding-rules.html
        https://github.com/junecastillote/Export-ExoMailForwardRules
		
	.EXAMPLE
		.\Export-ExoMailForwardRules.ps1

#>
$scriptVersion = "1.0"

#get root path of the script
$script_root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$WarningPreference = "SilentlyContinue"

#create folders if they don'e exist
if ((Test-Path "$($script_root)\Reports") -eq $false)
{
    New-Item -ItemType Directory "$($script_root)\Reports" | out-null
}

if ((Test-Path "$($script_root)\Logs") -eq $false)
{
    New-Item -ItemType Directory "$($script_root)\Logs" | out-null
}

#kill transcript if still running
try{
    stop-transcript|out-null
  }
  catch [System.InvalidOperationException]{}

#start transcribing
Start-Transcript -Path "$($script_root)\Logs\ScriptRun_$((get-date).tostring("yyyy_MM_dd")).txt"

Function New-EXOSession()
{
    param([parameter(mandatory=$true)]$exoCredential)

    #discard all current PSSession
    Get-PSSession | Remove-PSSession -Confirm:$false

    #create new Exchange Online Session
    $EXOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid -Credential $exoCredential -Authentication Basic -AllowRedirection
    $office365 = Import-PSSession $EXOSession -DisableNameChecking -AllowClobber
}



#<O365 CREDENTIALS
#Note: This uses an encrypted credential (XML). To store the credential:
#1. Login to the Server/Computer using the account that will be used to run the script/task
#2. Run this "Get-Credential | Export-CliXml Office365StoredCredential.xml"
#3. Make sure that Office365StoredCredential.xml is in the same folder as the script.
$onLineCredential = Import-Clixml "$($script_root)\Office365StoredCredential.xml"
#O365 CREDENTIALS>

#<mail variables
$sendEmail = $true
$sender = "healthmonitor@lazyexchangeadmin.com"
$recipients = "june.castillote@lazyexchangeadmin.com","june.castillote@gmail.com"
$subject = "Office 365 Mailbox Forwarding Rule Report"
$smtpServer = "smtp.office365.com"
$smtpPort = "587"
#mail variables>

#open new Exchange Online Session
Write-Host "Connecting Session" -ForegroundColor Green
New-EXOSession $onlineCredential

$outputFile = $script_root +"\Reports\ForwardRules_$((get-date).tostring("yyyy_MM_dd")).csv"

if (Test-Path $outputFile)
{
    Remove-Item $outputFile -Force -Confirm:$false
}

#retrieve all mailbox list
Write-Host "Retrieving Mailbox List" -ForegroundColor Green
$mailboxes = get-mailbox -RecipientTypeDetails UserMailbox -ResultSize 50 | Sort-Object UserPrincipalName | Select-Object UserPrincipalName,PrimarySMTPAddress,SamAccountName,DisplayName
$mailboxes | Export-Csv -NoTypeInformation "$($script_root)\Reports\mailboxes.csv"
$totalMailbox = $mailboxes.Count

#paging variable
$pageSize=100
$offSet=0

$i=1
$totalRules=0
do {
	$InboxRules = @()
	$mailboxList=@(import-csv "$($script_root)\Reports\mailboxes.csv" | select-object -skip $offset -First $pageSize)

    foreach ($mailbox in $mailboxList){

        Write-host "$($i) of $($totalMailbox) | $($mailbox.UserPrincipalname) | " -ForegroundColor Yellow -NoNewline
        $i++
        $rules = Get-InboxRule -Mailbox $mailbox.UserPrincipalname | Select-Object Name, Description, Enabled, Priority, ForwardTo, ForwardAsAttachmentTo, RedirectTo, DeleteMessage | Where-Object {($_.ForwardTo -ne $null) -or ($_.ForwardAsAttachmentTo -ne $null) -or ($_.RedirectsTo -ne $null)}
        $ruleCount = $rules | Measure-Object
        $totalRules=$totalRules+$ruleCount.count
        if ($rules){
            Write-Host "Found $($ruleCount.Count) rules" -ForegroundColor Green
            foreach ($rule in $rules){
                $temp = "" | Select-Object UserPrincipalName, DisplayName, RuleName, Description, Enabled, Priority, ForwardTo, ForwardAsAttachmentTo, RedirectTo, DeleteMessage
                $temp.UserPrincipalName = $mailbox.UserPrincipalName
                $temp.DisplayName = $mailbox.DisplayName
                $temp.RuleName = $rule.Name
                $temp.Description = $rule.Description
                $temp.Enabled = $rule.Enabled
                $temp.Priority = $rule.Priority
                $temp.ForwardTo = $rule.ForwardTo
                $temp.ForwardAsAttachmentTo = $rule.ForwardAsAttachmentTo
                $temp.RedirectTo = $rule.RedirectTo
                $temp.DeleteMessage = $rule.DeleteMessage
                $InboxRules += $temp
            }
        }
        else {
            Write-Host "Found 0 rules" -ForegroundColor Cyan
        }
    }

Write-Host "Total Rules Found: $($totalRules)" -ForegroundColor Green
$InboxRules | Export-Csv -NoTypeInformation $outputFile -Append

#increment the offset value
$offSet+=$pageSize
	
#call the new session function

if ($offset -lt $totalMailbox)
{
    Write-Host "Reconnecting Session" -ForegroundColor Green
    New-EXOSession $onlineCredential
}

} while($offset -lt $totalMailbox)

if ($sendEmail -eq $true -and $totalRules -gt 0 ){
    Write-Host "Sending Email" -ForegroundColor Green
    Send-MailMessage -SmtpServer $smtpServer -Port $smtpPort -From $sender -To $recipients -Subject $subject -Body "<a href=""https://github.com/junecastillote/Export-ExoMailForwardRules"">Export-ExoMailForwardRules v$($scriptVersion)</a>" -Attachments $outputFile -BodyAsHtml -Credential $onLineCredential -UseSsl
}

Stop-Transcript