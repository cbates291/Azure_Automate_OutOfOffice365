<#
  Set Out Of Office schedule for the IntlQuoteRequest Mailbox 
#>

$To = @("email@email.com","someone@emai.com")
$From = "email@email.com"

#Connect to Azure Automation
$Credentials = Get-AutomationPSCredential -Name 'credentials'
 
# Function: Connect to Exchange Online 
function Connect-ExchangeOnline {
    param (
        $Creds
    )
        Write-Output "Connecting to Exchange Online"
        Get-PSSession | Remove-PSSession       
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Creds -Authentication Basic -AllowRedirection
        $Commands = @("Set-MailboxAutoReplyConfiguration","Get-MailboxAutoReplyConfiguration")
        Import-PSSession -Session $Session -DisableNameChecking:$true -AllowClobber:$true -CommandName $Commands | Out-Null
    }
 
# Connect to Exchange Online
Connect-ExchangeOnline -Creds $Credentials
Connect-AzureAD -Credential $Credentials

$Error.Clear()

# Out of office message set for 6PM to 6AM following day.
# Need to account for time zones. Azure is UTC and need to have OOTO times in EST
$EST = [System.TimeZoneInfo]::FindSystemTimeZoneById("Eastern Standard Time")
$Start = Get-Date
$StartTime = [datetime]"$($Start.Month)/$($Start.Day)/$($Start.Year) 14:00:00"
$UTCStartTime = ([System.TimeZoneInfo]::ConvertTimeToUtc($StartTime, $EST))
$End = (Get-Date).AddDays(1)
$EndTime   = [datetime]"$($End.Month)/$($End.Day)/$($End.Year) 02:00:00"
$UTCEndTime = ([System.TimeZoneInfo]::ConvertTimeToUtc($EndTime, $EST))
$Mailboxes = (Get-AzureADGroupMember -ObjectId GroupObjectID).UserPrincipalName
$currentday = Get-Date -Format dddd

#weekend Times (6PM Friday to 6AM Monday)
$WeekendStartTime = [datetime]"$($Start.Month)/$($Start.Day)/$($Start.Year) 14:00:00"
$WeekendUTCStartTime = ([System.TimeZoneInfo]::ConvertTimeToUtc($WeekendStartTime, $EST))
$WeekendEnd = (Get-Date).AddDays(3)
$WeekendEndTime = [datetime]"$($WeekendEnd.Month)/$($WeekendEnd.Day)/$($WeekendEnd.Year) 02:00:00"
$WeekendUTCEndTime = ([System.TimeZoneInfo]::ConvertTimeToUtc($WeekendEndTime, $EST))

if ($currentday -match "Friday"){
foreach($Mailbox in $Mailboxes){
    Write-Host Setting OOTO for $Mailbox -ForegroundColor Yellow
    Set-MailboxAutoReplyConfiguration -Identity $Mailbox -AutoReplyState Scheduled -StartTime $WeekendUTCStartTime -EndTime $WeekendUTCEndTime -ExternalAudience All
}
}
elseif(($currentday -match "Saturday") -or ($currentday -match "Sunday")){

}
Else {
foreach($Mailbox in $Mailboxes){
    Write-Host Setting OOTO for $Mailbox -ForegroundColor Yellow
    Set-MailboxAutoReplyConfiguration -Identity $Mailbox -AutoReplyState Scheduled -StartTime $UTCStartTime -EndTime $UTCEndTime -ExternalAudience All
}
}

$Results = @()
foreach($Mailbox in $Mailboxes){
    Write-Host Setting OOTO for $Mailbox -ForegroundColor Yellow
    $Results += Get-MailboxAutoReplyConfiguration -Identity $Mailbox | Select Identity, AutoReplyState, StartTime, EndTime
}

$HTMLResults = $Results | Sort-Object Name | ConvertTo-Html -Fragment | Out-String
Send-MailMessage -From $From -To $To -Subject "Status: OOO Script" -Body $HTMLResults -BodyAsHtml -Port 587 -UseSsl -SmtpServer 'smtp.office365.com' -Credential $Credentials

if($Error){
    Write-Output "Job completed with errors."
    $ErrorLog += $Error | ForEach-Object {"$($_.Exception.Message)<br>"} | Out-String
    $Body = "Errors during script"
    $Body = "$Body $ErrorLog"
    Send-MailMessage -From $From -To $To -Subject "Error: OOO Script" -Body $Body -BodyAsHtml -Port 587 -UseSsl -SmtpServer 'smtp.office365.com' -Credential $Credentials
} else {
    Write-Output "Job completed. No Errors." 
}