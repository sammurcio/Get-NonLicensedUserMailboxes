param( 	
  [switch]$EmailResults,
  [ValidateRange(0,10)]
  [int32]$CreatedDaysBack
)

$user = "admin-smurcio@Valeant.onmicrosoft.com"
$keyFile = "C:\O365 License Script\key.txt"
$pwdFile = "C:\O365 License Script\cred.txt"
$url="https://outlook.office365.com/powershell-liveid/"
$pwd = Get-Content $pwdFile | ConvertTo-SecureString -Key (Get-Content $keyFile)
$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $user, $pwd
$date = Get-Date -Format MMddyyyy
$csv = "C:\O365 License Script\Reports\" + $date + "_NonLicensed_MailboxUser_Report.csv"


function sendReport {
  $subject = "Non-Licensed User Mailbox Report $date"
  #$recipient = "o365report@valeant.com","nsylvester@westmonroepartners.com","ksullivan@westmonroepartners.com","jvancleve@westmonroepartners.com"
  $recipient = "smurcio@westmonroepartners.com", "jvancleve@westmonroepartners.com"
  $sender = "Exchange@valeant.com"
  $server = "SMTPRELAY.valeant.com"
  Send-MailMessage -SMTPServer $server -To $recipient -From $sender -Subject $subject `
  -Body "Please see the attached results file." -Attachments $csv
}

function sendError {
  $subject = "ERROR: Non-Licensed User Mailbox Report $date"
  #$recipient = "o365report@valeant.com","nsylvester@westmonroepartners.com","ksullivan@westmonroepartners.com","jvancleve@westmonroepartners.com"
  $recipient = "smurcio@westmonroepartners.com"
  $sender = "Exchange@valeant.com"
  $server = "SMTPRELAY.valeant.com"
  Send-MailMessage -SMTPServer $server -To $recipient -From $sender -Subject $subject `
  -Body "Script is unable to run due to following error, please investigate: `n $error"
}


Get-PSSession | Remove-PSSession
Import-Module MSOnline
try { Connect-MSOLService -Credential $creds -ErrorAction Stop } Catch {$error = $_; sendError; Exit }
$exoSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $url -Credential $creds `
-Authentication Basic -AllowRedirection -Name "EXO"
try { Import-PSSession $exoSession -AllowClobber -ErrorAction Stop } Catch { $error = $_; sendError; Exit }

$test = Get-MsolCompanyInformation -ErrorAction SilentlyContinue
$exoSession = Get-PSSession -Name "EXO"

if ( $test -eq $null -and $exoSession.State -ne "Opened" -and $exoSession.Availability -ne "Available" ) {
  $error = "Unable to connect to one or more of the following services:`n  -MSOL`n  -ExchangeOnline"
  sendError 
}

$activity = "Processing Request..."
$status = "Getting All Mailboxes" 
Write-Progress -Activity $activity -Status $status -ID 1
$mbxs = Get-Mailbox #-ResultSize Unlimited

if ( $createddaysback ) {
  $checkDate = (Get-Date).AddDays(-$createddaysback)
  $mbxs = $mbxs | ? { $_.WhenMailboxCreated -ge $checkDate }
}
$userMbxs = $mbxs | ? {$_.RecipientTypeDetails -eq "UserMailbox"}
Clear-Variable mbxs
Write-Progress -Activity $activity -Status $status -ID 1 -Completed

$i = 0
$out = @()
$userCount = $userMbxs.Count
foreach ( $a in $userMbxs ) {
  $i++
  Write-Progress -Activity "Processing mailbox list" -Status "Currently on mailbox $($a.DisplayName)" -PercentComplete ($i / $userCount * 100)
  $msolUser = Get-MsolUser -UserPrincipalName $a.UserPrincipalName
  if ( $msolUser.IsLicensed -eq $false) {
    $result = "User is not licensed"
  } else {
    
    if ( $msolUser.Licenses.ServiceStatus.ServicePlan.ServiceName -like "Exchange*" ) {
      $result = "User is licensed for Exchange Online"
    } else {
      $result = "User is not licensed for Exchange Online"
    }

  }

  if ( $result -match "User is not licensed" ) {
    $obj = [pscustomobject]@{
      Username = $a.UserPrincipalName
      DisplayName = $a.DisplayName
      Result = $result
      Mailbox_Creation = $a.WhenMailboxCreated
    }
    $out += $obj
  }

}

Get-PSSession | Remove-PSSession
$out | Export-Csv -NoTypeInformation $csv

if ( $emailresults ) {
  sendReport
}