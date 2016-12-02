param( 	
  [switch]$EmailResults,
  [ValidateRange(0,10)]
  [int32]$CreatedDaysBack
)

# Variables that can be modified:
#=================================================================================
$csv = "C:\O365\Reports\" + $date + "_NonLicensed_MailboxUser_Report.csv"
$checkLic = "WESTMONROEPARTNERS1:ENTERPRISEPREMIUM_NOPSTNCONF"
#Variables if email results is enabled
$recipient = "smurcio@domain.com", "user@domain.com"
$sender = "Exchange@domain.com"
$server = "SMTPRELAY.domain.com"
#=================================================================================

$user = "asfa@domain.onmicrosoft.com"
$keyFile = "C:\O365 License Script\key.txt"
$pwdFile = "C:\O365 License Script\cred.txt"
$url="https://outlook.office365.com/powershell-liveid/"
$pwd = Get-Content $pwdFile | ConvertTo-SecureString -Key (Get-Content $keyFile)
$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $user, $pwd
$date = Get-Date -Format MMddyyyy

function sendReport {
  $subject = "Non-Licensed User Mailbox Report $date"
  Send-MailMessage -SMTPServer $server -To $recipient -From $sender -Subject $subject `
  -Body "Please see the attached results file." -Attachments $csv
}

function sendError {
  $subject = "ERROR: Non-Licensed User Mailbox Report $date"
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
$status = "Getting All User Mailboxes" 
Write-Progress -Activity $activity -Status $status -ID 1
$mbxs = Get-Mailbox -Filter "RecipientTypeDetails -eq 'UserMailbox'" -ResultSize Unlimited
if ( $createddaysback ) {
  $checkDate = (Get-Date).AddDays(-$createddaysback)
  $mbxs = $mbxs | ? { $_.WhenMailboxCreated -ge $checkDate }
}
Write-Progress -Activity $activity -Status $status -ID 1 -Completed

$i = 0
$mbxCount = $mbxs.Count
foreach ( $a in $mbxs ) {
  $i++
  if ( $msolUser ) { Clear-Variable msolUser }
  $upn = $a.UserPrincipalName
  $displayName = $a. displayName
  $created = $a.WhenMailboxCreated
  Write-Progress -Activity "Processing mailbox list" -Status "Currently on mailbox $displayName" -PercentComplete ($i / $mbxCount * 100)
  $msolUser = Get-MsolUser -UserPrincipalName $upn -ErrorAction SilentlyContinue

  if ( !$msolUser ) {
    $result = "MSOL User Not Found"
  } elseif ( $msolUser.IsLicensed -eq $false) {
    $result = "User is not licensed"
  } elseif ( $msolUser.IsLicensed -eq $true ) {

    if ( $msolUser.Licenses.AccountSkuId -contains $checkLic ) {
      $serviceStatus = $msolUser.Licenses |`
      ? { $_.AccountSkuid -eq $checkLic } | Select -ExpandProperty ServiceStatus |`
      ? {$_.serviceplan.servicename -like "EXCHANGE*" -and $_.serviceplan.servicename -notlike "EXCHANGE_ANALYTICS"}
      $isDisabled = $serviceStatus.ProvisioningStatus -contains "Disabled"
      
      if ( $isDisabled -eq $true ) {
        $result = "User is licensed with SKU $($checkLic) but service is not enabled"
      } else {
         $result = "User is licensed with SKU $($checkLic) and Exchange Online service is enabled"
      }

    } else {
      $result = "User is licensed but not with SKU $($checkLic)"
    }

  }

  $obj = [pscustomobject]@{
    Username = $upn
    DisplayName = $displayName
    Result = $result
    Mailbox_Created = $created
    }
    $obj | Export-Csv -Path $csv -NoTypeInformation -Append

}

Get-PSSession | Remove-PSSession

if ( $emailresults ) {
  sendReport
}
