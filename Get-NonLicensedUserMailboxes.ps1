

$i = 0
$out = @()

$activity = "Processing Request..."
$status = "Getting All Mailboxes" 

Write-Progress -Activity $activity -Status $status -ID 1
$mbxs = Get-mailbox
$userMbxs = $mbxs | ? {$_.RecipientTypeDetails -eq "UserMailbox"}
Write-Progress -Activity $activity -Status $status -ID 1 -Completed

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
      $result = "User is licensed, however not for for Exchange Online"
    }

  }

  $obj = [pscustomobject]@{
    Username = $a.UserPrincipalName
    DisplayName = $a.DisplayName
    Result = $result
   }
      
   $out += $obj

}

$out | Export-Csv -NoTypeInformation "$env:USERPROFILE\Desktop\Mailboxes_NonLicensed_User_Report.csv"