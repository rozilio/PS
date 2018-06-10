ipmo *azure*

$cred = Get-Credential
Connect-AzureRmAccount -Subscription f030c0c1-8cc5-4ec7-bd1e-977cd5162255 -Credential $cred

#Get-AzureRmADUser | Out-file AzureUsers.txt
#Get-AzureRmADUser | ? {$_.UserPrincipalName -like "*EXT*" -and $_.UserPrincipalName -notlike "mscloud.ofek_outlook.com*" -and $_.UserPrincipalName -notlike "msbluesky_outlook.com*"} | Out-File AzureGuestUsers.txt

$MNN = "lucasa"
$UPN = $MNN + "@mscloudofekoutlook.onmicrosoft.com"
$DN = "lucas aides"
$pass = ConvertTo-SecureString -String "Aa123456" -AsPlainText -Force
System.secu
New-AzureRmADUser -UserPrincipalName $UPN -DisplayName $DN -MailNickname $MNN -Password $pass -ForceChangePasswordNextLogin

Get-AzureRmADGroup | select DisplayName
$Group = Get-AzureRmADGroup -DisplayName
Add-AzureRmADGroupMember -MemberUserPrincipalName $UPN -TargetGroupDisplayName $Group