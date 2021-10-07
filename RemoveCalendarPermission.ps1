$ExchangeServer=""
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://+$ExchangeServer+/Powershell/ -Authentication Kerberos
Import-PSSession $Session -AllowClobber
$User_Giving_Permissions = @("", "", "") #בעל הלוח שנה
$User_Getting_Permissions = "" #למי לתת הרשאות
foreach($User_Giving in $User_Giving_Permissions)
{
   $CalendarFolder = Get-MailboxFolderStatistics -Identity $User_Giving -FolderScope Calendar | Where-Object {$_.FolderType -eq 'Calendar'} | Select-Object Name, FolderId
   $path = "$User_Giving"+":$($calendarFolder.FolderId)" 
   Get-MailboxFolderPermission -Identity "$path"
   Remove-MailboxFolderPermission -Identity "$path"
}
$CalendarFolder = Get-MailboxFolderStatistics -Identity $User_Giving_Permissions -FolderScope Calendar | Where-Object {$_.FolderType -eq 'Calendar'} | Select-Object Name, FolderId
Get-MailboxFolderPermission -Identity "$path"  
Remove-PSSession $Session