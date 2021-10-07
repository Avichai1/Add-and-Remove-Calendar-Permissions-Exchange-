$ExchangeServer=""
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://+$ExchangeServer+/Powershell/ -Authentication Kerberos
Import-PSSession $Session -AllowClobber
$User_Giving_Permissions = 'anniel@maof-group.co.il' #בעל הלוח שנה
$User_Getting_Permissions = 'annaa@maof-group.co.il' #למי לתת הרשאות
$CalendarFolder = Get-MailboxFolderStatistics -Identity $User_Giving_Permissions -FolderScope Calendar | Where-Object {$_.FolderType -eq 'Calendar'} | Select-Object Name, FolderId
$path = "$User_Giving_Permissions"+":$($calendarFolder.FolderId)" 
Get-MailboxFolderPermission -Identity "$path"
Add-MailboxFolderPermission -Identity "$path" -User $User_Getting_Permissions -AccessRights Author
Set-MailboxFolderPermission -Identity "$path" -User $User_Getting_Permissions -AccessRights Author
$CalendarFolder = Get-MailboxFolderStatistics -Identity $User_Giving_Permissions -FolderScope Calendar | Where-Object {$_.FolderType -eq 'Calendar'} | Select-Object Name, FolderId
Get-MailboxFolderPermission -Identity "$path"  
Remove-PSSession $Session