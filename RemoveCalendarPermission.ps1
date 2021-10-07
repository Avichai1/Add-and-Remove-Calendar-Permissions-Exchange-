$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exch33.maof-group.co.il/Powershell/ -Authentication Kerberos
Import-PSSession $Session -AllowClobber
$User_Giving_Permissions = @('yaronl@maof-group.co.il', "sapirbz@maof-group.co.il", "nasty@maof-group.co.il") #בעל הלוח שנה
$User_Getting_Permissions = 'avihai@maof-group.co.il' #למי לתת הרשאות
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