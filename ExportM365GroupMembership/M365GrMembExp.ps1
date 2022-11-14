<# Tayfun TEK
website: https://tayfuntech.com/ 
twitter: @tayfuntech
linkedin: www.linkedin.com/in/tayfuntek/
#>

function Green
{
    process { Write-Host $_ -ForegroundColor Green }
}

if ( New-Item "C:\ExportM365Folder" -itemType Directory)
{ Write-Output "The target folder (C:\ExportM365Folder) has been created successfully" | Green 
}

#Connect Exchange Online Powershell
Connect-ExchangeOnline

#Define all M365 groups to $Groups variable 
$Groups = Get-UnifiedGroup -ResultSize Unlimited

#This part will list group memberships and some attributes of members.
$Groups | ForEach-Object {
$group = $_
Get-UnifiedGroupLinks -Identity $group.Name -LinkType Members -ResultSize Unlimited | ForEach-Object {
New-Object -TypeName PSObject -Property @{
Group = $group.DisplayName
Member = $_.Name
EmailAddress = $_.PrimarySMTPAddress
RecipientType= $_.RecipientType
}}} | Export-CSV "C:\ExportM365Folder\M365GroupMembers.csv" -NoTypeInformation -Encoding UTF8 

#Check if the export file has been created successfully.
$Folder = "C:\ExportM365Folder\M365GroupMembers.csv" 
if (Test-Path -Path $Folder) { "The export file has been created! C:\ExportM365Folder\M365GroupMembers.csv"| Green } else {
    "The export file has not been created." 
}
