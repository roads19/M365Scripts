<# Tayfun TEK
website: https://tayfuntech.com/ 
twitter: @tayfuntech
linkedin: www.linkedin.com/in/tayfuntek/
#>

#The following commands will install and import the Exchange Onlime Management Module.

Install-Module -Name ExchangeOnlineManagement -RequiredVersion 2.0.5
Set-ExecutionPolicy RemoteSigned
Import-Module ExchangeOnlineManagement

if ( Connect-ExchangeOnline )
{
    Write-Output "Connected to Exchange Online"
}

#In this section, type the "Display Name" of the O365 user

$mailuser= "Display Name"


#In this section, you can define when mailbox items move to the archive by changing the Retention Tag day count.

if ( New-RetentionPolicyTag "Archive365Tag" -Type All -RetentionEnabled $true -AgeLimitForRetention 365 -RetentionAction MovetoArchive )
{
    Write-Output "The Retention Tag has been created successfully."
}
else {
    Write-Host "The Retention Tag already exists. No action is needed."
}

#This command creates a new Retention policy which includes the Retention Tag we defined above.
if ( New-RetentionPolicy "Archive365Policy" -RetentionPolicyTagLinks "Archive365Tag" )
{
    Write-Output "The retention policy has been created successfully."
}
else {
    Write-Host "The Retention policy already exists. No action is needed."
}

#Assinging the Retention Policy to the specified user.
if ( Set-Mailbox $mailuser -RetentionPolicy "Archive365Policy" )
{
    Write-Output "Retention policy assigned to user."
}

function Green
{
    process { Write-Host $_ -ForegroundColor Green }
}

#Activating the archive
if ( Enable-Mailbox $mailuser -Archive )
{
    Write-Output "In-place archive enabled successfully.!" | Green
}

Get-Mailbox $mailuser | Format-List Name,*Archive**

#It terminates the Exchange Online session.
Disconnect-ExchangeOnline -confirm:$false