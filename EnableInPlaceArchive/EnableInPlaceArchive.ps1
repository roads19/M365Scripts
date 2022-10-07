<# Tayfun TEK
website: https://tayfuntech.com/ 
twitter: @tayfuntech
linkedin: www.linkedin.com/in/tayfuntek/
#>

Install-Module -Name ExchangeOnlineManagement -RequiredVersion 2.0.5
Set-ExecutionPolicy RemoteSigned
Import-Module ExchangeOnlineManagement

if ( Connect-ExchangeOnline )
{
    Write-Output "Exchange Online'a bağlanıldı"
}


$mailuser= "Office Apps"

if ( New-RetentionPolicyTag "Archive365Tag" -Type All -RetentionEnabled $true -AgeLimitForRetention 365 -RetentionAction MovetoArchive )
{
    Write-Output "Saklama etiketi oluşturuldu"
}
else {
    Write-Host "Saklama etiketi zaten var. Bir şey yapmanıza gerek yok"
}


if ( New-RetentionPolicy "Archive365Policy" -RetentionPolicyTagLinks "Archive365Tag" )
{
    Write-Output "Saklama ilkesi oluşturuldu"
}
else {
    Write-Host "Saklama ilkesi zaten var. Bir şey yapmanıza gerek yok"
}


if ( Set-Mailbox $mailuser -RetentionPolicy "Archive365Policy" )
{
    Write-Output "Saklama ilkesi kullanıcıya atandı"
}

function Green
{
    process { Write-Host $_ -ForegroundColor Green }
}

if ( Enable-Mailbox $mailuser -Archive )
{
    Write-Output "Arşiv Etkinleştirildi!" | Green
}

Get-Mailbox $mailuser | Format-List Name,*Archive**