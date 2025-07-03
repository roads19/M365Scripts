# Connect to Exchange Online with the following cmd if you are not connected yet
# Connect-ExchangeOnline

# List for the output data
$allRules = @()

# Gather all mailboxes
$mailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited

Write-Host "Total $($mailboxes.Count) mailboxes found. Gathering the rules..."

foreach ($mbx in $mailboxes) {
    Write-Host "Mailbox: $($mbx.UserPrincipalName)"

    try {
        # Gather the inbox rules from all mailboxes
        $rules = Get-InboxRule -Mailbox $mbx.UserPrincipalName
        foreach ($rule in $rules) {
            $allRules += [PSCustomObject]@{
                Mailbox            = $mbx.UserPrincipalName
                RuleName           = $rule.Name
                Enabled            = $rule.Enabled
                Priority           = $rule.Priority
                Description        = $rule.Description
                From               = ($rule.From | ForEach-Object { $_.Name }) -join ", "
                SentTo             = ($rule.SentTo | ForEach-Object { $_.Name }) -join ", "
                CopyToFolder       = $rule.CopyToFolder
                MoveToFolder       = $rule.MoveToFolder
                DeleteMessage      = $rule.DeleteMessage
                ForwardTo          = ($rule.ForwardTo | ForEach-Object { $_.Name }) -join ", "
                RedirectTo         = ($rule.RedirectTo | ForEach-Object { $_.Name }) -join ", "
            }
        }
    }
    catch {
        Write-Warning "Couldn't get inbox rules: $($_.Exception.Message)"
    }
}

# Export to Csv file
$csvPath = "$env:USERPROFILE\Desktop\ExchangeOnline_InboxRules.csv"
$allRules | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

Write-Host "Inbox rules have been exported to CSV file: $csvPath"
