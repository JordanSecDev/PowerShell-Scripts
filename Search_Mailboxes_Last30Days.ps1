# Connect to Exchange Online using MFA
# Requires Exchange Online Management module
Param(
    [string]$UserPrincipalName
)

if (-not $UserPrincipalName) {
    Write-Host "Enter the user principal name for the account used to connect." -ForegroundColor Yellow
    $UserPrincipalName = Read-Host "User Principal Name"
}

# Establish the connection with MFA prompt
Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName

# Determine cutoff date
$cutoff = (Get-Date).AddDays(-30)

# Retrieve mailbox statistics and filter by last logon time
$mailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited |
    Get-MailboxStatistics |
    Where-Object { $_.LastLogonTime -ge $cutoff }

$mailboxes | Select-Object @{N='EmailAddress';E={$_.PrimarySmtpAddress}},
                         @{N='LastLogonTime';E={$_.LastLogonTime}} |
    Sort-Object LastLogonTime -Descending |
    Format-Table -AutoSize

Disconnect-ExchangeOnline -Confirm:$false
