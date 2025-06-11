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

# Retrieve mailboxes and handle cases where users have multiple mailboxes
$mailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited

# Track how many mailboxes each user has to flag duplicates
$mailboxCounts = @{}
foreach ($mb in $mailboxes) {
    if ($mailboxCounts.ContainsKey($mb.UserPrincipalName)) {
        $mailboxCounts[$mb.UserPrincipalName]++
    } else {
        $mailboxCounts[$mb.UserPrincipalName] = 1
    }
}

$results = foreach ($mb in $mailboxes) {
    $note = $null

    # Attempt to retrieve statistics using user principal name; if multiple
    # mailboxes exist this will throw an error which we catch silently
    try {
        $stats = Get-MailboxStatistics -Identity $mb.UserPrincipalName -ErrorAction Stop
    } catch {
        $stats = Get-MailboxStatistics -Identity $mb.Identity -ErrorAction SilentlyContinue
        $note = 'Multiple mailboxes detected'
    }

    if ($stats -and $stats.LastLogonTime -ge $cutoff) {
        [pscustomobject]@{
            EmailAddress  = $mb.PrimarySmtpAddress
            LastLogonTime = $stats.LastLogonTime
            Notes         = if ($mailboxCounts[$mb.UserPrincipalName] -gt 1) { 'Multiple mailboxes detected' } else { $note }
        }
    }
}

$results |
    Sort-Object LastLogonTime -Descending |
    Format-Table -AutoSize

Disconnect-ExchangeOnline -Confirm:$false
