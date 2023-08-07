# [START EDIT] UPDATE THESE VALUES
# Specify which days remaining will be notified.
$PasswordNotificationWindowInDays = @(31, 17, 14, 10, 5, 3, 1)
# Specify the sender email address.
$SenderEmailAddress = 'PasswordExpirationNotification@lazyexchangeadmin.cyou'
# [END EDIT]

# Get all domain password expiration policies
$domains = Get-MgDomain | Where-Object { $_.PasswordValidityPeriodInDays -ne 2147483647 } | Select-Object Id, PasswordValidityPeriodInDays
$domains | ForEach-Object { if (!$_.PasswordValidityPeriodInDays) { $_.PasswordValidityPeriodInDays = 90 } }

# Retrieve user password expiration dates
$properties = "UserPrincipalName", "mail", "displayName", "PasswordPolicies", "LastPasswordChangeDateTime", "CreatedDateTime"
$users = Get-MgUser -Filter "userType eq 'member' and accountEnabled eq true" `
    -Property $properties -CountVariable userCount `
    -ConsistencyLevel Eventual -All -PageSize 999 -Verbose | `
    Select-Object $properties | Where-Object {
    $_.PasswordPolicies -ne 'DisablePasswordExpiration' -and "$(($_.userPrincipalName).Split('@')[1])" -in $($domains.id)
}

# Add properties to the $users objects
$users | Add-Member -MemberType NoteProperty -Name Domain -Value $null
$users | Add-Member -MemberType NoteProperty -Name MaxPasswordAge -Value 0
$users | Add-Member -MemberType NoteProperty -Name PasswordAge -Value 0
$users | Add-Member -MemberType NoteProperty -Name ExpiresOn -Value (Get-Date '1970-01-01')
$users | Add-Member -MemberType NoteProperty -Name DaysRemaining -Value 0


# Get the current datetime for calculation
$timeNow = Get-Date

foreach ($user in $users) {
    # Get the user's domain
    $userDomain = ($user.userPrincipalName).Split('@')[1]
    # Get the maximum password age based on the domain password policy.
    $maxPasswordAge = ($domains | Where-Object { $_.id -eq $userDomain }).PasswordValidityPeriodInDays

    # Skip the user if the PasswordValidityPeriodInDays is 2147483647, which means no expiration.
    if ($maxPasswordAge -eq 2147483647) {
        continue;
    }

    $passwordAge = (New-TimeSpan -Start $user.LastPasswordChangeDateTime -End $timeNow).Days
    $expiresOn = (Get-Date $user.LastPasswordChangeDateTime).AddDays($maxPasswordAge)
    $user.Domain = $userDomain
    $user.MaxPasswordAge = $maxPasswordAge
    $user.PasswordAge = $passwordAge
    $user.ExpiresOn = $expiresOn
    $user.DaysRemaining = $(
        # If the remaining days is negative, show 0 instead.
        if (($daysRemaining = (New-TimeSpan -Start $timeNow -End $expiresOn).Days) -lt 1) { 0 }
        else { $daysRemaining }
    )
}

# Send Office 365 password expiration notification
foreach ($user in $users) {
    # Guard clause if the user's DaysRemaining value is not within -
    # the $PasswordNotificationWindowInDays
    # and has no email address (can't send the user an email)
    if ($user.DaysRemaining -notin $PasswordNotificationWindowInDays -or !$user.Mail) {
        continue;
    }

    # Compose the message
    $mailBody = @()
    $mailBody += '<!DOCTYPE html><html><body>'
    $mailBody += "<p>Dear, $($user.DisplayName)</p>"
    $mailBody += "<p>The password for your Office 365 account ($($user.UserPrincipalName)) will expire on <b>$(get-date $user.ExpiresOn -Format D)</b>."
    $mailBody += "<br>Please change your password soon to avoid interruption to your access.</p>"
    $mailBody += "<p>Thank you. - The IT Team</p>"

    # Create the mail object
    $mailObject = @{
        Message         = @{
            ToRecipients = @(
                @{
                    EmailAddress = @{
                        Address = $($user.Mail)
                    }
                }
            )
            Subject      = "Your Office 365 password will expire in $($user.DaysRemaining) day(s)"
            Body         = @{
                ContentType = "HTML"
                Content     = ($mailBody -join "`n")
            }
        }
        SaveToSentItems = "false"
    }

    # Send the Office 365 Password Expiration Notification Email
    try {
        "Sending password expiration notice to [$($user.displayName)] [Expires in: $($user.daysRemaining) days] [Expires on: $($user.expiresOn)]" | Out-Default
        Send-MgUserMail -BodyParameter $mailObject -UserId $SenderEmailAddress
    }
    catch {
        "There was an error sending the notification to $($user.displayName)" | Out-Default
        $_.Exception.Message | Out-Default
    }
}