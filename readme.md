# Office 365 Password Expiration Notification

## Step 1: Connect to Microsoft Graph PowerShell

You must first connect to Microsoft Graph.

* *Using Delegated Access*

  ```PowerShell
  Connect-MgGraph -Scopes 'User.Read.All', 'Mail.Send', 'Domain.Read.All'
  ```

* *Using App-Only Access with Certificate*

  ```PowerShell
  Connect-MgGraph -ClientId <client ID> -TenantId <tenant ID> -CertificateThumbprint <thumbprint>
  ```

## Step 2: Update the Password Expiration Threshold and Sender Email in the Script

Open the script in the editor and modify these values.

> **Note**: If you're using *Delegated Access*, the **`$SenderEmailAddress`** must be the logged in user's email address.

```PowerShell
$PasswordNotificationWindowInDays = @(31, 17, 14, 10, 5, 3, 1)
$SenderEmailAddress = 'EMAIL ADDRESS HERE'
```

## Step 3: Run the Script

Run the script. No parameters required.

```PowerShell
.\Invoke-PasswordExpirationNotification.ps1
```
