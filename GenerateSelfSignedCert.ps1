# Specify the app or certificate name
$CertificateOrAppName = 'Office 365 Password Expiration Notification'

# Generate a self-signed certificate
$certSplat = @{
    Subject           = $CertificateOrAppName
    NotBefore         = ((Get-Date).AddDays(-1))
    NotAfter          = ((Get-Date).AddYears(3))
    CertStoreLocation = "Cert:\CurrentUser\My"
    Provider          = "Microsoft Enhanced RSA and AES Cryptographic Provider"
    HashAlgorithm     = "SHA256"
    KeySpec           = "KeyExchange"
    KeyExportPolicy   = "Exportable"
}
$selfSignedCertificate = New-SelfSignedCertificate @certSplat

# Export the public certificate.
$selfSignedCertificate | Export-Certificate -FilePath ".\$CertificateOrAppName.cer"


# Display the certificate details
# $selfSignedCertificate | Format-List PSParentPath, ThumbPrint, Subject, NotAfter

# Export the certificate to PFX.
# $selfSignedCertificate | Export-PfxCertificate -FilePath ".\$CertificateOrAppName.pfx" -Password $(ConvertTo-SecureString -String "Transpire^Struck8^Impurity^Draw" -AsPlainText -Force)

