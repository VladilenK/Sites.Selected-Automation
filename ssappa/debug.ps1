
return
#######################################
# backup
# get MS Graph API access token
$tenantName     # e.g. contoso.onmicrosoft.com
$adminAppClientId       # App (client) Id of the admin (FullControl) App
$clientSc       # App (client) secret of the admin (FullControl) App

$ReqTokenBody = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    Client_Id     = $adminAppClientId
    Client_Secret = $clientSc
} 
$TokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token" -Method POST -Body $ReqTokenBody
$AccessToken = $TokenResponse.access_token
if ($TokenResponse.expires_in -gt 0) {
    # Write-Host "Got token. Token expires at  " $(Get-date).AddMinutes($TokenResponse.expires_in)
    Write-Host "Got token. Token expires in " $TokenResponse.expires_in
}

# with Cert:
$CertificateBytes = [System.Convert]::FromBase64String($secretPlainText)
$CertificateObject = New-Object System.Security.Cryptography.x509Certificates.x509Certificate2Collection
$CertificateObject.Import($CertificateBytes,$null,[System.Security.Cryptography.x509Certificates.x509KeyStorageFlags]::Exportable)
$ProtectedCertificateBytes = $CertificateObject.Export([System.Security.Cryptography.x509Certificates.x509ContentType]::Pkcs12,"")
[System.IO.File]::WriteAllBytes("c:\temp\Certificate.pfx",$ProtectedCertificateBytes)
Import-PfxCertificate -FilePath "c:\temp\Certificate.pfx" Cert:\CurrentUser\My

$msalToken = Get-MsalToken -ClientId $adminAppClientId -ClientCertificate "c:\temp\Certificate.pfx" -TenantId $tenantId
$AccessToken = $msalToken.AccessToken

#######################################


[System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object {$_.FullName -match "Microsoft.Identity.Client"} | Select-Object -Property FullName, Location | fl

$null = [System.Reflection.Assembly]::LoadFrom("C:\Program Files\PowerShell\Modules\ExchangeOnlineManagement\3.1.0\netCore\Microsoft.Identity.Client.dll")

$DLL_MicrosoftIdentityClient = Get-ChildItem 'Microsoft.Identity.Client.dll' -Path "$home\Documents\PowerShell\Modules\" -Recurse
$LastVersion = $DLL_MicrosoftIdentityClient.VersionInfo | Where-Object { $_.FileName -match "core" } | Sort-Object FileVersion -Descending | Select-Object -First 1
[System.Reflection.Assembly]::LoadFrom("$($LastVersion.FileName)")

# Get-PnPTimeZoneId | Out-Null

