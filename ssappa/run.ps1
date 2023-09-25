# Input bindings are passed in via param block.
param($Timer)
$currentUTCtime = (Get-Date).ToUniversalTime()
if ($Timer.IsPastDue) { Write-Host "PowerShell timer is running late!"}
Write-Host "PowerShell timer trigger function ran! TIME: $currentUTCtime"
Write-Host "=================================================================="

Get-Module -Name PnP.PowerShell -ListAvailable
Get-Module -Name Az -ListAvailable
# Import-Module -Name PnP.PowerShell

$VaultName = "SharePointAutomationDemo"
$certName = "SharePointAutomationAppCert" # ...
Get-AzKeyVault -VaultName $vaultName 
$secretSecureString = Get-AzKeyVaultSecret -VaultName $vaultName -Name $certName 
$secretPlainText = ConvertFrom-SecureString -AsPlainText -SecureString $secretSecureString.SecretValue
$secretPlainText.Substring(0, 4)

$orgName = "s5dz3"
$tenantName = "$orgName.onmicrosoft.com"
$adminUrl = "https://$orgName-admin.sharepoint.com"
$tenantId = '7ddc7314-9f01-45d5-b012-71665bb1c544'
$adminAppClientId = '7e60c372-ec15-4069-a4df-0ab47912da46'

$connAdmin = Connect-PnPOnline -Url $adminUrl -ClientId $adminAppClientId -CertificateBase64Encoded $secretPlainText -Tenant $tenantId -ReturnConnection 
Get-PnPTenant -Connection $connAdmin 
Get-PnPTenantSite -Url $adminUrl -Connection $connAdmin | ft -a
Get-PnPSite -Connection $connAdmin | ft -a

$intakeSiteUrl = "https://s5dz3.sharepoint.com/sites/spoa"
$intakeListId  = "4c488eca-fd1a-4c42-8b5c-7d5979cc0c24"
$connIntakeSite = Connect-PnPOnline -Url $intakeSiteurl -ClientId $adminAppClientId -CertificateBase64Encoded $secretPlainText -Tenant $tenantId -ReturnConnection 
$connIntakeSite.Url

# Get-PnPList -Connection $connIntakeSite # to get intake list Id
# $intakeList = Get-PnPList -Connection $connIntakeSite -Identity $intakeListId -Includes Fields
# $intakeList.Fields | clip
$intakeListItems = @()
$intakeListItems += Get-PnPListItem -Connection $connIntakeSite -List $intakeListId 
$intakeListItems = $intakeListItems | ?{!$_.AutomationOutput}
$intakeListItems.Count

$intakeListItem = $intakeListItems[-1]; $intakeListItem
$intakeListItem.FieldValues['Title']
$intakeListItem.FieldValues['AppId']
$intakeListItem.FieldValues['Permissions']
$intakeListItem.FieldValues['AutomationOutput']

# providing permissions
foreach ($intakeListItem in $intakeListItems) {
  Write-Host "Intake list item: " $intakeListItem.Id $intakeListItem.FieldValues['Title'], $intakeListItem.FieldValues['AppId']

  $clientAppId = $intakeListItem.FieldValues['AppId']
  $clientApp = Get-PnPAzureADApp -Identity $clientAppId -Connection $connAdmin
  $clientAppName = $clientApp.DisplayName

  $request = $intakeListItem.FieldValues['Permissions'].ToLower()
  $clientSiteUrl = $intakeListItem.FieldValues['Title']
  
  $connClientSite = Connect-PnPOnline -Url $clientSiteUrl -ClientId $adminAppClientId -CertificateBase64Encoded $secretPlainText -Tenant $tenantId -ReturnConnection  
  $clientSite = Get-PnPSite -Connection $connClientSite -Includes Id
  $clientSiteId = $clientSite.Id.Guid

  if ($request -match "report") {
    Write-Host "Getting all site permissions"  
    $rawPermissions = Get-PnPAzureADAppSitePermission -Connection $connAdmin -Site $clientSiteId 
    $permissionsText = "Application permissions to site (AppId:permissions):"
    foreach ($rawPermission in $rawPermissions) {
      $permissionsText += "`n"
      $permissionsText += ($rawPermission.Apps.id -join ",") + ":"
      $permissionsText += $rawPermission.Roles -join "," 
    }
    $permissionsText
    $values = @{AutomationOutput = $permissionsText}
    Set-PnPListItem -List $intakeListId -Identity $intakeListItem.Id -Values $values -Connection $connIntakeSite
  }

  if ($request -match "provision") {
    Write-Host "Adding app permissions to site"  
    $role = "Read"
    if ($intakeListItem.FieldValues['Permissions'].ToLower() -match "write") {
        $role = "Write"
    }

    $rawPermission = Grant-PnPAzureADAppSitePermission -AppId $clientAppId -DisplayName $clientAppName -Permissions $role -Site $clientSiteId -Connection $connAdmin
    $permission = [PSCustomObject]@{
      PermissionId      = $rawPermission.Id 
      AppRole    = $rawPermission.roles -join ";"
      AppName = $rawPermission.Apps.DisplayName -join ";"
      AppId   = $rawPermission.Apps.id -join ";"
      SiteId   = $clientSiteId
      SiteUrl  = $clientSite.Url
    }
    $permission | fl
    $values = @{AutomationOutput = "Permissions provided:`n" + $permission}
    Set-PnPListItem -List $intakeListId -Identity $intakeListItem.Id -Values $values -Connection $connIntakeSite
  }

  if ($request -match "remove") {
    Write-Host "Removing app permissions to site"  
    $rawPermissions = Get-PnPAzureADAppSitePermission -Connection $connAdmin -Site $clientSiteId 
    $rawPermission = $rawPermissions | ?{$_.Apps.Id -eq $clientAppId}
    Revoke-PnPAzureADAppSitePermission -PermissionId $rawPermission.Id -Site $clientSiteId -Connection $connAdmin -Force
    $permission = [PSCustomObject]@{
        Role    = $rawPermission.roles -join ";"
        AppName = $rawPermission.Apps.DisplayName -join ";"
        AppId   = $rawPermission.Apps.id -join ";"
    }
    $permission
    $values = @{AutomationOutput = "Permissions removed:`n" + $permission}
    Set-PnPListItem -List $intakeListId -Identity $intakeListItem.Id -Values $values -Connection $connIntakeSite
  }

}

# DisConnect-AzAccount 
# Connect-AzAccount 
