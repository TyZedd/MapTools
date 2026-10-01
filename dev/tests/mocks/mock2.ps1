. ./mock.ps1
$E3='05e9a617-0261-4cee-bb44-138d3ef5d965'; $PBI='f8a1db68-be16-40ed-86d5-cb42ce701560'
foreach($usr in $global:U){ $usr | Add-Member -NotePropertyName AssignedLicenses -NotePropertyValue @() -Force }
$global:U[0].AssignedLicenses=@([pscustomobject]@{SkuId=$E3},[pscustomobject]@{SkuId=$PBI})
$global:U[1].AssignedLicenses=@([pscustomobject]@{SkuId=$E3})
$global:U[3].AssignedLicenses=@([pscustomobject]@{SkuId=$E3},[pscustomobject]@{SkuId='11111111-2222-3333-4444-555555555555'})
function Get-MgSubscribedSku { param([switch]$All,$ErrorAction) @([pscustomobject]@{SkuId=$E3;SkuPartNumber='SPE_E3'},[pscustomobject]@{SkuId=$PBI;SkuPartNumber='POWER_BI_PRO'}) }
function Get-MgReportAuthenticationMethodUserRegistrationDetail { param([switch]$All,$Property,$ErrorAction)
  @([pscustomobject]@{Id='u1';IsMfaRegistered=$true;IsAdmin=$false;MethodsRegistered=@('microsoftAuthenticatorPush','softwareOneTimePasscode')},
    [pscustomobject]@{Id='u2';IsMfaRegistered=$false;IsAdmin=$false;MethodsRegistered=@()},
    [pscustomobject]@{Id='u3';IsMfaRegistered=$false;IsAdmin=$true;MethodsRegistered=@()}) }
