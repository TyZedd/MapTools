function Get-MgContext { [pscustomobject]@{ TenantId='00000000-test'; Account='paj-a2@contoso.ch'; Scopes=@('GroupMember.Read.All','User.Read.All','AuditLog.Read.All','LicenseAssignment.Read.All','RoleManagement.Read.Directory') } }
function Connect-MgGraph {}; function Disconnect-MgGraph {}
$now=Get-Date
function SIA($succ,$int){ [pscustomobject]@{ LastSuccessfulSignInDateTime=$succ; LastSignInDateTime=$int; LastNonInteractiveSignInDateTime=$null } }
$global:U=@(
 [pscustomobject]@{Id='u1';DisplayName='Marie Dupont';UserPrincipalName='marie.dupont@contoso.ch';AccountEnabled=$true;UserType='Member';CreatedDateTime=$now.AddYears(-3);OnPremisesSyncEnabled=$true;SignInActivity=(SIA $now.AddDays(-3) $now.AddDays(-3))}
 [pscustomobject]@{Id='u2';DisplayName='Luc Ancien';UserPrincipalName='luc.ancien@contoso.ch';AccountEnabled=$true;UserType='Member';CreatedDateTime=$now.AddYears(-3);OnPremisesSyncEnabled=$true;SignInActivity=(SIA $null $now.AddDays(-400))}
 [pscustomobject]@{Id='u3';DisplayName='Jean (admin)';UserPrincipalName='jdo-a2@contoso.ch';AccountEnabled=$true;UserType='Member';CreatedDateTime=$now.AddYears(-2);OnPremisesSyncEnabled=$false;SignInActivity=(SIA $null $null)}
 [pscustomobject]@{Id='u4';DisplayName='Parti';UserPrincipalName='parti@contoso.ch';AccountEnabled=$false;UserType='Member';CreatedDateTime=$now.AddYears(-2);OnPremisesSyncEnabled=$true;SignInActivity=$null}
 [pscustomobject]@{Id='u5';DisplayName='Stagiaire';UserPrincipalName='stag@contoso.ch';AccountEnabled=$true;UserType='Member';CreatedDateTime=$now.AddDays(-20);OnPremisesSyncEnabled=$true;SignInActivity=(SIA $null $null)}
 [pscustomobject]@{Id='u6';DisplayName='Guest Gmail';UserPrincipalName='g_gmail.com#EXT#@contoso.onmicrosoft.com';AccountEnabled=$true;UserType='Guest';CreatedDateTime=$now.AddYears(-1);OnPremisesSyncEnabled=$null;SignInActivity=(SIA $null $null)}
)
function Get-MgUser { param([switch]$All,$Property,$ErrorAction) $global:U }
$global:G=@(
 [pscustomobject]@{Id='g1aaaaaaaa';DisplayName='GG_Finance';GroupTypes=@();SecurityEnabled=$true;MailEnabled=$false;OnPremisesSyncEnabled=$true;CreatedDateTime=$now.AddYears(-5)}
 [pscustomobject]@{Id='g2bbbbbbbb';DisplayName='GG_Compta';GroupTypes=@();SecurityEnabled=$true;MailEnabled=$false;OnPremisesSyncEnabled=$true;CreatedDateTime=$now.AddYears(-5)}
 [pscustomobject]@{Id='g3cccccccc';DisplayName='Equipe Finance';GroupTypes=@('Unified');SecurityEnabled=$false;MailEnabled=$true;OnPremisesSyncEnabled=$null;CreatedDateTime=$now.AddYears(-2)}
 [pscustomobject]@{Id='g4dddddddd';DisplayName='SG_Dyn';GroupTypes=@('DynamicMembership');SecurityEnabled=$true;MailEnabled=$false;OnPremisesSyncEnabled=$null;CreatedDateTime=$now.AddYears(-1)}
 [pscustomobject]@{Id='g5eeeeeeee';DisplayName='GG_Compta';GroupTypes=@();SecurityEnabled=$true;MailEnabled=$false;OnPremisesSyncEnabled=$false;CreatedDateTime=$now.AddYears(-1)}
)
function M($id,$type,$extra){ $ap=@{'@odata.type'="#microsoft.graph.$type"}; if($extra){$extra.GetEnumerator()|%{$ap[$_.Key]=$_.Value}}; [pscustomobject]@{Id=$id;AdditionalProperties=$ap} }
$global:MB=@{
 'g1aaaaaaaa'=@((M 'g2bbbbbbbb' 'group' @{displayName='GG_Compta'}),(M 'u2' 'user'),(M 'u4' 'user'))
 'g2bbbbbbbb'=@((M 'u1' 'user'),(M 'u5' 'user'),(M 'dev1' 'device'))
 'g3cccccccc'=@((M 'u1' 'user'),(M 'u6' 'user'))
 'g4dddddddd'=@((M 'u1' 'user'),(M 'u3' 'user'))
 'g5eeeeeeee'=@()
}
function Get-MgGroup { param([switch]$All,$Property) $global:G }
function Get-MgGroupMember { param($GroupId,[switch]$All,$ErrorAction) $global:MB[$GroupId] }
