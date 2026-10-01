function Import-Module {}
$now=Get-Date
$D='DC=contoso,DC=local'
function G($n,$ou,$cat,$members,$admin){ [pscustomobject]@{Name=$n;SamAccountName=$n;DistinguishedName="CN=$n,$ou,$D";GroupCategory=$cat;GroupScope='Global';adminCount=$admin;whenCreated=$now.AddYears(-4);member=$members;Description=''} }
function Us($sam,$ou,$en,$ll,$cr,$adm,$nev=$false,$pls=100){ [pscustomobject]@{Name=$sam;SamAccountName=$sam;DisplayName="User $sam";UserPrincipalName="$sam@contoso.ch";DistinguishedName="CN=$sam,$ou,$D";Enabled=$en;LastLogonDate=$ll;whenCreated=$cr;adminCount=$adm;PasswordNeverExpires=$nev;PasswordLastSet=$now.AddDays(-$pls)} }
$global:UU=@((Us 'mdupont' 'OU=Users' $true $now.AddDays(-2) $now.AddYears(-3) $null $false 500),(Us 'adm-bob' 'OU=Admins' $true $now.AddDays(-5) $now.AddYears(-3) 1 $true 800),(Us 'lancien' 'OU=Users' $true $now.AddDays(-500) $now.AddYears(-3) $null),(Us 'svc-sql' 'OU=Svc' $true $null $now.AddYears(-5) $null $true 900),(Us 'jdo-a2' 'OU=Admins' $false $null $now.AddYears(-1) 1))
$global:GG=@(
 (G 'GG_Finance' 'OU=Groupes' 'Security' @("CN=GG_Compta,OU=Groupes,$D","CN=lancien,OU=Users,$D","CN=PC01,OU=PC,$D") $null),
 (G 'GG_Compta' 'OU=Groupes' 'Security' @("CN=mdupont,OU=Users,$D") $null),
 (G 'Domain Admins' 'CN=Users' 'Security' @("CN=jdo-a2,OU=Admins,$D","CN=svc-sql,OU=Svc,$D","CN=adm-bob,OU=Admins,$D") 1),
 (G 'DL_Info' 'OU=Groupes' 'Distribution' @("CN=GG_Finance,OU=Groupes,$D") $null)
)
function Get-ADGroup { param($Filter,$Properties,$Server) $global:GG }
function Get-ADUser { param($Filter,$Properties,$Server) $global:UU }
