function Connect-ExchangeOnline {}
function Get-ConnectionInformation { [pscustomobject]@{State='Connected'} }
function Get-MgDomain { param([switch]$All,$ErrorAction) @([pscustomobject]@{Id='contoso.ch'},[pscustomobject]@{Id='contoso.onmicrosoft.com'}) }
$global:calls=@{}
function Search-UnifiedAuditLog { param($StartDate,$EndDate,$RecordType,$Operations,$SessionId,$SessionCommand,$ResultSize,$ErrorAction)
  if($global:calls[$SessionId]){ return @() }; $global:calls[$SessionId]=1
  if($StartDate -gt (Get-Date).AddDays(-20)){ return @() }
  $d=$StartDate.AddDays(1)
  $mk={ param($id,$op,$user,$tgt,$tt,$obj,$item) [pscustomobject]@{Identity=$id;AuditData=(@{CreationTime=$d.ToString('o');Operation=$op;UserId=$user;TargetUserOrGroupName=$tgt;TargetUserOrGroupType=$tt;ObjectId=$obj;SourceFileName=($obj -split '/')[-1];ItemType=$item;SiteUrl='https://contoso.sharepoint.com/sites/Finance/';Workload='SharePoint';Permission='Read'}|ConvertTo-Json)} }
  $k=$StartDate.Ticks
  @(
   (& $mk "a$k" 'AnonymousLinkCreated' 'marie@contoso.ch' '' '' 'https://contoso.sharepoint.com/sites/Finance/Shared Documents/Budget.xlsx' 'File'),
   (& $mk "b$k" 'SharingInvitationCreated' 'marie@contoso.ch' 'john@gmail.com' 'Guest' 'https://contoso.sharepoint.com/sites/Finance/Shared Documents/Rapport.docx' 'File'),
   (& $mk "c$k" 'AddedToSecureLink' 'luc@contoso.ch' 'paul@partner.ch' '' 'https://contoso.sharepoint.com/sites/Finance/Shared Documents/Contrat.pdf' 'File'),
   (& $mk "d$k" 'CompanyLinkCreated' 'luc@contoso.ch' '' 'Member' 'https://contoso.sharepoint.com/sites/Finance/Shared Documents/Note.docx' 'File'),
   (& $mk "e$k" 'AnonymousLinkRemoved' 'marie@contoso.ch' '' '' 'https://contoso.sharepoint.com/sites/Finance/Shared Documents/Budget.xlsx' 'File')
  ) | Select-Object -First ($(if($StartDate -gt (Get-Date).AddDays(-25)){5}else{4}))
}
function Get-SPOTenant { [pscustomobject]@{SharingCapability='ExternalUserAndGuestSharing';RequireAnonymousLinksExpireInDays=30;DefaultSharingLinkType='AnonymousAccess'} }
function Get-SPOSite { param($Limit,$IncludePersonalSite,$ErrorAction)
  @([pscustomobject]@{Url='https://contoso.sharepoint.com/sites/Finance';Title='Finance';Template='GROUP#0';SharingCapability='ExternalUserAndGuestSharing';GroupId='11111111-1111-1111-1111-111111111111';Owner='marie@contoso.ch';LastContentModifiedDate=(Get-Date);StorageUsageCurrent=1200},
    [pscustomobject]@{Url='https://contoso.sharepoint.com/sites/RH';Title='RH';Template='STS#3';SharingCapability='Disabled';GroupId='00000000-0000-0000-0000-000000000000';Owner='rh@contoso.ch';LastContentModifiedDate=(Get-Date).AddDays(-400);StorageUsageCurrent=50}) }
function Connect-SPOService {}
