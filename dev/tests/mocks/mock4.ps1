. ./mock3.ps1
function Get-MgContext { [pscustomobject]@{ TenantId='00000000-test'; Account='paj-a2@contoso.ch'; Scopes=@('GroupMember.Read.All','User.Read.All','LicenseAssignment.Read.All','RoleManagement.Read.Directory','Reports.Read.All') } }
function Get-MgReportOffice365ActiveUserDetail { param($Period,$OutFile,$ErrorAction)
  $d=(Get-Date); @(
   [pscustomobject]@{'User Principal Name'='marie.dupont@contoso.ch';'Exchange Last Activity Date'=$d.AddDays(-2).ToString('yyyy-MM-dd');'Teams Last Activity Date'=$d.AddDays(-1).ToString('yyyy-MM-dd');'OneDrive Last Activity Date'='';'SharePoint Last Activity Date'='';'Skype For Business Last Activity Date'='';'Yammer Last Activity Date'=''},
   [pscustomobject]@{'User Principal Name'='luc.ancien@contoso.ch';'Exchange Last Activity Date'='';'Teams Last Activity Date'='';'OneDrive Last Activity Date'='';'SharePoint Last Activity Date'='';'Skype For Business Last Activity Date'='';'Yammer Last Activity Date'=''},
   [pscustomobject]@{'User Principal Name'='stag@contoso.ch';'Exchange Last Activity Date'='';'Teams Last Activity Date'='';'OneDrive Last Activity Date'='';'SharePoint Last Activity Date'='';'Skype For Business Last Activity Date'='';'Yammer Last Activity Date'=''}
  ) | Export-Csv $OutFile -NoTypeInformation }
