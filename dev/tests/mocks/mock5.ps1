. ./mock4.ps1
function Get-MgReportOffice365ActiveUserDetail { param($Period,$OutFile,$ErrorAction)
  @([pscustomobject]@{'User Principal Name'='A1B2C3D4E5F6A7B8C9D0E1F2A3B4C5D6';'Exchange Last Activity Date'='2026-09-01'}) | Export-Csv $OutFile -NoTypeInformation }
