<#
.SYNOPSIS
    Teste les scripts d'export MapTools avec des cmdlets Graph, AD, Exchange et SharePoint simulées (dossier mocks).
    Aucun tenant n'est contacté. À lancer depuis PowerShell 7 :  .\dev\tests\Test-Scripts.ps1
#>
$ErrorActionPreference = 'Stop'
$scripts = Join-Path (Split-Path (Split-Path $PSScriptRoot)) 'scripts'
$mocks = Join-Path $PSScriptRoot 'mocks'
$sortie = Join-Path ([System.IO.Path]::GetTempPath()) "MapToolsTests_$([guid]::NewGuid().ToString('N').Substring(0,8))"
New-Item -ItemType Directory -Path $sortie | Out-Null
$cas = @(
    @{ Nom = 'Entra (MFA, licences, rôles, propriétaires)'; Mock = 'mock3.ps1'; Script = 'Export-MapToolsEntra.ps1'; Args = '-Prefixe E1'; Attendus = @('E1_*_MapTools_Liens.csv', 'E1_*_MapTools_Attributs.csv') }
    @{ Nom = 'Entra sans P1 (rapport d''activité)';         Mock = 'mock4.ps1'; Script = 'Export-MapToolsEntra.ps1'; Args = '-Prefixe E2'; Attendus = @('E2_*_MapTools_Attributs.csv') }
    @{ Nom = 'Active Directory';                            Mock = 'mockad.ps1'; Script = 'Export-MapToolsAD.ps1';   Args = '-Prefixe A1'; Attendus = @('A1_*_AD_MapTools_Liens.csv', 'A1_*_AD_MapTools_Attributs.csv') }
    @{ Nom = 'Partages et sites';                           Mock = 'mockp.ps1'; Script = 'Export-MapToolsPartages.ps1'; Args = '-Prefixe P1 -Jours 30'; Attendus = @('P1_*_MapTools_Partages.csv', 'P1_*_MapTools_Sites.csv') }
)
$echecs = 0
foreach ($c in $cas) {
    $cmd = "Set-Location '$mocks'; . ./$($c.Mock); & '$(Join-Path $scripts $c.Script)' $($c.Args) -DossierExport '$sortie' *>&1 | Out-Null"
    pwsh -NoProfile -Command $cmd
    $manquants = @($c.Attendus | Where-Object { -not (Get-ChildItem (Join-Path $sortie $_) -ErrorAction SilentlyContinue) })
    if ($manquants.Count) { $echecs++; Write-Host "ÉCHEC $($c.Nom) : $($manquants -join ', ') absent(s)" -ForegroundColor Red }
    else { Write-Host "OK    $($c.Nom)" -ForegroundColor Green }
}
Remove-Item $sortie -Recurse -Force
if ($echecs) { exit 1 }
