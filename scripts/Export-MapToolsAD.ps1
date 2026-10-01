<#
.SYNOPSIS
    Exporte les groupes Active Directory (on-prem), leurs imbrications et leurs membres
    utilisateurs (avec statut d'activité) au format attendu par MapTools.html.

.DESCRIPTION
    Deux CSV dans -DossierExport :
      - <prefixe>_<date>_AD_MapTools_Liens.csv     : Parent ; Enfant ; TypeEnfant
      - <prefixe>_<date>_AD_MapTools_Attributs.csv : une ligne par groupe et par utilisateur

    Statut (colonne Categorie) : Désactivé, Actif, Inactif > N mois, Jamais connecté,
    Nouveau (jamais connecté). Date de référence : LastLogonDate (lastLogonTimestamp),
    répliquée avec un décalage possible de 9 à 14 jours : suffisant pour un seuil en mois.

    Seuls les objets du domaine interrogé sont reliés ; les membres d'autres domaines
    ou forêts (Foreign Security Principals) sont comptés dans NbAutres.

    Dans la carte : couleur = type de groupe / statut utilisateur,
    pointillés = Protege (adminCount = 1).

.PREREQUIS
    Module ActiveDirectory (RSAT ou DC). Lecture simple de l'annuaire.

.EXAMPLE
    .\Export-MapToolsAD.ps1 -Prefixe Contoso -DossierExport C:\Temp
    .\Export-MapToolsAD.ps1 -Server dc01.contoso.local -SearchBase "OU=Groupes,DC=contoso,DC=local" -SansUtilisateurs
.NOTES
    Outil     : MapTools 1.1
    Auteur    : Jimmy Paul
    Copyright : © 2026 Jimmy Paul. Voir LICENCE.txt.
#>

[CmdletBinding()]
param(
    [string]$Server,
    [string]$SearchBase,
    [string]$DossierExport = ".",
    [string]$Prefixe = "AD",
    [int]$MoisInactivite = 6,
    [switch]$SansUtilisateurs
)

Write-Host "MapTools 1.1 · export Active Directory" -ForegroundColor DarkCyan

# Règles de profil (heuristique, testées dans l'ordre sur le SamAccountName en minuscules)
$ReglesProfil = [ordered]@{
    'Service / sync'     = '^(svc[-_.]|sa[-_.]|msol_|aad_|sync_)'
    'Admin'              = '-a2$|-admin$|^adm[-_.]|^admin[-_.]'
    'Salle / équipement' = '^(salle|room|scan|copieur|printer)'
    'Technique / test'   = '^(test|install|krbtgt$|guest$|defaultaccount$)'
}

Import-Module ActiveDirectory -ErrorAction Stop
$commun = @{}
if ($Server) { $commun.Server = $Server }

$maintenant = Get-Date
$dateSeuil  = $maintenant.AddMonths(-$MoisInactivite)
$fmt = 'yyyy-MM-dd HH:mm'

# --- Groupes (toujours tout le domaine pour résoudre les imbrications ; SearchBase filtre l'affichage) ---
Write-Host "Lecture des groupes AD..." -ForegroundColor Cyan
$groupes = Get-ADGroup @commun -Filter * -Properties member, GroupCategory, GroupScope, adminCount, whenCreated, Description
$parDN = @{}; $groupes | ForEach-Object { $parDN[$_.DistinguishedName] = $_ }
$cible = if ($SearchBase) { $groupes | Where-Object { $_.DistinguishedName -like "*,$SearchBase" } } else { $groupes }
Write-Host "$($groupes.Count) groupes dans le domaine, $(@($cible).Count) exportés" -ForegroundColor Cyan

# --- Utilisateurs ---
$parDNUser = @{}
if (-not $SansUtilisateurs) {
    Write-Host "Lecture des utilisateurs AD..." -ForegroundColor Cyan
    Get-ADUser @commun -Filter * -Properties DisplayName, UserPrincipalName, Enabled, LastLogonDate, whenCreated, adminCount, PasswordNeverExpires, PasswordLastSet |
        ForEach-Object { $parDNUser[$_.DistinguishedName] = $_ }
    Write-Host "$($parDNUser.Count) utilisateurs" -ForegroundColor Cyan
}

function Get-OU($dn) { ($dn -split '(?<!\\),', 2)[1] }
$doublons = $groupes | Group-Object Name | Where-Object Count -gt 1 | ForEach-Object Name
function Get-CleGroupe($g) { if ($g.Name -in $doublons) { "$($g.Name) [$((Get-OU $g.DistinguishedName) -replace ',DC=.*$','')]" } else { $g.Name } }
function Get-CleUser($u)   { if ($u.UserPrincipalName) { $u.UserPrincipalName } else { $u.SamAccountName } }
function Get-Statut($u) {
    if (-not $u.Enabled) { return 'Désactivé' }
    if (-not $u.LastLogonDate) { if ($u.whenCreated -gt $dateSeuil) { return 'Nouveau (jamais connecté)' } return 'Jamais connecté' }
    if ($u.LastLogonDate -lt $dateSeuil) { return "Inactif > $MoisInactivite mois" }
    return 'Actif'
}
function Get-Profil($u) {
    $sam = "$($u.SamAccountName)".ToLower()
    foreach ($k in $ReglesProfil.Keys) { if ($sam -match $ReglesProfil[$k]) { return $k } }
    return 'Standard'
}

# --- Liens ---
# Piège : Get-ADGroupMember échoue au-delà de 5000 membres (limite ADWS MaxGroupOrMemberEntries).
# Ce script lit la propriété member via Get-ADGroup ; contrôler les très gros groupes.
$liens = [System.Collections.Generic.List[object]]::new()
$stats = @{}; $nbGroupesUser = @{}; $usersRelies = @{}
foreach ($g in $cible) {
    $s = @{ Sous = 0; Users = 0; Autres = 0 }
    if ($g.member.Count -ge 5000) { Write-Warning "'$($g.Name)' : $($g.member.Count) membres, vérifier la limite ADWS." }
    foreach ($dn in $g.member) {
        if ($parDN.ContainsKey($dn)) {
            $s.Sous++
            $liens.Add([PSCustomObject]@{ Parent = Get-CleGroupe $g; Enfant = Get-CleGroupe $parDN[$dn]; TypeEnfant = 'Groupe' })
        } elseif ($parDNUser.ContainsKey($dn)) {
            $s.Users++
            $u = $parDNUser[$dn]
            $nbGroupesUser[$dn] = 1 + [int]$nbGroupesUser[$dn]; $usersRelies[$dn] = $true
            $liens.Add([PSCustomObject]@{ Parent = Get-CleGroupe $g; Enfant = Get-CleUser $u; TypeEnfant = 'Utilisateur' })
        } else { $s.Autres++ }   # ordinateurs, contacts, FSP, utilisateurs ignorés (-SansUtilisateurs)
    }
    $stats[$g.DistinguishedName] = $s
}

# --- Attributs (mêmes colonnes partout) ---
function New-Ligne { [ordered]@{
    Cle=''; Libelle=''; TypeElement=''; Categorie=''; Profil=''; UPN=''; SamAccountName=''; Protege=''; Etendue=''
    DerniereConnexion=''; JoursInactivite=''; MdpNExpireJamais=''; MdpDernierChangement=''; CreeLe=''; NbGroupesDirects=''
    NbMembresDirects=''; NbSousGroupes=''; NbUtilisateurs=''; NbAutres=''; OU='' } }

$attributs = [System.Collections.Generic.List[object]]::new()
foreach ($g in $cible) {
    $l = New-Ligne; $s = $stats[$g.DistinguishedName]
    $l.Cle = Get-CleGroupe $g; $l.Libelle = $g.Name; $l.TypeElement = 'Groupe'
    $l.Categorie = if ($g.GroupCategory -eq 'Security') { 'Sécurité' } else { 'Distribution' }
    $l.SamAccountName = $g.SamAccountName; $l.Protege = ($g.adminCount -eq 1); $l.Etendue = $g.GroupScope
    $l.CreeLe = $g.whenCreated.ToString('yyyy-MM-dd')
    $l.NbSousGroupes = $s.Sous; $l.NbUtilisateurs = $s.Users; $l.NbAutres = $s.Autres; $l.NbMembresDirects = $s.Sous + $s.Users + $s.Autres
    $l.OU = Get-OU $g.DistinguishedName
    $attributs.Add([PSCustomObject]$l)
}
$compteStatut = @{}
foreach ($dn in $parDNUser.Keys) {
    $u = $parDNUser[$dn]
    if ($SearchBase -and -not $usersRelies[$dn] -and $dn -notlike "*,$SearchBase") { continue }
    $l = New-Ligne
    $l.Cle = Get-CleUser $u; $l.Libelle = if ($u.DisplayName) { $u.DisplayName } else { $u.Name }; $l.TypeElement = 'Utilisateur'
    $l.Categorie = Get-Statut $u; $l.Profil = Get-Profil $u
    $l.UPN = $u.UserPrincipalName; $l.SamAccountName = $u.SamAccountName; $l.Protege = ($u.adminCount -eq 1)
    if ($u.LastLogonDate) { $l.DerniereConnexion = $u.LastLogonDate.ToString($fmt); $l.JoursInactivite = [int]($maintenant - $u.LastLogonDate).TotalDays }
    $l.MdpNExpireJamais = $u.PasswordNeverExpires
    if ($u.PasswordLastSet) { $l.MdpDernierChangement = $u.PasswordLastSet.ToString('yyyy-MM-dd') }
    $l.CreeLe = $u.whenCreated.ToString('yyyy-MM-dd'); $l.NbGroupesDirects = [int]$nbGroupesUser[$dn]
    $l.OU = Get-OU $dn
    $attributs.Add([PSCustomObject]$l)
    $compteStatut[$l.Categorie] = 1 + [int]$compteStatut[$l.Categorie]
}

# --- Export ---
$encodage = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
$horodatage = Get-Date -Format 'yyyyMMdd_HHmm'
$fLiens = Join-Path $DossierExport "$($Prefixe)_$($horodatage)_AD_MapTools_Liens.csv"
$fAttr  = Join-Path $DossierExport "$($Prefixe)_$($horodatage)_AD_MapTools_Attributs.csv"
$liens     | Export-Csv $fLiens -NoTypeInformation -Delimiter ';' -Encoding $encodage
$attributs | Sort-Object TypeElement, Cle | Export-Csv $fAttr -NoTypeInformation -Delimiter ';' -Encoding $encodage

Write-Host "`n=== Résumé ===" -ForegroundColor Green
Write-Host "Groupes exportés : $(@($cible).Count) | liens groupe → groupe : $(@($liens | Where-Object TypeEnfant -eq 'Groupe').Count)"
if (-not $SansUtilisateurs) {
    $total = ($compteStatut.Values | Measure-Object -Sum).Sum
    Write-Host "Utilisateurs exportés : $total | seuil : $MoisInactivite mois (avant le $($dateSeuil.ToString('dd.MM.yyyy')))"
    $compteStatut.GetEnumerator() | ForEach-Object {
        [PSCustomObject]@{ Statut = $_.Key; Nombre = $_.Value; 'Pct total' = '{0:P1}' -f ($_.Value / [math]::Max(1,$total)) }
    } | Sort-Object Statut | Format-Table -AutoSize
}
Write-Host "Liens     : $fLiens" -ForegroundColor Green
Write-Host "Attributs : $fAttr" -ForegroundColor Green
Write-Host "Dépose les deux fichiers dans MapTools.html (Pointillés si vrai = Protege)" -ForegroundColor Green
