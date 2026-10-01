<#
.SYNOPSIS
    Liste les comptes Entra ID inactifs depuis plus de X mois, en séparant
    les comptes activés (à traiter) des comptes déjà désactivés.

.DESCRIPTION
    Utilise Microsoft Graph PowerShell pour lire signInActivity.
    Date de référence : lastSuccessfulSignInDateTime (connexion réussie uniquement).
    Repli sur lastSignInDateTime / lastNonInteractiveSignInDateTime si absente
    (propriété alimentée seulement depuis décembre 2023) -> colonne SourceDate.

    Deux exports CSV :
      - <prefixe>_Actifs_Inactifs.csv   : comptes ACTIVÉS sans connexion depuis > X mois
      - <prefixe>_Desactives.csv        : TOUS les comptes DÉSACTIVÉS, avec leur dernière activité

    Les comptes créés après la date seuil et jamais connectés ne sont pas
    considérés comme inactifs (nouveaux arrivants).

.PREREQUIS
    - Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Users -Scope CurrentUser
    - Scopes : User.Read.All + AuditLog.Read.All
    - Rôle Entra du compte connecté : Reports Reader, Security Reader ou Global Reader minimum
      (sans rôle, signInActivity revient vide / 403 même avec le consentement des scopes)
    - Licence Entra ID P1/P2 dans le tenant pour que signInActivity soit renseigné

.EXAMPLE
    .\Get-EntraIDComptesInactifs.ps1 -TenantId "client.onmicrosoft.com"
    .\Get-EntraIDComptesInactifs.ps1 -TenantId "xxxxxxxx-xxxx-..." -MoisInactivite 3 -InclureInvites
.NOTES
    Outil     : MapTools 1.1
    Auteur    : Jimmy Paul
    Copyright : © 2026 Jimmy Paul. Voir LICENCE.txt.
#>

[CmdletBinding()]
param(
    # Tenant cible (obligatoire en accès délégué GDAP, sinon connexion au tenant du compte)
    [string]$TenantId,

    [int]$MoisInactivite = 6,

    [string]$DossierExport = ".",

    [switch]$InclureInvites
)

Write-Host "MapTools 1.1 · audit des comptes inactifs" -ForegroundColor DarkCyan

# --- 1. Connexion (réutilise la session existante si présente) ---
$scopes = @("User.Read.All", "AuditLog.Read.All")
$ctx = Get-MgContext
$connexionParScript = $false

if (-not $ctx) {
    Write-Host "Aucune session Graph : connexion..." -ForegroundColor Cyan
    $params = @{ Scopes = $scopes; NoWelcome = $true }
    if ($TenantId) { $params.TenantId = $TenantId }
    Connect-MgGraph @params
    $connexionParScript = $true
    $ctx = Get-MgContext
} else {
    $manquants = $scopes | Where-Object { $_ -notin $ctx.Scopes }
    if ($manquants) {
        Write-Warning "Session existante sans les scopes : $($manquants -join ', '). Reconnecte-toi avec :`nConnect-MgGraph -Scopes '$($scopes -join "','")'"
        return
    }
}
Write-Host "Tenant connecté : $($ctx.TenantId) | Compte : $($ctx.Account)" -ForegroundColor Cyan

$maintenant = Get-Date
$dateSeuil  = $maintenant.AddMonths(-$MoisInactivite)
Write-Host "Seuil : aucune connexion réussie depuis le $($dateSeuil.ToString('dd/MM/yyyy'))" -ForegroundColor Cyan

# --- 2. Récupération ---
$proprietes = @(
    "Id", "DisplayName", "UserPrincipalName", "AccountEnabled",
    "UserType", "CreatedDateTime", "OnPremisesSyncEnabled", "SignInActivity"
)

Write-Host "Récupération des utilisateurs..." -ForegroundColor Cyan
try {
    $tousLesUtilisateurs = Get-MgUser -All -Property $proprietes -ErrorAction Stop
} catch {
    Write-Error "Échec Get-MgUser : $($_.Exception.Message)`nVérifier le rôle Entra (Reports Reader / Global Reader) et la licence P1/P2."
    if ($connexionParScript) { Disconnect-MgGraph | Out-Null }
    return
}

# --- 3. Analyse ---
$actifsInactifs = [System.Collections.Generic.List[object]]::new()
$desactives     = [System.Collections.Generic.List[object]]::new()
$nbActifsAnalyses = 0

foreach ($u in $tousLesUtilisateurs) {

    if (-not $InclureInvites -and $u.UserType -eq "Guest") { continue }

    $sia = $u.SignInActivity

    # Date de référence : connexion réussie en priorité
    if ($sia.LastSuccessfulSignInDateTime) {
        $derniereActivite = [datetime]$sia.LastSuccessfulSignInDateTime
        $source = "Succès"
    } else {
        $dates = @($sia.LastSignInDateTime, $sia.LastNonInteractiveSignInDateTime) |
                 Where-Object { $_ } | ForEach-Object { [datetime]$_ }
        if ($dates) {
            $derniereActivite = ($dates | Sort-Object -Descending)[0]
            $source = "Tentative (succès inconnu)"
        } else {
            $derniereActivite = $null
            $source = "Aucune"
        }
    }

    $jours = if ($derniereActivite) { [math]::Round(($maintenant - $derniereActivite).TotalDays) } else { $null }

    $ligne = [PSCustomObject]@{
        DisplayName       = $u.DisplayName
        UserPrincipalName = $u.UserPrincipalName
        AccountEnabled    = $u.AccountEnabled
        UserType          = $u.UserType
        SyncAD            = [bool]$u.OnPremisesSyncEnabled
        CreatedDateTime   = $u.CreatedDateTime
        DerniereConnexion = $derniereActivite
        SourceDate        = $source
        JoursInactivite   = $jours
        Statut            = ""
    }

    # --- Comptes désactivés : liste séparée, tous ---
    if ($u.AccountEnabled -eq $false) {
        $ligne.Statut = if (-not $derniereActivite) { "Désactivé - jamais connecté" }
                        elseif ($derniereActivite -lt $dateSeuil) { "Désactivé - inactif > $MoisInactivite mois" }
                        else { "Désactivé - activité récente" }
        $desactives.Add($ligne)
        continue
    }

    # --- Comptes activés ---
    $nbActifsAnalyses++

    if (-not $derniereActivite) {
        if ($u.CreatedDateTime -and [datetime]$u.CreatedDateTime -gt $dateSeuil) { continue }  # nouveau compte
        $ligne.Statut = "Jamais connecté"
        $actifsInactifs.Add($ligne)
    }
    elseif ($derniereActivite -lt $dateSeuil) {
        $ligne.Statut = "Inactif depuis $jours jours"
        $actifsInactifs.Add($ligne)
    }
}

# --- 4. Résumé ---
Write-Host "`n=== Résumé ===" -ForegroundColor Green
Write-Host "Utilisateurs dans le tenant                  : $($tousLesUtilisateurs.Count)"
Write-Host "Comptes activés analysés                     : $nbActifsAnalyses"
Write-Host "Comptes activés inactifs > $MoisInactivite mois          : $($actifsInactifs.Count)"
if ($nbActifsAnalyses -gt 0) {
    $pct = [math]::Round(($actifsInactifs.Count / $nbActifsAnalyses) * 100, 1)
    Write-Host ">>> $pct % des comptes activés sont inactifs <<<" -ForegroundColor Yellow
}
Write-Host "Comptes désactivés                           : $($desactives.Count)"

# --- 5. Export (séparateur ; pour Excel FR, BOM pour les accents) ---
$encodage = if ($PSVersionTable.PSVersion.Major -ge 6) { "utf8BOM" } else { "UTF8" }
$prefixe  = Join-Path $DossierExport "EntraID_$(Get-Date -Format 'yyyyMMdd_HHmmss')"

if ($actifsInactifs.Count) {
    Write-Host "`n--- Comptes ACTIVÉS inactifs ---" -ForegroundColor Yellow
    $actifsInactifs | Sort-Object DerniereConnexion |
        Format-Table DisplayName, UserPrincipalName, DerniereConnexion, SourceDate, Statut -AutoSize
    $actifsInactifs | Sort-Object DerniereConnexion |
        Export-Csv "$($prefixe)_Actifs_Inactifs.csv" -NoTypeInformation -Delimiter ';' -Encoding $encodage
    Write-Host "Export : $($prefixe)_Actifs_Inactifs.csv" -ForegroundColor Green
}

if ($desactives.Count) {
    Write-Host "`n--- Comptes DÉSACTIVÉS ---" -ForegroundColor DarkGray
    $desactives | Sort-Object DerniereConnexion |
        Format-Table DisplayName, UserPrincipalName, DerniereConnexion, Statut -AutoSize
    $desactives | Sort-Object DerniereConnexion |
        Export-Csv "$($prefixe)_Desactives.csv" -NoTypeInformation -Delimiter ';' -Encoding $encodage
    Write-Host "Export : $($prefixe)_Desactives.csv" -ForegroundColor Green
}

if ($connexionParScript) { Disconnect-MgGraph | Out-Null }
