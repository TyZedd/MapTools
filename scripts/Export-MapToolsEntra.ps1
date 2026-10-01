<#
.SYNOPSIS
    Exporte les groupes Entra ID, leurs imbrications et leurs membres utilisateurs
    (avec statut d'activité) au format attendu par MapTools.html.

.DESCRIPTION
    Deux CSV dans -DossierExport :
      - <prefixe>_<date>_MapTools_Liens.csv     : Parent ; Enfant ; TypeEnfant
      - <prefixe>_<date>_MapTools_Attributs.csv : une ligne par groupe et par utilisateur

    Statut d'un utilisateur (colonne Categorie) :
      Désactivé                  AccountEnabled = false
      Actif                      connexion réussie depuis moins de -MoisInactivite mois
      Inactif > N mois           dernière connexion plus ancienne
      Jamais connecté            aucune connexion, compte créé avant le seuil
      Nouveau (jamais connecté)  aucune connexion, compte créé après le seuil
      Activité inconnue          signInActivity illisible (scope, rôle ou licence manquant)

    Date de référence : lastSuccessfulSignInDateTime ; à défaut (données antérieures à
    décembre 2023) la plus récente de lastSignInDateTime / lastNonInteractiveSignInDateTime,
    signalée par SourceDate = « Tentative (succès inconnu) ».

    Profil (colonne Profil) : classification heuristique sur le nom / UPN, réglable dans
    $ReglesProfil. À confirmer côté Exchange pour les boîtes partagées et salles.

.PREREQUIS
    Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Groups, Microsoft.Graph.Reports, Microsoft.Graph.Identity.DirectoryManagement, Microsoft.Graph.Identity.Governance -Scope CurrentUser
    Scopes : GroupMember.Read.All, User.Read.All, AuditLog.Read.All, LicenseAssignment.Read.All, RoleManagement.Read.Directory
    MFA (userRegistrationDetails) : AuditLog.Read.All, même rôle Entra et licence P1/P2 que signInActivity
    Licences (subscribedSkus) : LicenseAssignment.Read.All, rôle Directory Readers ou Global Reader
    Secours sans licence P1 : Reports.Read.All (rapport d'activité Microsoft 365, 180 jours) ; les noms ne doivent pas être anonymisés
      (centre d'administration Microsoft 365 > Paramètres > Paramètres de l'organisation > Rapports)
    Rôles Entra : RoleManagement.Read.Directory ; les attributions éligibles et activées (PIM) nécessitent Entra ID P2
    Propriétaires : lus pour les groupes cloud de sécurité et Microsoft 365 (Graph ne les expose pas pour les groupes synchronisés ni les listes de distribution)
    Rôle Entra : Reports Reader, Security Reader ou Global Reader (pour signInActivity)
    Licence Entra ID P1/P2 dans le tenant (pour signInActivity)

.EXAMPLE
    Connect-MgGraph -Scopes "GroupMember.Read.All","User.Read.All","AuditLog.Read.All" -TenantId "contoso.ch" -NoWelcome
    .\Export-MapToolsEntra.ps1 -Prefixe Contoso -DossierExport C:\Temp

.EXAMPLE
    .\Export-MapToolsEntra.ps1 -SansUtilisateurs      # groupes et imbrications uniquement
    .\Export-MapToolsEntra.ps1 -SansMFA -SansLicences  # sans statut MFA ni licences
    .\Export-MapToolsEntra.ps1 -SansRoles -SansProprietaires
    .\Export-MapToolsEntra.ps1 -Prefixe Contoso -AvecPartages      # + partages et sites (Export-MapToolsPartages.ps1)
.NOTES
    Outil     : MapTools 1.1
    Auteur    : Jimmy Paul
    Copyright : © 2026 Jimmy Paul. Voir LICENCE.txt.
#>

[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$DossierExport = ".",
    [string]$Prefixe = "Entra",
    [int]$MoisInactivite = 6,
    [switch]$SansUtilisateurs,
    [switch]$SansMFA,
    [switch]$SansLicences,
    [switch]$SansRoles,
    [switch]$SansProprietaires,
    # Enchaîne Export-MapToolsPartages.ps1 (audit des partages et paramètres des sites) avec le même préfixe et horodatage
    [switch]$AvecPartages,
    [string]$SharePointAdminUrl
)

Write-Host "MapTools 1.1 · export Entra ID" -ForegroundColor DarkCyan

# --- Règles de profil (heuristique, testées dans l'ordre sur l'UPN en minuscules) ---
$ReglesProfil = [ordered]@{
    'Service / sync'     = '^(svc[-_.]|sync_|adsyncsvc)|sbc'
    'Admin'              = '-a2@|-admin@|^adm[-_.]|^admin[-_.]'
    'Salle / équipement' = '^(salle|room|booking|wireless|ipad|tv\.|parking|massage|scan)'
    'Technique / test'   = '^(test|install|no-?reply|dmarc|audit\d|intune\.)'
}

# --- 1. Connexion (réutilise la session existante) ---
$scopesVoulus = @('GroupMember.Read.All', 'User.Read.All', 'AuditLog.Read.All', 'LicenseAssignment.Read.All', 'RoleManagement.Read.Directory', 'Reports.Read.All')
$ctx = Get-MgContext
$connexionParScript = $false
if (-not $ctx) {
    $p = @{ Scopes = $scopesVoulus; NoWelcome = $true }
    if ($TenantId) { $p.TenantId = $TenantId }
    Connect-MgGraph @p
    $connexionParScript = $true
    $ctx = Get-MgContext
}
$aScope = { param($liste) [bool]($ctx.Scopes | Where-Object { $_ -in $liste }) }
if (-not (& $aScope @('GroupMember.Read.All', 'Group.Read.All', 'Directory.Read.All'))) {
    Write-Warning "Scope manquant pour lire les groupes. Reconnecte-toi :`nConnect-MgGraph -Scopes '$($scopesVoulus -join "','")'"
    return
}
if (-not $SansUtilisateurs -and -not (& $aScope @('User.Read.All', 'Directory.Read.All'))) {
    Write-Warning "Scope manquant pour lire les utilisateurs (User.Read.All). Relance avec -SansUtilisateurs ou reconnecte-toi."
    return
}
$lireConnexions = (-not $SansUtilisateurs) -and (& $aScope @('AuditLog.Read.All'))
if (-not $SansUtilisateurs -and -not $lireConnexions) {
    Write-Warning "AuditLog.Read.All absent : pas de date de connexion Entra (secours possible par le rapport d'activité Microsoft 365)."
}
Write-Host "Tenant : $($ctx.TenantId) | Compte : $($ctx.Account)" -ForegroundColor Cyan

$maintenant = Get-Date
$dateSeuil  = $maintenant.AddMonths(-$MoisInactivite)
$fmt = 'yyyy-MM-dd HH:mm'

# --- 2. Utilisateurs ---
$parIdUser = @{}
if (-not $SansUtilisateurs) {
    Write-Host "Récupération des utilisateurs..." -ForegroundColor Cyan
    $propsU = @('Id', 'DisplayName', 'UserPrincipalName', 'AccountEnabled', 'UserType', 'CreatedDateTime', 'OnPremisesSyncEnabled', 'AssignedLicenses')
    $users = $null
    if ($lireConnexions) {
        try { $users = Get-MgUser -All -Property ($propsU + 'SignInActivity') -ErrorAction Stop }
        catch {
            Write-Warning "signInActivity illisible ($($_.Exception.Message.Split("`n")[0])). Vérifier le rôle Entra et la licence P1/P2. Export sans dates de connexion."
            $lireConnexions = $false
        }
    }
    if (-not $users) { $users = Get-MgUser -All -Property $propsU }
    $users | ForEach-Object { $parIdUser[$_.Id] = $_ }
    Write-Host "$($users.Count) utilisateurs" -ForegroundColor Cyan
}

# --- 2b. MFA enregistrée (rapport userRegistrationDetails) ---
$mfaParId = @{}
$lireMFA = (-not $SansUtilisateurs) -and (-not $SansMFA)
if ($lireMFA) {
    if (-not (Get-Command Get-MgReportAuthenticationMethodUserRegistrationDetail -ErrorAction SilentlyContinue)) {
        Write-Warning "Module Microsoft.Graph.Reports absent : export sans statut MFA (Install-Module Microsoft.Graph.Reports -Scope CurrentUser)."
        $lireMFA = $false
    } elseif (-not (& $aScope @('AuditLog.Read.All'))) {
        Write-Warning "AuditLog.Read.All absent : export sans statut MFA."
        $lireMFA = $false
    } else {
        Write-Host "Lecture de l'enregistrement MFA..." -ForegroundColor Cyan
        try {
            Get-MgReportAuthenticationMethodUserRegistrationDetail -All -Property Id, IsMfaRegistered, IsAdmin, MethodsRegistered -ErrorAction Stop |
                ForEach-Object { $mfaParId[$_.Id] = $_ }
            Write-Host "$($mfaParId.Count) fiches MFA" -ForegroundColor Cyan
        } catch {
            Write-Warning "Statut MFA illisible ($($_.Exception.Message.Split("`n")[0])). Vérifier le rôle Entra et la licence P1/P2. Export sans MFA."
            $lireMFA = $false
        }
    }
}

# --- 2c. Licences (noms des SKU du tenant) ---
$skuParId = @{}
$lireLicences = (-not $SansUtilisateurs) -and (-not $SansLicences)
# Noms commerciaux des SKU courants ; les autres gardent leur référence technique (SkuPartNumber)
$NomsSku = @{
    'SPE_E3'='Microsoft 365 E3'; 'SPE_E5'='Microsoft 365 E5'; 'SPE_F1'='Microsoft 365 F3'
    'SPB'='Microsoft 365 Business Premium'; 'O365_BUSINESS_PREMIUM'='Microsoft 365 Business Standard'; 'O365_BUSINESS_ESSENTIALS'='Microsoft 365 Business Basic'
    'O365_BUSINESS'='Microsoft 365 Apps for business'; 'OFFICESUBSCRIPTION'='Microsoft 365 Apps for enterprise'
    'ENTERPRISEPACK'='Office 365 E3'; 'ENTERPRISEPREMIUM'='Office 365 E5'; 'STANDARDPACK'='Office 365 E1'; 'DESKLESSPACK'='Office 365 F3'
    'EXCHANGESTANDARD'='Exchange Online (Plan 1)'; 'EXCHANGEENTERPRISE'='Exchange Online (Plan 2)'; 'EXCHANGEDESKLESS'='Exchange Online Kiosk'
    'AAD_PREMIUM'='Entra ID P1'; 'AAD_PREMIUM_P2'='Entra ID P2'; 'EMS'='EMS E3'; 'EMSPREMIUM'='EMS E5'; 'INTUNE_A'='Intune Plan 1'
    'POWER_BI_PRO'='Power BI Pro'; 'PBI_PREMIUM_PER_USER'='Power BI Premium par utilisateur'; 'POWER_BI_STANDARD'='Power BI (gratuit)'
    'VISIOCLIENT'='Visio Plan 2'; 'PROJECTPROFESSIONAL'='Project Plan 3'; 'FLOW_FREE'='Power Automate (gratuit)'
    'MCOEV'='Teams Phone Standard'; 'MCOPSTN1'='Teams Calling Plan (national)'; 'PHONESYSTEM_VIRTUALUSER'='Teams Phone Resource Account'
    'Microsoft_Teams_Rooms_Pro'='Teams Rooms Pro'; 'Microsoft_Teams_Rooms_Basic'='Teams Rooms Basic'; 'TEAMS_EXPLORATORY'='Teams Exploratory'
}
if ($lireLicences) {
    if (-not (Get-Command Get-MgSubscribedSku -ErrorAction SilentlyContinue)) {
        Write-Warning "Module Microsoft.Graph.Identity.DirectoryManagement absent : licences exportées par identifiant seulement."
    } else {
        try {
            Get-MgSubscribedSku -All -ErrorAction Stop | ForEach-Object { $skuParId["$($_.SkuId)"] = $_.SkuPartNumber }
            Write-Host "$($skuParId.Count) abonnements (SKU)" -ForegroundColor Cyan
        } catch {
            Write-Warning "Abonnements illisibles ($($_.Exception.Message.Split("`n")[0])). Scope LicenseAssignment.Read.All et rôle Directory Readers requis. Licences exportées par identifiant."
        }
    }
}
function Get-Licences($u) {
    $refs = @($u.AssignedLicenses | ForEach-Object { $id = "$($_.SkuId)"; if ($skuParId[$id]) { $skuParId[$id] } else { $id } })
    $noms = @($refs | ForEach-Object { if ($NomsSku[$_]) { $NomsSku[$_] } else { $_ } })
    return ($noms -join ' | '), ($refs -join ' | ')
}

# --- 2a. Secours : rapport d'activité Microsoft 365 (sans licence Entra ID P1) ---
# Dernière activité par service (Exchange, OneDrive, SharePoint, Teams…) sur 180 jours, pour les comptes licenciés.
$actParUpn = @{}; $sourceActivite = if ($lireConnexions) { 'connexions' } else { '' }
if (-not $SansUtilisateurs -and -not $lireConnexions) {
    if (-not (Get-Command Get-MgReportOffice365ActiveUserDetail -ErrorAction SilentlyContinue)) {
        Write-Warning "Pas de dates de connexion et module Microsoft.Graph.Reports absent : statuts d'activité inconnus."
    } elseif (-not (& $aScope @('Reports.Read.All'))) {
        Write-Warning "Pas de dates de connexion ; ajouter le scope Reports.Read.All pour utiliser le rapport d'activité Microsoft 365 en secours."
    } else {
        Write-Host "Dates de connexion indisponibles : lecture du rapport d'activité Microsoft 365 (180 jours)..." -ForegroundColor Cyan
        $tmp = [System.IO.Path]::GetTempFileName()
        try {
            Get-MgReportOffice365ActiveUserDetail -Period 'D180' -OutFile $tmp -ErrorAction Stop | Out-Null
            $colsAct = 'Exchange Last Activity Date','OneDrive Last Activity Date','SharePoint Last Activity Date','Skype For Business Last Activity Date','Yammer Last Activity Date','Teams Last Activity Date'
            foreach ($r in (Import-Csv $tmp)) {
                $d = $colsAct | ForEach-Object { $r.$_ } | Where-Object { $_ } | ForEach-Object { [datetime]$_ } | Sort-Object -Descending | Select-Object -First 1
                $actParUpn["$($r.'User Principal Name')".ToLower()] = $d
            }
            $connus = @($users | Where-Object { $actParUpn.ContainsKey("$($_.UserPrincipalName)".ToLower()) }).Count
            if ($actParUpn.Count -and -not $connus) {
                Write-Warning "Rapport d'activité anonymisé : aucun UPN reconnu. Désactiver l'anonymisation (centre d'administration Microsoft 365 > Paramètres > Paramètres de l'organisation > Rapports), puis relancer."
            } else {
                $sourceActivite = 'rapport'
                Write-Host "$connus comptes trouvés dans le rapport d'activité" -ForegroundColor Cyan
            }
        } catch {
            Write-Warning "Rapport d'activité illisible ($($_.Exception.Message.Split("`n")[0])). Statuts d'activité inconnus."
        } finally { Remove-Item $tmp -ErrorAction SilentlyContinue }
    }
}

function Get-DerniereActivite($u) {
    if ($sourceActivite -eq 'rapport') {
        $cle = "$($u.UserPrincipalName)".ToLower()
        if (-not $actParUpn.ContainsKey($cle)) { return $null, 'Hors rapport d''activité (sans licence)' }
        if ($actParUpn[$cle]) { return [datetime]$actParUpn[$cle], 'Rapport d''activité M365' }
        return $null, 'Rapport d''activité M365 (aucune activité sur 180 j)'
    }
    $sia = $u.SignInActivity
    if (-not $sia) { return $null, 'Aucune' }
    if ($sia.LastSuccessfulSignInDateTime) { return [datetime]$sia.LastSuccessfulSignInDateTime, 'Succès' }
    $d = @($sia.LastSignInDateTime, $sia.LastNonInteractiveSignInDateTime) | Where-Object { $_ } |
         ForEach-Object { [datetime]$_ } | Sort-Object -Descending | Select-Object -First 1
    if ($d) { return $d, 'Tentative (succès inconnu)' }
    return $null, 'Aucune'
}
function Get-Statut($u, $der, $src) {
    if ($u.AccountEnabled -eq $false) { return 'Désactivé' }
    if ($sourceActivite -eq 'rapport') {
        if ($src -like 'Hors*') { return 'Activité inconnue' }
        if ($der) { if ($der -lt $dateSeuil) { return "Inactif > $MoisInactivite mois" } else { return 'Actif' } }
        if ($u.CreatedDateTime -and [datetime]$u.CreatedDateTime -gt $dateSeuil) { return 'Nouveau (jamais connecté)' }
        # Aucune activité sur 180 jours : inactif si le seuil ne dépasse pas 6 mois, sinon indéterminé
        if ($MoisInactivite -le 6) { return "Inactif > $MoisInactivite mois" } else { return 'Activité inconnue' }
    }
    if (-not $lireConnexions)         { return 'Activité inconnue' }
    if (-not $der) {
        if ($u.CreatedDateTime -and [datetime]$u.CreatedDateTime -gt $dateSeuil) { return 'Nouveau (jamais connecté)' }
        return 'Jamais connecté'
    }
    if ($der -lt $dateSeuil) { return "Inactif > $MoisInactivite mois" }
    return 'Actif'
}
function Get-Profil($u) {
    if ($u.UserType -eq 'Guest') { return 'Invité' }
    $upn = "$($u.UserPrincipalName)".ToLower()
    foreach ($k in $ReglesProfil.Keys) { if ($upn -match $ReglesProfil[$k]) { return $k } }
    if ("$($u.DisplayName)" -like 'svc*') { return 'Service / sync' }
    return 'Standard'
}

# --- 3. Groupes ---
Write-Host "Récupération des groupes..." -ForegroundColor Cyan
$groupes = Get-MgGroup -All -Property Id, DisplayName, GroupTypes, SecurityEnabled, MailEnabled, OnPremisesSyncEnabled, CreatedDateTime
$parIdGroupe = @{}; $groupes | ForEach-Object { $parIdGroupe[$_.Id] = $_ }
Write-Host "$($groupes.Count) groupes" -ForegroundColor Cyan

# Noms de groupes en double : suffixe court de l'Id (sinon fusionnés sur la carte)
$doublons = $groupes | Group-Object DisplayName | Where-Object Count -gt 1 | ForEach-Object Name
function Get-CleGroupe($g) { if ($g.DisplayName -in $doublons) { "$($g.DisplayName) [$($g.Id.Substring(0,8))]" } else { $g.DisplayName } }

# --- 4. Membres ---
$liens = [System.Collections.Generic.List[object]]::new()
$stats = @{}                                  # Id groupe -> compteurs
$nbGroupesUser = @{}                          # Id user -> nb de groupes directs
$proprios = @{}                               # Id groupe -> UPN des propriétaires
$lireProprios = -not $SansProprietaires
$i = 0
foreach ($g in $groupes) {
    $i++
    Write-Progress -Activity "Lecture des membres" -Status $g.DisplayName -PercentComplete ($i / $groupes.Count * 100)
    $s = @{ Sous = 0; Users = 0; Autres = 0 }
    try { $membres = Get-MgGroupMember -GroupId $g.Id -All -ErrorAction Stop }
    catch { Write-Warning "Membres illisibles pour '$($g.DisplayName)' : $($_.Exception.Message.Split("`n")[0])"; $stats[$g.Id] = $s; continue }
    foreach ($mb in $membres) {
        switch ($mb.AdditionalProperties.'@odata.type') {
            '#microsoft.graph.group' {
                $s.Sous++
                $enf = $parIdGroupe[$mb.Id]
                $liens.Add([PSCustomObject]@{ Parent = Get-CleGroupe $g; Enfant = if ($enf) { Get-CleGroupe $enf } else { $mb.AdditionalProperties.displayName }; TypeEnfant = 'Groupe' })
            }
            '#microsoft.graph.user' {
                $s.Users++
                if ($SansUtilisateurs) { continue }
                $u = $parIdUser[$mb.Id]
                $cle = if ($u) { $u.UserPrincipalName } else { $mb.AdditionalProperties.userPrincipalName }
                if (-not $cle) { continue }
                $nbGroupesUser[$mb.Id] = 1 + [int]$nbGroupesUser[$mb.Id]
                $liens.Add([PSCustomObject]@{ Parent = Get-CleGroupe $g; Enfant = $cle; TypeEnfant = 'Utilisateur' })
            }
            default { $s.Autres++ }   # appareils, principaux de service, contacts
        }
    }
    # Propriétaires : groupes cloud de sécurité ou Microsoft 365 uniquement
    if ($lireProprios -and -not $g.OnPremisesSyncEnabled -and ($g.SecurityEnabled -or $g.GroupTypes -contains 'Unified')) {
        try {
            $own = @(Get-MgGroupOwner -GroupId $g.Id -All -ErrorAction Stop | ForEach-Object {
                $o = $parIdUser[$_.Id]; if ($o) { $o.UserPrincipalName } elseif ($_.AdditionalProperties.userPrincipalName) { $_.AdditionalProperties.userPrincipalName } })
            $proprios[$g.Id] = $own
        } catch { $proprios[$g.Id] = $null }
    }
    $stats[$g.Id] = $s
}
Write-Progress -Activity "Lecture des membres" -Completed

# --- 4b. Rôles d'administration Entra (attributions actives, et éligibles si PIM) ---
$RolesPrivilegies = @('Global Administrator','Privileged Role Administrator','Privileged Authentication Administrator','Security Administrator',
    'Exchange Administrator','SharePoint Administrator','Teams Administrator','User Administrator','Authentication Administrator',
    'Helpdesk Administrator','Password Administrator','Application Administrator','Cloud Application Administrator',
    'Conditional Access Administrator','Intune Administrator','Hybrid Identity Administrator','Groups Administrator',
    'Billing Administrator','Compliance Administrator','Security Operator','Azure DevOps Administrator')
$roleNom = @{}; $roleStats = @{}; $lireRoles = -not $SansRoles
function Add-LienRole($roleId, $principalId, $type, $scope) {
    $nom = $roleNom[$roleId]; if (-not $nom) { return }
    $u = $parIdUser[$principalId]; $gp = $parIdGroupe[$principalId]
    if (-not $roleStats[$roleId]) { $roleStats[$roleId] = @{ Users = 0; Groupes = 0; Autres = 0 } }
    if ($scope -and $scope -ne '/') { $type = "$type (périmètre restreint)" }
    if ($u)      { $roleStats[$roleId].Users++;   $script:liens.Add([PSCustomObject]@{ Parent = "Rôle · $nom"; Enfant = $u.UserPrincipalName; TypeEnfant = 'Utilisateur'; Attribution = $type }) }
    elseif ($gp) { $roleStats[$roleId].Groupes++; $script:liens.Add([PSCustomObject]@{ Parent = "Rôle · $nom"; Enfant = (Get-CleGroupe $gp); TypeEnfant = 'Groupe'; Attribution = $type }) }
    else         { $roleStats[$roleId].Autres++ }   # principaux de service, identités managées
}
if ($lireRoles) {
    if (-not (Get-Command Get-MgRoleManagementDirectoryRoleDefinition -ErrorAction SilentlyContinue)) {
        Write-Warning "Module Microsoft.Graph.Identity.Governance absent : export sans rôles (Install-Module Microsoft.Graph.Identity.Governance -Scope CurrentUser)."
        $lireRoles = $false
    } elseif (-not (& $aScope @('RoleManagement.Read.Directory', 'RoleManagement.Read.All', 'Directory.Read.All'))) {
        Write-Warning "Scope RoleManagement.Read.Directory absent : export sans rôles."
        $lireRoles = $false
    } else {
        Write-Host "Lecture des rôles d'administration..." -ForegroundColor Cyan
        try {
            Get-MgRoleManagementDirectoryRoleDefinition -All -Property Id, DisplayName -ErrorAction Stop | ForEach-Object { $roleNom[$_.Id] = $_.DisplayName }
            # Instances PIM : distinguent permanent, temporaire et activé (sinon « Active »)
            $inst = @{}
            try { Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance -All -ErrorAction Stop | ForEach-Object { $inst["$($_.RoleDefinitionId)|$($_.PrincipalId)"] = $_ } } catch { }
            Get-MgRoleManagementDirectoryRoleAssignment -All -ErrorAction Stop | ForEach-Object {
                $i = $inst["$($_.RoleDefinitionId)|$($_.PrincipalId)"]
                $type = if (-not $i) { 'Active' } elseif ($i.AssignmentType -eq 'Activated') { 'Activée' } elseif ($i.EndDateTime) { 'Temporaire' } else { 'Permanente' }
                Add-LienRole $_.RoleDefinitionId $_.PrincipalId $type $_.DirectoryScopeId
            }
            try {
                Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All -ErrorAction Stop | ForEach-Object { Add-LienRole $_.RoleDefinitionId $_.PrincipalId 'Éligible' $_.DirectoryScopeId }
            } catch { Write-Warning "Attributions éligibles (PIM) non lues : licence Entra ID P2 ou droits manquants. Seules les attributions actives sont exportées." }
            Write-Host "$($roleStats.Count) rôles attribués" -ForegroundColor Cyan
        } catch {
            Write-Warning "Rôles illisibles ($($_.Exception.Message.Split("`n")[0])). Export sans rôles."
            $lireRoles = $false
        }
    }
}

# --- 5. Attributs (mêmes colonnes pour groupes et utilisateurs : Export-Csv prend celles de la 1re ligne) ---
function New-Ligne { [ordered]@{
    Cle=''; Libelle=''; TypeElement=''; Categorie=''; Profil=''; UPN=''; UserType=''; SyncAD=''; Dynamique=''
    DerniereConnexion=''; JoursInactivite=''; SourceDate=''; CreeLe=''; NbGroupesDirects=''
    NbMembresDirects=''; NbSousGroupes=''; NbUtilisateurs=''; NbAutres=''; Licences=''; LicencesSku=''
    MFA=''; MethodesMFA=''; RoleAdmin=''; Proprietaires=''; NbProprietaires=''; Id='' } }

$attributs = [System.Collections.Generic.List[object]]::new()
foreach ($g in $groupes) {
    $l = New-Ligne; $s = $stats[$g.Id]
    $l.Cle = Get-CleGroupe $g; $l.Libelle = $g.DisplayName; $l.TypeElement = 'Groupe'
    $l.Categorie = if ($g.GroupTypes -contains 'Unified') { 'M365' }
                   elseif ($g.SecurityEnabled -and $g.MailEnabled) { 'Sécurité (messagerie)' }
                   elseif ($g.SecurityEnabled) { 'Sécurité' } else { 'Distribution' }
    $l.SyncAD = [bool]$g.OnPremisesSyncEnabled; $l.Dynamique = $g.GroupTypes -contains 'DynamicMembership'
    $l.CreeLe = if ($g.CreatedDateTime) { ([datetime]$g.CreatedDateTime).ToString('yyyy-MM-dd') }
    if ($s) { $l.NbSousGroupes = $s.Sous; $l.NbUtilisateurs = $s.Users; $l.NbAutres = $s.Autres; $l.NbMembresDirects = $s.Sous + $s.Users + $s.Autres }
    if ($proprios.ContainsKey($g.Id) -and $null -ne $proprios[$g.Id]) { $l.Proprietaires = $proprios[$g.Id] -join ' | '; $l.NbProprietaires = $proprios[$g.Id].Count }
    $l.Id = $g.Id
    $attributs.Add([PSCustomObject]$l)
}
foreach ($rid in $roleStats.Keys) {
    $l = New-Ligne; $st = $roleStats[$rid]; $nom = $roleNom[$rid]
    $l.Cle = "Rôle · $nom"; $l.Libelle = $nom; $l.TypeElement = 'Rôle'
    $l.Categorie = if ($nom -in $RolesPrivilegies) { 'Rôle privilégié' } else { 'Rôle' }
    $l.NbUtilisateurs = $st.Users; $l.NbSousGroupes = $st.Groupes; $l.NbAutres = $st.Autres; $l.NbMembresDirects = $st.Users + $st.Groupes + $st.Autres
    $l.Id = $rid
    $attributs.Add([PSCustomObject]$l)
}
$compteStatut = @{}
if (-not $SansUtilisateurs) {
    foreach ($u in $parIdUser.Values) {
        $der, $src = Get-DerniereActivite $u
        $l = New-Ligne
        $l.Cle = $u.UserPrincipalName; $l.Libelle = $u.DisplayName; $l.TypeElement = 'Utilisateur'
        $l.Categorie = Get-Statut $u $der $src; $l.Profil = Get-Profil $u
        $l.UPN = $u.UserPrincipalName; $l.UserType = $u.UserType; $l.SyncAD = [bool]$u.OnPremisesSyncEnabled
        if ($der) { $l.DerniereConnexion = $der.ToString($fmt); $l.JoursInactivite = [int]($maintenant - $der).TotalDays }
        $l.SourceDate = if ($sourceActivite) { $src } else { 'Non lue' }
        $l.CreeLe = if ($u.CreatedDateTime) { ([datetime]$u.CreatedDateTime).ToString('yyyy-MM-dd') }
        $l.NbGroupesDirects = [int]$nbGroupesUser[$u.Id]
        if ($lireLicences) { $l.Licences, $l.LicencesSku = Get-Licences $u }
        if ($lireMFA) {
            $r = $mfaParId[$u.Id]
            if ($r) {
                $l.MFA = if ($r.IsMfaRegistered) { 'Oui' } else { 'Non' }
                $l.MethodesMFA = @($r.MethodsRegistered) -join ' | '
                $l.RoleAdmin = if ($r.IsAdmin) { 'Oui' } else { 'Non' }
            }
        }
        $l.Id = $u.Id
        $attributs.Add([PSCustomObject]$l)
        $cleStat = "$($l.Categorie)|$($u.UserType)"
        $compteStatut[$cleStat] = 1 + [int]$compteStatut[$cleStat]
    }
}

# --- 6. Export ---
$encodage = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
$horodatage = Get-Date -Format 'yyyyMMdd_HHmm'
$fLiens = Join-Path $DossierExport "$($Prefixe)_$($horodatage)_MapTools_Liens.csv"
$fAttr  = Join-Path $DossierExport "$($Prefixe)_$($horodatage)_MapTools_Attributs.csv"
$liens | Select-Object Parent, Enfant, TypeEnfant, Attribution | Export-Csv $fLiens -NoTypeInformation -Delimiter ';' -Encoding $encodage
$attributs | Sort-Object TypeElement, Cle | Export-Csv $fAttr -NoTypeInformation -Delimiter ';' -Encoding $encodage

# --- 7. Résumé ---
Write-Host "`n=== Résumé ===" -ForegroundColor Green
Write-Host "Groupes : $($groupes.Count) | liens groupe → groupe : $(@($liens | Where-Object TypeEnfant -eq 'Groupe').Count)"
if (-not $SansUtilisateurs) {
    $total = $parIdUser.Count
    Write-Host "Utilisateurs : $total | appartenances : $(@($liens | Where-Object TypeEnfant -eq 'Utilisateur').Count)"
    Write-Host "Seuil d'inactivité : $MoisInactivite mois (avant le $($dateSeuil.ToString('dd.MM.yyyy')))"
    $compteStatut.GetEnumerator() | ForEach-Object {
        $st, $ty = $_.Key -split '\|'
        [PSCustomObject]@{ Statut = $st; Type = $ty; Nombre = $_.Value; 'Pct total' = '{0:P1}' -f ($_.Value / [math]::Max(1,$total)) }
    } | Sort-Object Type, Statut | Format-Table -AutoSize
    $actives = @($attributs | Where-Object { $_.TypeElement -eq 'Utilisateur' -and $_.Categorie -ne 'Désactivé' })
    if ($lireMFA) {
        $comptesSansMfa = @($actives | Where-Object { $_.MFA -eq 'Non' })
        Write-Host "Comptes activés sans MFA enregistrée : $($comptesSansMfa.Count) sur $($actives.Count) (dont $(@($comptesSansMfa | Where-Object { $_.RoleAdmin -eq 'Oui' }).Count) avec un rôle admin)"
    }
    if ($lireRoles) {
        $ga = @($liens | Where-Object { $_.Parent -eq 'Rôle · Global Administrator' })
        Write-Host "Rôles attribués : $($roleStats.Count) | attributions Global Administrator : $($ga.Count) (dont $(@($ga | Where-Object { $_.Attribution -like 'Éligible*' }).Count) éligibles)"
    }
    if ($lireLicences) {
        $licDormants = @($attributs | Where-Object { $_.TypeElement -eq 'Utilisateur' -and $_.Licences -and $_.Categorie -match 'Désactivé|Inactif|Jamais' })
        Write-Host "Comptes inactifs ou désactivés avec licence : $($licDormants.Count)"
    }
}
Write-Host "Liens      : $fLiens" -ForegroundColor Green
Write-Host "Attributs  : $fAttr" -ForegroundColor Green
Write-Host "Dépose les deux fichiers dans MapTools.html" -ForegroundColor Green

# --- 8. Partages (facultatif) : n'interrompt jamais l'export principal ---
if ($AvecPartages) {
    $scriptPartages = Join-Path $PSScriptRoot 'Export-MapToolsPartages.ps1'
    if (-not (Test-Path $scriptPartages)) { Write-Warning "Export-MapToolsPartages.ps1 introuvable à côté de ce script : partages non exportés." }
    else {
        $pp = @{ Prefixe = $Prefixe; DossierExport = $DossierExport; Horodatage = $horodatage }
        if ($SharePointAdminUrl) { $pp.SharePointAdminUrl = $SharePointAdminUrl }
        try { $init = (Get-MgOrganization -ErrorAction Stop).VerifiedDomains | Where-Object IsInitial | Select-Object -First 1; if ($init) { $pp.Organisation = $init.Name } } catch { }
        try { & $scriptPartages @pp } catch { Write-Warning "Export des partages interrompu ($($_.Exception.Message.Split("`n")[0])). L'export principal est complet." }
    }
}

if ($connexionParScript) { Disconnect-MgGraph | Out-Null }
