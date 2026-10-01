<#
.SYNOPSIS
    Exporte les partages SharePoint / OneDrive (journal d'audit Purview) et les paramètres
    de partage externe des sites, au format attendu par MapTools.html.

.DESCRIPTION
    Deux CSV dans -DossierExport :
      - <prefixe>_<date>_MapTools_Partages.csv : un partage par ligne (lien « Tout le monde »,
        lien organisation, lien personnes spécifiques, invitation, accès accordé)
      - <prefixe>_<date>_MapTools_Sites.csv    : paramètre de partage externe de chaque site
        (et une ligne pour le paramètre du tenant)

    Le journal d'audit montre les partages CRÉÉS pendant la période (180 jours au plus avec
    Audit Standard), pas l'état actuel des liens : un lien créé avant la période n'apparaît pas,
    un lien créé puis supprimé est signalé (colonne SupprimeEnsuite).

    Lecture seule. Chaque partie est facultative : si un module ou un droit manque, elle est
    signalée et ignorée, l'autre est exportée.

.PREREQUIS
    Audit : module ExchangeOnlineManagement (v3), rôle View-Only Audit Logs, Audit Logs ou Global Reader ;
            audit unifié activé (par défaut). En accès délégué : -Organisation <domaine du client>.
    Sites : module Microsoft.Online.SharePoint.PowerShell, rôle SharePoint Administrator.
            Sous PowerShell 7, le module est chargé avec -UseWindowsPowerShell (Windows uniquement).
    Install-Module ExchangeOnlineManagement, Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser

.EXAMPLE
    .\Export-MapToolsPartages.ps1 -Prefixe Contoso -DossierExport C:\Temp -SharePointAdminUrl https://contoso-admin.sharepoint.com
    .\Export-MapToolsPartages.ps1 -Prefixe Contoso -Organisation contoso.onmicrosoft.com -Jours 90 -SansSites

.NOTES
    Outil     : MapTools 1.1
    Auteur    : Jimmy Paul
    Copyright : © 2026 Jimmy Paul. Voir LICENCE.txt.
#>

[CmdletBinding()]
param(
    [string]$DossierExport = ".",
    [string]$Prefixe = "Entra",
    # Horodatage imposé (utilisé par Export-MapToolsEntra.ps1 -AvecPartages pour aligner les noms)
    [string]$Horodatage,
    # Domaine du tenant client, pour Connect-ExchangeOnline -DelegatedOrganization (accès délégué)
    [string]$Organisation,
    [string]$SharePointAdminUrl,
    [ValidateRange(1, 180)][int]$Jours = 180,
    [switch]$AvecOneDrive,
    [switch]$AvecUtilisations,
    [switch]$SansAudit,
    [switch]$SansSites
)

Write-Host "MapTools 1.1 · export des partages" -ForegroundColor DarkCyan

if (-not $Horodatage) { $Horodatage = Get-Date -Format 'yyyyMMdd_HHmm' }
$encodage = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }

# --- Domaines du tenant (pour reconnaître un destinataire externe) ---
$domaines = @()
if (Get-Command Get-MgDomain -ErrorAction SilentlyContinue) {
    try { $domaines = @(Get-MgDomain -All -ErrorAction Stop | ForEach-Object { $_.Id.ToLower() }) } catch { }
}

# ===================== 1. Journal d'audit : partages =====================
$TypesPartage = [ordered]@{
    'AnonymousLinkCreated'     = 'Lien « Tout le monde » créé'
    'AnonymousLinkUpdated'     = 'Lien « Tout le monde » modifié'
    'AnonymousLinkRemoved'     = 'Lien « Tout le monde » supprimé'
    'CompanyLinkCreated'       = 'Lien organisation créé'
    'SecureLinkCreated'        = 'Lien personnes spécifiques créé'
    'AddedToSecureLink'        = 'Ajout à un lien personnes spécifiques'
    'SharingInvitationCreated' = 'Invitation de partage'
    'SharingSet'               = 'Accès accordé'
}
if ($AvecUtilisations) { $TypesPartage['AnonymousLinkUsed'] = 'Lien « Tout le monde » utilisé' }

$partages = [System.Collections.Generic.List[object]]::new()
if (-not $SansAudit) {
    if (-not (Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue)) {
        Write-Warning "Module ExchangeOnlineManagement absent : partages non exportés (Install-Module ExchangeOnlineManagement -Scope CurrentUser)."
    } else {
        $connecte = $false
        try { $connecte = [bool](Get-ConnectionInformation -ErrorAction Stop | Where-Object { $_.State -eq 'Connected' }) } catch { }
        if (-not $connecte) {
            $p = @{ ShowBanner = $false }
            if ($Organisation) { $p.DelegatedOrganization = $Organisation }
            try { Connect-ExchangeOnline @p -ErrorAction Stop; $connecte = $true }
            catch { Write-Warning "Connexion Exchange Online impossible ($($_.Exception.Message.Split("`n")[0])) : partages non exportés." }
        }
        if ($connecte) {
            $fin = Get-Date; $debut = $fin.AddDays(-$Jours)
            Write-Host "Lecture du journal d'audit ($Jours jours)..." -ForegroundColor Cyan
            $vus = @{}; $erreur = $false
            # Fenêtres de 7 jours : ReturnLargeSet plafonne à 50 000 résultats par session
            for ($d = $debut; $d -lt $fin -and -not $erreur; $d = $d.AddDays(7)) {
                $f = if ($d.AddDays(7) -gt $fin) { $fin } else { $d.AddDays(7) }
                $session = "MapTools_$([guid]::NewGuid())"
                Write-Progress -Activity "Journal d'audit" -Status "$($d.ToString('dd.MM.yyyy')) → $($f.ToString('dd.MM.yyyy'))" -PercentComplete ((($d - $debut).TotalDays / $Jours) * 100)
                do {
                    try {
                        $lot = @(Search-UnifiedAuditLog -StartDate $d -EndDate $f -RecordType SharePointSharingOperation -Operations @($TypesPartage.Keys) `
                                 -SessionId $session -SessionCommand ReturnLargeSet -ResultSize 5000 -ErrorAction Stop)
                    } catch {
                        Write-Warning "Recherche d'audit impossible ($($_.Exception.Message.Split("`n")[0])). Vérifier le rôle View-Only Audit Logs et que l'audit est activé."
                        $erreur = $true; break
                    }
                    foreach ($e in $lot) {
                        if ($vus.ContainsKey($e.Identity)) { continue }; $vus[$e.Identity] = $true
                        $a = $e.AuditData | ConvertFrom-Json
                        $dest = "$($a.TargetUserOrGroupName)"; $typeDest = "$($a.TargetUserOrGroupType)"
                        $externe = if ($a.Operation -like 'AnonymousLink*') { 'Tout le monde' }
                                   elseif ($typeDest -eq 'Guest' -or $dest -match '#ext#') { 'Oui' }
                                   elseif ($dest -match '@([^@]+)$' -and $domaines.Count -and ($Matches[1].ToLower() -notin $domaines)) { 'Oui' }
                                   else { 'Non' }
                        $partages.Add([PSCustomObject]@{
                            Date             = ([datetime]$a.CreationTime).ToString('yyyy-MM-dd HH:mm')
                            Operation        = $a.Operation
                            TypePartage      = $TypesPartage[$a.Operation]
                            PartagePar       = $a.UserId
                            Destinataire     = $dest
                            TypeDestinataire = $typeDest
                            Externe          = $externe
                            NomElement       = if ($a.SourceFileName) { $a.SourceFileName } else { ("$($a.ObjectId)" -split '/')[-1] }
                            Element          = $a.ObjectId
                            TypeElementPartage = $a.ItemType
                            Site             = $a.SiteUrl
                            Service          = $a.Workload
                            Permission       = $a.Permission
                            SupprimeEnsuite  = ''
                        })
                    }
                } while ($lot.Count -gt 0 -and -not $erreur)
            }
            Write-Progress -Activity "Journal d'audit" -Completed
            # Lien « Tout le monde » supprimé après sa création : à signaler
            $suppr = @{}
            $partages | Where-Object Operation -eq 'AnonymousLinkRemoved' | ForEach-Object { $k = "$($_.Element)"; if (-not $suppr[$k] -or $_.Date -gt $suppr[$k]) { $suppr[$k] = $_.Date } }
            $partages | Where-Object Operation -eq 'AnonymousLinkCreated' | ForEach-Object { $_.SupprimeEnsuite = if ($suppr["$($_.Element)"] -gt $_.Date) { 'Oui' } else { 'Non' } }
            Write-Host "$($partages.Count) événements de partage" -ForegroundColor Cyan
        }
    }
}

# ===================== 2. Paramètres de partage des sites =====================
$LibellesPartage = @{
    'Disabled'                        = 'Désactivé (interne uniquement)'
    'ExistingExternalUserSharingOnly' = 'Invités existants uniquement'
    'ExternalUserSharingOnly'         = 'Nouveaux et existants invités'
    'ExternalUserAndGuestSharing'     = 'Tout le monde (liens anonymes)'
}
$sites = [System.Collections.Generic.List[object]]::new()
if (-not $SansSites) {
    $spoOk = $false
    if (-not (Get-Command Get-SPOSite -ErrorAction SilentlyContinue)) {
        try {
            if ($PSVersionTable.PSVersion.Major -ge 6) { Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell -WarningAction SilentlyContinue -ErrorAction Stop }
            else { Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction Stop }
        } catch { }
        if (-not (Get-Command Get-SPOSite -ErrorAction SilentlyContinue)) {
            Write-Warning "Module Microsoft.Online.SharePoint.PowerShell indisponible (Install-Module Microsoft.Online.SharePoint.PowerShell ; sous PowerShell 7, Windows uniquement) : paramètres des sites non exportés."
        }
    }
    if (Get-Command Get-SPOSite -ErrorAction SilentlyContinue) {
        try { Get-SPOTenant -ErrorAction Stop | Out-Null; $spoOk = $true } catch {
            if (-not $SharePointAdminUrl -and (Get-Command Get-MgOrganization -ErrorAction SilentlyContinue)) {
                try { $init = (Get-MgOrganization -ErrorAction Stop).VerifiedDomains | Where-Object IsInitial | Select-Object -First 1
                      if ($init) { $SharePointAdminUrl = "https://$($init.Name.Split('.')[0])-admin.sharepoint.com" } } catch { }
            }
            if ($SharePointAdminUrl) {
                try { Connect-SPOService -Url $SharePointAdminUrl -ErrorAction Stop; $spoOk = $true }
                catch { Write-Warning "Connexion SharePoint impossible ($($_.Exception.Message.Split("`n")[0])). Rôle SharePoint Administrator requis." }
            } else { Write-Warning "Indiquer -SharePointAdminUrl (https://<tenant>-admin.sharepoint.com) : paramètres des sites non exportés." }
        }
    }
    if ($spoOk) {
        Write-Host "Lecture des paramètres de partage des sites..." -ForegroundColor Cyan
        try {
            $t = Get-SPOTenant -ErrorAction Stop
            $sites.Add([PSCustomObject]@{ Url = 'Tenant'; Titre = 'Paramètre du tenant'; Modele = 'Tenant'; PartageExterne = $LibellesPartage["$($t.SharingCapability)"]
                ValeurBrute = "$($t.SharingCapability)"; GroupeM365 = ''; Proprietaire = ''; DerniereModification = ''; StockageMo = ''
                ExpirationLiensAnonymesJours = $t.RequireAnonymousLinksExpireInDays; LienParDefaut = "$($t.DefaultSharingLinkType)" })
            Get-SPOSite -Limit All -IncludePersonalSite ([bool]$AvecOneDrive) -ErrorAction Stop | ForEach-Object {
                $sites.Add([PSCustomObject]@{ Url = $_.Url; Titre = $_.Title; Modele = $_.Template; PartageExterne = $LibellesPartage["$($_.SharingCapability)"]
                    ValeurBrute = "$($_.SharingCapability)"; GroupeM365 = if ("$($_.GroupId)" -ne '00000000-0000-0000-0000-000000000000') { "$($_.GroupId)" } else { '' }
                    Proprietaire = $_.Owner; DerniereModification = if ($_.LastContentModifiedDate) { ([datetime]$_.LastContentModifiedDate).ToString('yyyy-MM-dd') } else { '' }
                    StockageMo = $_.StorageUsageCurrent; ExpirationLiensAnonymesJours = ''; LienParDefaut = '' })
            }
            Write-Host "$($sites.Count - 1) sites" -ForegroundColor Cyan
        } catch { Write-Warning "Lecture des sites impossible ($($_.Exception.Message.Split("`n")[0]))." }
    }
}

# ===================== Export et résumé =====================
if ($partages.Count) {
    $fP = Join-Path $DossierExport "$($Prefixe)_$($Horodatage)_MapTools_Partages.csv"
    $partages | Sort-Object Date | Export-Csv $fP -NoTypeInformation -Delimiter ';' -Encoding $encodage
    $anon = @($partages | Where-Object Operation -eq 'AnonymousLinkCreated')
    $ext  = @($partages | Where-Object { $_.Externe -eq 'Oui' })
    Write-Host "`nPartages : $($partages.Count) | liens « Tout le monde » créés : $($anon.Count) (dont $(@($anon | Where-Object SupprimeEnsuite -eq 'Oui').Count) supprimés ensuite) | vers des externes : $($ext.Count)" -ForegroundColor Green
    Write-Host "Partages : $fP" -ForegroundColor Green
} elseif (-not $SansAudit) { Write-Host "Aucun événement de partage exporté." -ForegroundColor Yellow }
if ($sites.Count) {
    $fS = Join-Path $DossierExport "$($Prefixe)_$($Horodatage)_MapTools_Sites.csv"
    $sites | Export-Csv $fS -NoTypeInformation -Delimiter ';' -Encoding $encodage
    $ouverts = @($sites | Where-Object { $_.ValeurBrute -eq 'ExternalUserAndGuestSharing' -and $_.Url -ne 'Tenant' })
    Write-Host "Sites autorisant les liens « Tout le monde » : $($ouverts.Count) sur $($sites.Count - 1)" -ForegroundColor Green
    Write-Host "Sites    : $fS" -ForegroundColor Green
}
if ($partages.Count -or $sites.Count) { Write-Host "Dépose ces fichiers dans MapTools.html avec les autres exports" -ForegroundColor Green }
