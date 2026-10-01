# MapTools · Référence technique

Version 1.1 · © 2026 Jimmy Paul

Document de consultation pour adapter MapTools ou l'utiliser avec d'autres sources. L'utilisation courante est décrite dans `GUIDE.md`.

**Sommaire** : 1. Paramètres des scripts · 2. Fichiers produits · 3. Formats CSV pour d'autres sources · 4. Réglages · 5. Pièges connus · 6. Intégrité et signature · 7. Historique · 8. Composants tiers

## 1. Paramètres des scripts

### `Export-MapToolsEntra.ps1`

| Paramètre | Rôle | Défaut |
|---|---|---|
| `-TenantId` | Tenant à utiliser si aucune session Graph n'est ouverte | tenant du compte |
| `-DossierExport` | Dossier de sortie | dossier courant |
| `-Prefixe` | Préfixe des fichiers | `Entra` |
| `-MoisInactivite` | Seuil d'inactivité en mois | `6` |
| `-SansUtilisateurs` | Groupes et imbrications uniquement | désactivé |
| `-SansMFA` | Ne pas lire le statut MFA | désactivé |
| `-SansLicences` | Ne pas lire les licences | désactivé |
| `-SansRoles` | Ne pas lire les rôles d'administration | désactivé |
| `-SansProprietaires` | Ne pas lire les propriétaires de groupes | désactivé |

Modules : Microsoft.Graph.Authentication, .Users, .Groups, .Reports (MFA), .Identity.DirectoryManagement (noms des licences) et .Identity.Governance (rôles).

Scopes : `GroupMember.Read.All` (ou `Group.Read.All`, `Directory.Read.All`), `User.Read.All` (ou `Directory.Read.All`), `AuditLog.Read.All` (dates de connexion et MFA), `LicenseAssignment.Read.All` (abonnements), `RoleManagement.Read.Directory` (rôles ; attributions éligibles avec Entra ID P2), `Reports.Read.All` (secours sans P1 : rapport d'activité Microsoft 365). Si une session existe sans les scopes nécessaires à la lecture des groupes ou des utilisateurs, le script s'arrête et indique la commande de reconnexion.

### `Export-MapToolsAD.ps1`

| Paramètre | Rôle | Défaut |
|---|---|---|
| `-Server` | Contrôleur de domaine | découvert automatiquement |
| `-SearchBase` | Limiter les groupes exportés à une OU (les imbrications restent résolues sur tout le domaine) | tout le domaine |
| `-DossierExport` | Dossier de sortie | dossier courant |
| `-Prefixe` | Préfixe des fichiers | `AD` |
| `-MoisInactivite` | Seuil d'inactivité en mois | `6` |
| `-SansUtilisateurs` | Groupes et imbrications uniquement | désactivé |

### `Export-MapToolsPartages.ps1`

| Paramètre | Rôle | Défaut |
|---|---|---|
| `-DossierExport`, `-Prefixe` | Dossier et préfixe des fichiers | dossier courant, `Entra` |
| `-Horodatage` | Horodatage imposé (aligné sur l'export Entra par `-AvecPartages`) | date et heure courantes |
| `-Organisation` | Domaine du tenant pour `Connect-ExchangeOnline -DelegatedOrganization` (accès délégué) | aucun |
| `-SharePointAdminUrl` | URL d'administration SharePoint | déduite du domaine initial |
| `-Jours` | Période d'audit (1 à 180) | `180` |
| `-AvecOneDrive` | Inclure les sites OneDrive dans la liste des sites | désactivé |
| `-AvecUtilisations` | Ajouter les utilisations de liens anonymes (`AnonymousLinkUsed`, volumineux) | désactivé |
| `-SansAudit`, `-SansSites` | Ne pas lire le journal d'audit, ou les sites | désactivé |

Côté export Entra : `-AvecPartages` (et `-SharePointAdminUrl`) lancent ce script à la fin.

Opérations d'audit lues (type d'enregistrement `SharePointSharingOperation`) : AnonymousLinkCreated, AnonymousLinkUpdated, AnonymousLinkRemoved, CompanyLinkCreated, SecureLinkCreated, AddedToSecureLink, SharingInvitationCreated, SharingSet. La recherche se fait par fenêtres de 7 jours, car `Search-UnifiedAuditLog -SessionCommand ReturnLargeSet` plafonne à 50 000 résultats par session.

### `Get-EntraIDComptesInactifs.ps1`

| Paramètre | Rôle | Défaut |
|---|---|---|
| `-TenantId` | Tenant à utiliser si aucune session Graph n'est ouverte | tenant du compte |
| `-MoisInactivite` | Seuil d'inactivité en mois | `6` |
| `-DossierExport` | Dossier de sortie | dossier courant |
| `-InclureInvites` | Inclure les comptes invités | désactivé |

## 2. Fichiers produits

Séparateur `;`, UTF-8 avec BOM, noms horodatés `<Prefixe>_<AAAAMMJJ_HHmm>_MapTools_Liens.csv` et `…_MapTools_Attributs.csv` (`…_AD_MapTools_…` pour l'AD). Le préfixe sert de nom de client dans les rapports, l'horodatage de date dans la comparaison avant / après : garder ces noms tels quels.

**Liens** : `Parent ; Enfant ; TypeEnfant ; Attribution`. TypeEnfant vaut Groupe ou Utilisateur. Attribution n'est renseignée que pour les rôles : Permanente, Active (sans détail PIM), Temporaire, Activée ou Éligible, suivi de « (périmètre restreint) » pour une unité administrative.

**Attributs** : une ligne par groupe et par utilisateur, mêmes colonnes pour tous.

| Colonne | Contenu |
|---|---|
| `Cle` | Groupe : nom (suffixé si homonyme). Utilisateur : UPN (SamAccountName à défaut côté AD) |
| `Libelle` | Nom affiché sur la carte |
| `TypeElement` | Groupe ou Utilisateur |
| `Categorie` | Groupe : type. Utilisateur : statut |
| `Profil` | Profil probable de l'utilisateur (section 4.2) |
| `DerniereConnexion`, `JoursInactivite` | Date de référence et ancienneté en jours |
| `CreeLe` | Date de création |
| `NbGroupesDirects` | Utilisateur : nombre de groupes dont il est membre direct |
| `NbMembresDirects`, `NbSousGroupes`, `NbUtilisateurs`, `NbAutres` | Groupe : composition directe (Autres = appareils, principaux de service, contacts, membres d'autres domaines) |

Colonnes propres à Entra : `UPN`, `UserType`, `SyncAD`, `Dynamique`, `SourceDate` (Succès, Tentative (succès inconnu), Aucune, Non lue ; en secours : Rapport d'activité M365, Rapport d'activité M365 (aucune activité sur 180 j), Hors rapport d'activité (sans licence)), `Licences` (noms commerciaux, séparés par « | »), `LicencesSku` (références techniques), `MFA` (Oui, Non, vide si non lu), `MethodesMFA`, `RoleAdmin` (Oui si un rôle d'administration Entra est attribué), `Proprietaires` (UPN séparés par « | »), `NbProprietaires` (vide si non lus : groupes synchronisés, listes de distribution), `Id`.

Les rôles sont des lignes à part : `Cle` = « Rôle · <nom> », `TypeElement` = Rôle, `Categorie` = Rôle privilégié ou Rôle selon la table `$RolesPrivilegies` du script.

Le statut MFA provient du rapport d'enregistrement des méthodes (`userRegistrationDetails`) : « Oui » signifie qu'une méthode forte est **enregistrée**, pas que la MFA est **exigée** à chaque connexion ; c'est l'accès conditionnel qui l'impose. Les noms commerciaux des licences courantes sont définis dans la table `$NomsSku` du script.
Colonnes propres à l'AD : `UPN`, `SamAccountName`, `Protege` (adminCount = 1), `Etendue`, `MdpNExpireJamais`, `MdpDernierChangement`, `OU`.

## 3. Formats CSV pour d'autres sources

MapTools accepte toute hiérarchie parent → enfant : organigramme (responsable → collaborateur), partages de fichiers (dossier → sous-dossier), dépendances (application → dépend de), rôles (rôle → groupe éligible).

Le séparateur (`;`, `,`, tabulation) et l'encodage UTF-8 avec ou sans BOM sont détectés. La première ligne contient les noms de colonnes.

### Fichier de liens (obligatoire)

Une ligne par relation, deux colonnes suffisent. Les doublons et les lignes incomplètes sont ignorés.

```
Parent;Enfant
GG_Finance;GG_Comptabilite
GG_Comptabilite;marie.dupont@contoso.ch
```

### Fichier d'attributs (optionnel)

Une ligne par élément ; la colonne clé doit contenir exactement les valeurs utilisées dans les liens.

```
Cle;Libelle;TypeElement;Categorie;SyncAD
GG_Finance;GG_Finance;Groupe;Sécurité;True
marie.dupont@contoso.ch;Marie Dupont;Utilisateur;Actif;True
```

| Rôle de colonne | Effet |
|---|---|
| Clé | Jointure avec les liens (obligatoire) |
| Libellé affiché | Texte affiché à la place de la clé |
| Type d'élément | Valeur contenant user, utilisateur, compte, account, person ou contact : l'élément est un utilisateur (point) ; sinon un groupe |
| Catégorie / statut | Couleur (statuts à couleur fixe, sinon 8 couleurs par fréquence puis gris) |
| Pointillés si vrai | Bordure en pointillés des groupes |
| MFA (vrai / faux) | Utilisateur activé à faux : dessiné en losange ; alimente les constats MFA du rapport |
| Ovale si vrai | Forme ovale des groupes |

Valeurs lues comme vraies : `True`, `Vrai`, `Oui`, `Yes`, `1`, `Y`, `O`. Toutes les colonnes non vides s'affichent dans le panneau de détails.

### Colonnes reconnues automatiquement

Sans tenir compte de la casse ni des accents.

| Rôle | Noms reconnus |
|---|---|
| Parent | Parent, ParentGroup, Groupe parent, Source, From, De, Manager, Responsable |
| Enfant | Enfant, Child, ChildGroup, Member, Membre, Sous-groupe, Target, To, Vers, Cible |
| Clé | Cle, Key, DisplayName, Name, Nom, Libelle, Label, Element, SamAccountName, Groupe, Group, Id (la première présente dans l'ordre du fichier) |
| Libellé affiché | Libelle, Label, DisplayName, NomAffiche (autre que la clé) |
| Type d'élément | TypeElement, TypeObjet, ObjectType, Kind, Nature |
| Catégorie | Categorie, Category, Type, GroupCategory, GroupType, Statut, Status |
| Pointillés | SyncAD, OnPremisesSyncEnabled, Synced, Synchronise, Protege, adminCount |
| Ovale | Dynamique, Dynamic, IsDynamic |
| MFA | MFA, MfaEnregistre, IsMfaRegistered, MfaRegistered |
| Type de lien (fichier de liens) | Attribution, TypeLien, TypeAttribution, Assignment, AssignmentType |

Un fichier avec une colonne parent et une colonne enfant est pris comme liens ; un fichier avec une colonne clé, comme attributs. Tout se corrige dans la fiche du fichier.

## 4. Réglages

### 4.1 Outil (`MapTools.html`)

Constantes de la section « Configuration », en tête du script :

| Constante | Rôle | Défaut |
|---|---|---|
| `ALIAS` | Noms de colonnes reconnus automatiquement | section 3 |
| `USER_KIND` | Valeurs de « Type d'élément » traitées comme utilisateurs | user, utilisateur, compte… |
| `ROLE_KIND` | Valeurs de « Type d'élément » traitées comme rôles d'administration | Rôle, Role, DirectoryRole |
| `LINK_PRIO` | Attribution retenue quand un lien apparaît plusieurs fois | Permanente > Active > Temporaire > Activée > Éligible |
| `PWD_MAX_DAYS` | Ancienneté maximale d'un mot de passe dans le rapport (export AD) | 365 jours |
| `SEMANTIC` | Couleurs fixes par motif de catégorie | Actif, Nouveau, Jamais, Inactif, Désactivé, Inconnu |
| `WHEEL_K` | Sensibilité de la molette (augmenter pour zoomer plus vite) | 0.0006 (≈ 6 % par cran) |
| `ZOOM_MIN`, `ZOOM_MAX` | Bornes du zoom | 0.08 et 2.5 |
| `ZOOM_STEP` | Facteur des boutons et touches +/− | 1.3 |
| `FIT_MAX` | Zoom maximal en « tout afficher » | 1.1 |
| `SPACING` | Réglages des trois espacements : écart entre niveaux (`lvl`), entre éléments d'un niveau (`node`) et entre arbres (`tree`), taille du texte, longueur des libellés. `node` et `tree` se mesurent entre centres et doivent dépasser la hauteur d'un élément | Compact 175/30/32, Normal 250/40/42, Aéré 300/52/64 |
| `FLAT_AUTO_HIDE` | Nombre de groupes sans imbrication au-delà duquel ils sont masqués au chargement | 100 |
| `AUTO_FREEZE` | Vue Réseau figée après calcul de la disposition | true |
| `MAX_LABEL` | Longueur maximale d'un libellé | 32 caractères |
| `PALETTE` | Couleurs de catégorie avant le gris | 8 |
| `BIG_GRAPH` | Seuil d'avertissement (éléments affichés) | 800 |
| `MAX_LIST` | Éléments listés dans le panneau de détails | 300 |
| `DATA_F` | Constats classés dans la partie « Partages et sites » | voir le code |
| `GROUP_F`, `PRIV_F` | Constats classés dans la partie « Groupes », et constats comptés dans la composante « accès à privilèges » de l'indice | voir le code |
| `SENSITIVE` | Motif des noms de groupes considérés comme sensibles dans le rapport et la vue rapide « Périmètre d'un groupe à privilèges » (en plus des groupes marqués « Protege ») | admin, adm_, pim, privil, domain/enterprise/schema admins, account/backup operators, break glass, exclusion, tier 0 |

Les couleurs sont des variables CSS (`--c1` à `--c8`, `--s-ok`, `--s-warn`…), avec une variante pour le thème sombre. Les données sont mémorisées dans la base IndexedDB `maptools` (clé `courant`) ; les prix des licences saisis dans le rapport sous la clé `prix`.

### 4.2 Scripts : règles de profil

La variable `$ReglesProfil`, en tête des deux scripts d'export, associe à chaque profil une expression régulière testée dans l'ordre (sur l'UPN pour Entra, sur le SamAccountName pour l'AD). Les invités sont toujours classés « Invité » ; ce qui ne correspond à aucune règle est « Standard ».

| Profil | Règle Entra (UPN) | Règle AD (SamAccountName) |
|---|---|---|
| Service / sync | `^(svc[-_.]\|sync_\|adsyncsvc)\|sbc` | `^(svc[-_.]\|sa[-_.]\|msol_\|aad_\|sync_)` |
| Admin | `-a2@\|-admin@\|^adm[-_.]\|^admin[-_.]` | `-a2$\|-admin$\|^adm[-_.]\|^admin[-_.]` |
| Salle / équipement | `^(salle\|room\|booking\|wireless\|ipad\|tv\.\|parking\|massage\|scan)` | `^(salle\|room\|scan\|copieur\|printer)` |
| Technique / test | `^(test\|install\|no-?reply\|dmarc\|audit\d\|intune\.)` | `^(test\|install\|krbtgt$\|guest$\|defaultaccount$)` |

Adapter ces règles aux conventions de nommage de chaque client.

### Partages et sites

**Partages** (`…_MapTools_Partages.csv`) : `Date`, `Operation`, `TypePartage`, `PartagePar`, `Destinataire`, `TypeDestinataire`, `Externe` (Oui, Non, Tout le monde), `NomElement`, `Element` (URL), `TypeElementPartage`, `Site`, `Service`, `Permission`, `SupprimeEnsuite` (lien « Tout le monde » supprimé plus tard dans la période).

Un destinataire est « externe » s'il est de type invité, si son identifiant contient `#EXT#`, ou si son domaine n'est pas un domaine du tenant (lu dans la session Graph si elle est ouverte). Les invités sont rapprochés des comptes de l'annuaire par leur UPN ou leur adresse d'origine (`prenom.nom_domaine.ch#EXT#@…` ↔ `prenom.nom@domaine.ch`).

**Sites** (`…_MapTools_Sites.csv`) : `Url` (« Tenant » pour la ligne du tenant), `Titre`, `Modele`, `PartageExterne` (libellé), `ValeurBrute` (SharingCapability), `GroupeM365`, `Proprietaire`, `DerniereModification`, `StockageMo`, `ExpirationLiensAnonymesJours` et `LienParDefaut` (tenant).

### Listes exportées

Séparateur `;`, UTF-8 avec BOM, une ligne d'en-tête avec les noms de colonnes choisis. Nom : `<client>_<liste>_<date>.csv`. Les valeurs de plusieurs éléments (groupes, rôles, licences) sont séparées par « | ».

Tous les CSV produits par l'outil (listes, plan d'action, journal) neutralisent les **formules** : une valeur qui commence par `=`, `+`, `-` ou `@` est précédée d'une apostrophe, pour qu'un nom de compte ou de groupe piégé ne s'exécute pas à l'ouverture dans Excel.

### Journal de la comparaison avant / après

Colonnes : `Categorie` (Constat, Compte, Appartenance, Groupe, Imbrication), `Changement` (Traité, Nouveau, Statut modifié, Supprimé, Créé, Retirée, Ajoutée), `Element`, `Libelle`, `Avant`, `Apres`, `Detail`. Les éléments sont rapprochés par leur clé (UPN pour les utilisateurs, nom pour les groupes) : un compte renommé apparaît comme supprimé puis créé.

## 5. Pièges connus

- **Partages** : le journal d'audit ne contient que les partages créés pendant la période (180 jours au plus en Audit Standard) ; il ne dit pas si un lien est encore actif. Les deux exports d'un avant / après couvrent des périodes qui se recouvrent.

- **Imbrication Entra** : les groupes M365 ne peuvent pas contenir d'autres groupes, et un groupe dynamique ne peut pas avoir de membres imbriqués.
- **Héritage Entra** : l'accès conditionnel suit l'imbrication, mais l'attribution de licences par groupe et l'assignation aux applications d'entreprise ne la suivent pas. Un utilisateur d'un sous-groupe peut figurer sur la carte sans recevoir la licence ou l'accès.
- **Groupes synchronisés** : leur composition et la désactivation de leurs membres se gèrent dans l'AD, pas dans Entra.
- **Homonymes** : deux groupes de même nom seraient fusionnés sur la carte ; les scripts les suffixent (début de l'Id côté Entra, OU côté AD).
- **AD multi-domaines** : seuls les objets du domaine interrogé sont reliés ; les autres sont comptés dans `NbAutres`.
- **Très gros groupes AD** : `Get-ADGroupMember` échoue au-delà de 5000 membres (limite ADWS). Les scripts lisent l'attribut `member` et signalent les groupes concernés.
- **Propriétaires** : Graph ne les expose pas pour les groupes synchronisés ni pour les listes de distribution ; la colonne `NbProprietaires` reste alors vide et ces groupes ne sont pas contrôlés.
- **Rôles éligibles** : sans licence Entra ID P2, seules les attributions actives sont lisibles ; elles apparaissent en « Active ».
- **`lastLogonTimestamp`** : répliqué avec 9 à 14 jours de retard, donc impropre à un suivi au jour près.

## 6. Intégrité et signature

`SHA256SUMS.txt` contient l'empreinte de chaque fichier du paquet original. Vérification, depuis le dossier `MapTools` :

```powershell
Get-Content .\SHA256SUMS.txt | ForEach-Object {
    $h, $f = $_ -split '  ', 2
    $ok = (Get-FileHash -Algorithm SHA256 $f).Hash -eq $h
    '{0,-5} {1}' -f $(if ($ok) {'OK'} else {'DIFF'}), $f
}
```

Pour garantir l'origine des scripts, les signer avec un certificat de signature de code :

```powershell
$cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert | Select-Object -First 1
Get-ChildItem .\scripts\*.ps1 | Set-AuthenticodeSignature -Certificate $cert -TimestampServer 'http://timestamp.digicert.com'
```

Toute modification ultérieure invalide la signature, et une stratégie d'exécution `AllSigned` refuse alors le script. La signature modifie les fichiers : régénérer `SHA256SUMS.txt` ensuite.

## 6 bis. Profil client (JSON)

Fichier `MapTools_Profil_<client>.json`, produit par **Profil client > Exporter** :

| Clé | Contenu |
|---|---|
| `mapTools`, `v` | Marqueur `profil` et version du format (1) |
| `client`, `author`, `confid` | Nom du client, intervenant, mention de confidentialité |
| `logoProvider`, `logoClient` | Logos en image PNG intégrée (data URI, 320 px de large au plus) |
| `sensitive` | Mots ajoutés à la détection des groupes sensibles, séparés par des virgules |
| `exceptions` | Risques acceptés : `fid` (identifiant du constat, ou `*` pour tous), `key` (élément), `label`, `finding`, `just`, `by`, `date`, `review` |
| `plan` | Plan d'action par identifiant de constat : `owner`, `due`, `status` |

Dans le navigateur, le profil est mémorisé sous la clé `profil:<préfixe des fichiers>` de la base `maptools`. Le plan d'action CSV ajoute les colonnes `Responsable`, `Echeance` et `StatutPlan`.

## 7. Historique

| Version | Date | Changements |
|---|---|---|
| 1.1 | 01.10.2026 | Partages SharePoint et OneDrive et paramètres des sites (`Export-MapToolsPartages.ps1`, `-AvecPartages`) ; rôles Entra et PIM, propriétaires de groupes, MFA, licences et économies, secours sans licence P1 ; contrôles de mots de passe AD ; rapport décisionnel avec page de garde, synthèse pour la direction, indice d'hygiène, plan d'action pilotable, risques acceptés ; avant / après avec suivi du plan ; profil client (JSON) ; « Pourquoi a-t-il accès ? » ; exports de listes ; nouvelle interface (barre du haut, menu Outils, volets repliables, panneau de détails en onglets, bandeau des filtres, infobulles), 15 vues rapides, niveaux et espacement ; dossier `dev` (sources et tests) |
| 1.0 | 24.09.2026 | Première version publiée |

## 8. Composants tiers

- vis-network 9.1.9 : double licence Apache-2.0 / MIT (`lib\LICENSE-vis-network-*.txt`)
- PapaParse 5.4.1 : MIT (`lib\LICENSE-papaparse.txt`)
