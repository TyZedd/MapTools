# MapTools · Guide d'utilisation

Version 1.1 · © 2026 Jimmy Paul

Ce guide suit le déroulé d'une utilisation : préparer le poste, exporter, charger, lire la carte, naviguer. Pour adapter l'outil ou utiliser d'autres sources, voir `REFERENCE.md`.

**Sommaire** : 1. Préparer le poste · 2. Exporter les données · 3. Charger les fichiers · 4. Lire la carte · 5. Naviguer · 6. Rapport décisionnel · 7. Comparer avant / après · 8. Profil client · 9. Mémorisation · 10. Dépannage

## 1. Préparer le poste

Les scripts téléchargés portent la marque « provenant d'Internet » et sont bloqués par la stratégie d'exécution. Depuis le dossier `MapTools` :

```powershell
Get-ChildItem .\scripts\*.ps1 | Unblock-File
```

Pour Entra ID, installer une fois les modules Microsoft Graph :

```powershell
Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Groups, Microsoft.Graph.Reports, Microsoft.Graph.Identity.DirectoryManagement, Microsoft.Graph.Identity.Governance -Scope CurrentUser
```

PowerShell 5.1 et 7 ont des dossiers de modules séparés : installer les modules dans la version utilisée ensuite. Tous les modules `Microsoft.Graph.*` doivent avoir la même version.

Pour Active Directory, il faut le module ActiveDirectory (RSAT ou contrôleur de domaine). Un compte en simple lecture suffit.

## 2. Exporter les données

Les scripts ne font que des lectures. Ils produisent deux CSV horodatés (liens et attributs), séparateur `;`, UTF-8 avec BOM, et affichent un résumé des statuts avec leurs pourcentages.

### Entra ID

```powershell
Connect-MgGraph -Scopes "GroupMember.Read.All","User.Read.All","AuditLog.Read.All","LicenseAssignment.Read.All","RoleManagement.Read.Directory","Reports.Read.All" -TenantId "<domaine du client>" -NoWelcome
.\scripts\Export-MapToolsEntra.ps1 -Prefixe <Client> -DossierExport C:\Temp
```

En accès délégué (GDAP), `-TenantId` est indispensable : sans lui, la connexion se fait sur le tenant du compte technicien. Le script réutilise la session ouverte et fait un appel par groupe : compter quelques minutes pour plusieurs centaines de groupes.

Pour obtenir les dates de connexion, trois conditions sont nécessaires :

- le scope `AuditLog.Read.All` ;
- un rôle Entra **Reports Reader**, **Security Reader** ou **Global Reader** sur le compte connecté (en GDAP, ce rôle doit faire partie de la relation) ;
- une licence **Entra ID P1 ou P2** dans le tenant.

S'il en manque une, le script le signale et bascule sur un **secours** : le rapport d'activité Microsoft 365 (scope `Reports.Read.All`, sans licence P1). Il donne la dernière activité par service (Exchange, OneDrive, SharePoint, Teams) sur 180 jours, pour les comptes licenciés. Il est moins précis qu'une date de connexion : un compte sans licence reste en « Activité inconnue », un compte sans aucune activité sur 180 jours est classé inactif (si le seuil ne dépasse pas 6 mois), et la colonne `SourceDate` indique l'origine de la date. Le rapport décisionnel le précise.

Piège : Microsoft **anonymise** par défaut les noms dans ces rapports. Le script le détecte et le signale ; pour l'éviter, décocher l'option d'affichage de noms masqués dans le centre d'administration Microsoft 365, sous Paramètres > Paramètres de l'organisation > Rapports.

Sans P1 et sans ce secours, tous les comptes activés sont en « Activité inconnue » : en environnement hybride, l'export AD donne l'activité des comptes synchronisés sans condition de licence.

Le script lit aussi, sauf si on le lui interdit :

- le **statut MFA** de chaque compte (méthode enregistrée ou non, méthodes, rôle d'administration), avec les mêmes conditions que les dates de connexion ;
- les **licences** attribuées, avec le scope `LicenseAssignment.Read.All` et un rôle Directory Readers ou Global Reader. Les licences courantes sont affichées sous leur nom commercial, les autres sous leur référence technique.

- les **rôles d'administration** Entra (Global Administrator, Exchange Administrator…), attribués à des comptes ou à des groupes, avec le type d'attribution : permanente, temporaire, activée ou éligible (PIM). Scope `RoleManagement.Read.Directory` ; les attributions éligibles demandent Entra ID P2 ;
- les **propriétaires** des groupes cloud de sécurité et Microsoft 365. Graph ne les expose pas pour les groupes synchronisés ni pour les listes de distribution. Cette lecture fait un appel de plus par groupe.

`-SansMFA`, `-SansLicences`, `-SansRoles` et `-SansProprietaires` désactivent ces lectures. Si un module ou un droit manque, le script le signale et exporte le reste.

### Active Directory

```powershell
.\scripts\Export-MapToolsAD.ps1 -Prefixe <Client> -DossierExport C:\Temp
```

Ajouter `-Server` pour cibler un contrôleur de domaine, `-SearchBase` pour limiter les groupes exportés à une OU. La date de référence est `LastLogonDate`, qui peut avoir 9 à 14 jours de retard : sans importance pour un seuil exprimé en mois.

Un export AD ne contient ni MFA, ni licences, ni rôles Entra, ni propriétaires : l'outil reconnaît la source, grise les vues concernées, et le rapport liste ces contrôles comme « non réalisés ». Il ajoute à la place les contrôles de mots de passe (expiration désactivée, ancienneté). Pour un environnement hybride, faire les deux exports et produire un rapport pour chacun.

### Partages SharePoint et OneDrive (facultatif)

```powershell
Install-Module ExchangeOnlineManagement, Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser   # une fois
.\scripts\Export-MapToolsEntra.ps1 -Prefixe <Client> -DossierExport C:\Temp -AvecPartages
```

`-AvecPartages` enchaîne `Export-MapToolsPartages.ps1` à la fin de l'export Entra, avec le même préfixe et le même horodatage. Le script peut aussi être lancé seul. Il produit deux fichiers de plus :

- **Partages** : les partages créés pendant la période d'audit (180 jours au plus avec Audit Standard), lus dans le journal d'audit Purview : liens « Tout le monde », liens organisation, liens personnes spécifiques, invitations, accès accordés. Pour chacun : date, auteur, destinataire, externe ou non, élément et site. Un lien « Tout le monde » supprimé après sa création est signalé ;
- **Sites** : le niveau de partage externe de chaque site (désactivé, invités existants, nouveaux invités, tout le monde) et le paramètre du tenant, avec l'expiration imposée aux liens anonymes.

Prérequis et pièges :

- **Audit** : rôle View-Only Audit Logs, Audit Logs ou Global Reader ; en accès délégué, le script se connecte à Exchange Online avec `-DelegatedOrganization` (domaine du tenant déduit de la session Graph, ou `-Organisation`) ;
- **Sites** : rôle **SharePoint Administrator**, et l'URL d'administration (`https://<tenant>-admin.sharepoint.com`), déduite du domaine initial du tenant ou passée par `-SharePointAdminUrl` ;
- **PowerShell 7** : le module SharePoint Online n'y fonctionne qu'importé avec `-UseWindowsPowerShell` (Windows uniquement), ce que fait le script ;
- **Ce que l'audit ne dit pas** : il montre les partages **créés** pendant la période, pas l'état actuel des liens. Un lien créé avant la période n'apparaît pas. Pour l'inventaire exhaustif d'un site précis, un parcours ciblé (PnP PowerShell) reste nécessaire ;
- si un module ou un rôle manque, la partie concernée est signalée et ignorée : l'export principal n'est jamais interrompu.

### Audit seul des comptes (Entra ID, sans les groupes)

```powershell
.\scripts\Get-EntraIDComptesInactifs.ps1 -TenantId "<domaine du client>" -MoisInactivite 6 -InclureInvites
```

Produit `EntraID_<date>_Actifs_Inactifs.csv` (comptes activés inactifs ou jamais connectés) et `EntraID_<date>_Desactives.csv`. Utile pour un rapport Excel rapide.

Tous les paramètres des scripts sont détaillés dans `REFERENCE.md`.

## 3. Charger les fichiers

Ouvrir `MapTools.html` et déposer **tous les CSV d'un export en même temps** (liens et attributs, plus partages et sites s'ils existent) dans la zone prévue (ou cliquer dessus pour les choisir). Les fichiers produits par les scripts sont reconnus automatiquement et la carte s'affiche aussitôt.

Chaque fichier chargé apparaît dans une fiche indiquant son rôle (liens, attributs ou ignoré) et les colonnes utilisées. Pour un CSV d'une autre source, corriger ces choix puis cliquer sur **Afficher la carte** (voir `REFERENCE.md`, « Formats CSV »).

**Voir un exemple** charge un jeu fictif ; si des fichiers sont déjà chargés, une confirmation est demandée car l'exemple les remplace. **Tout effacer** vide l'outil et la mémoire du navigateur, après confirmation. Dans ces boîtes de confirmation, Entrée et Échap annulent : seul un clic sur le bouton d'action valide.

## 4. Lire la carte

### Éléments

- **Groupes** : rectangles colorés par type (Sécurité, M365, Distribution…). Bordure en pointillés : groupe synchronisé depuis l'AD (export Entra) ou groupe protégé par AdminSDHolder (export AD). Forme ovale : groupe dynamique.
- **Rôles d'administration** (export Entra) : rectangles foncés aux coins arrondis, à gauche. Trait plein vers un compte ou un groupe : attribution active ou permanente ; pointillés : attribution éligible (PIM). Le panneau de détails d'un compte liste ses rôles, directs ou hérités d'un groupe.
- **Utilisateurs** : points colorés par statut, nom affiché dessous. Un **losange** signale un compte activé sans méthode MFA enregistrée (export Entra).
- **Disposition** : racines à gauche, un niveau d'imbrication par colonne. Un élément présent à plusieurs niveaux est placé au plus profond.
- **Boucle** : flèche orange en pointillés, quand un groupe se contient lui-même par imbrication. Le panneau de détails indique le chemin de la boucle.

### Statuts des utilisateurs

Le seuil N vaut 6 mois par défaut (paramètre `-MoisInactivite`).

| Statut | Couleur | Signification |
|---|---|---|
| Actif | vert | Dernière connexion réussie il y a moins de N mois |
| Nouveau (jamais connecté) | turquoise | Jamais connecté, compte créé il y a moins de N mois |
| Jamais connecté | jaune | Jamais connecté, compte créé il y a plus de N mois |
| Inactif > N mois | orange | Dernière connexion plus ancienne que N mois |
| Désactivé | rouge | Compte désactivé, quelle que soit son activité |
| Activité inconnue | gris | Dates de connexion non lues à l'export |

Côté Entra, la référence est la **dernière connexion réussie** : une tentative échouée, par exemple une pulvérisation de mots de passe, ne rend pas un compte actif. Cette information n'existe que depuis décembre 2023 ; pour les comptes plus anciens, la dernière tentative est utilisée et la colonne `SourceDate` le précise.

### Profil

Chaque utilisateur reçoit un profil probable : Invité, Service / sync, Admin, Salle / équipement, Technique / test ou Standard. C'est une **déduction à partir du nom ou de l'UPN**, pas une information de l'annuaire.

Les boîtes partagées et les salles sont souvent des comptes activés jamais connectés, par défaut dans Exchange Online : elles gonflent le nombre de « Jamais connecté ». Pour les confirmer :

```powershell
Connect-ExchangeOnline
Get-EXOMailbox -RecipientTypeDetails SharedMailbox,RoomMailbox,EquipmentMailbox -ResultSize Unlimited |
    Select-Object DisplayName, UserPrincipalName, RecipientTypeDetails
```

### Statistiques

La barre du bas indique le nombre de groupes reliés, de liens d'imbrication, d'utilisateurs membres et d'appartenances, de racines et de niveaux. S'affichent aussi, en orange, les boucles détectées et les éléments absents du fichier d'attributs (souvent une différence de nom entre les deux fichiers).

## 5. Naviguer

### Organisation de l'écran

- **Barre du haut** : recherche, vue Arborescence / Réseau, liste des vues rapides, le bouton **Rapport décisionnel**, le menu **Outils ▾** (Avant / après, Exporter des listes, Profil client, Exporter la vue en PNG) et **?** (aide). Les actions qui demandent des données restent grisées tant que rien n'est chargé.
- **Panneau de gauche**, en trois volets repliables d'un clic sur leur titre : **Données** (dépôt des fichiers, fiches de correspondance, exemple, effacement, mémorisation ; replié automatiquement une fois la carte affichée, avec le nombre de fichiers et la source), **Affichage** (niveaux, espacement, utilisateurs, interrupteurs, recentrage) et **Légende**.
- **Carte** au centre ; la description de la vue rapide active s'affiche en haut de la carte, avec les boutons ◀ ▶ pour la vue « Périmètre d'un groupe à privilèges ».
- **Bandeau des filtres actifs**, en haut à gauche de la carte : vue rapide, chemin, niveaux, mode utilisateurs, catégories masquées, branche isolée… Le bouton **Tout réinitialiser** revient à l'affichage par défaut. Si un élément attendu n'apparaît pas, c'est là qu'on comprend pourquoi.
- **Infobulles** : le survol d'un élément affiche l'essentiel (statut, profil, dernière connexion, MFA, nombre de groupes ou de membres, propriétaires) sans cliquer.
- **Panneau de détails** à droite, ouvert à la sélection d'un élément et fermé par le bouton × (la carte et la sélection ne changent pas).
- **Barre du bas** : les chiffres clés (groupes, liens, utilisateurs, appartenances, rôles, racines, niveaux, boucles) et la source reconnue. Certaines pastilles sont cliquables : groupes (tout réinitialiser), utilisateurs (les afficher tous), rôles (vue des rôles), boucles (les isoler), absents des attributs (les afficher).

### Sélection et détails

Un **clic** sur un élément ouvre ses détails à droite, en onglets (le dernier onglet consulté reste ouvert d'un élément à l'autre) :

- **Résumé** : l'essentiel en clair. Pour un utilisateur : statut, profil, type de compte, dernière connexion, MFA, licences, nombre de groupes, rôles, partages. Pour un groupe : catégorie, membres directs et tous niveaux, propriétaires, synchronisation, rôles, et la répartition des utilisateurs par statut ;
- **Accès** : « Pourquoi a-t-il accès ? », rôles hérités, groupes parents, sous-groupes et utilisateurs directs, avec le bouton d'export correspondant ;
- **Partages** (si les partages sont chargés) : partages reçus et effectués ;
- **Attributs** : toutes les colonnes de l'export avec des libellés lisibles ; les identifiants techniques sont repliés sous « Détails techniques ».

Chaque nom listé est cliquable. La **recherche** accepte un nom de groupe, un nom complet ou un UPN, même partiel ; valider avec Entrée.

### Vues rapides

La liste **Vue rapide** applique en un clic un ensemble de réglages adaptés à une question précise. Tout réglage modifié ensuite à la main la repasse en « Personnalisée ».

| Vue rapide | Question traitée | Réglages appliqués |
|---|---|---|
| **Structure des groupes** | Comment les groupes s'imbriquent-ils ? Y a-t-il des boucles, des niveaux trop profonds ? | Utilisateurs et groupes sans imbrication masqués, toute la légende visible |
| **Comptes à traiter** | Où se trouvent les comptes dormants dans la structure ? | Tous les utilisateurs, statuts Actif et Nouveau masqués |
| **Désactivés encore membres** | Quels comptes désactivés gardent des appartenances ? | Tous les utilisateurs, seul le statut Désactivé visible |
| **Périmètre d'un groupe à privilèges** | Qui hérite des droits d'un groupe sensible, directement ou par imbrication, et avec quel statut ? | Premier groupe sensible sélectionné, branche isolée avec tous ses utilisateurs ; les boutons ◀ ▶ passent au groupe sensible suivant |
| **Rôles d'administration** | Qui détient un rôle d'administration Entra, directement ou par un groupe, et avec quelle attribution ? | Rôles et tout ce qui en dépend, avec tous les utilisateurs concernés |
| **Comptes activés sans MFA** | Quels comptes activés n'ont aucune méthode MFA, et dans quels groupes ? | Seuls les utilisateurs activés sans MFA sont affichés, avec les groupes |
| **Invités et leurs groupes** | Quels invités ont accès à quoi ? | Comptes invités et les groupes qui les contiennent directement |
| **Invités destinataires de partages** | Quels invités ont reçu des documents ? | Invités ayant au moins un partage reçu, avec leurs groupes (partages chargés) |
| **Licences sur comptes inactifs ou désactivés** | Où se trouvent les licences à récupérer ? | Comptes inactifs, jamais connectés ou désactivés encore licenciés, avec leurs groupes |
| **Comptes de service** | Où sont utilisés les comptes de service ? | Comptes de profil « Service » et leurs groupes |
| **Comptes d'administration** | Qui a des droits élevés ? | Profil admin, rôle Entra, compte protégé ou membre d'un groupe sensible, avec leurs groupes |
| **Groupes Microsoft 365 sans propriétaire** | Quelles équipes ne sont gérées par personne ? | Groupes M365 dont le nombre de propriétaires est nul |
| **Boucles d'imbrication** | Quels groupes se contiennent eux-mêmes ? | Uniquement les groupes pris dans une boucle |
| **Groupes isolés et vides** | Quels groupes n'ont aucun lien ? | Utilisateurs masqués, éléments sans lien affichés : en arborescence, les groupes isolés apparaissent en colonne à gauche, sans flèche |

Les vues « Comptes à traiter » et « Désactivés encore membres » sont aussi parlantes en vue Réseau : les groupes qui concentrent des comptes à traiter ressortent. Les groupes sensibles sont détectés comme dans le rapport (section 6). Les vues qui demandent une donnée absente de l'export (rôles, MFA, licences, profils, propriétaires, partages, boucles) sont grisées. Une vue rapide repart toujours de « Niveaux : tous » et efface un éventuel chemin affiché. Elle garde la disposition en cours : passer d'Arborescence à Réseau (ou l'inverse) ne change ni la vue rapide ni les filtres. Les réglages faits ensuite (niveaux, espacement, légende, utilisateurs, interrupteurs, sélection d'un élément) s'ajoutent à la vue rapide sans l'annuler : le bandeau des filtres les affiche ensemble, et seuls le choix d'une autre vue, « Personnalisée » ou **Tout réinitialiser** la retirent. La vue rapide et ses ajustements sont mémorisés avec les données ; un nouveau chargement de fichiers repart d'un affichage par défaut.

### Pourquoi a-t-il accès ?

En haut du panneau de détails d'un utilisateur ou d'un groupe, le bloc **Pourquoi a-t-il accès ?** liste tous les rôles et groupes dont il hérite, classés en rôles, groupes sensibles et autres groupes. Choisir une cible puis **Afficher** :

- la carte n'affiche plus que les **chemins d'héritage** entre l'élément et la cible, surlignés, avec le type d'attribution sur les liens de rôle (pointillés si éligible) ;
- le panneau liste chaque chemin, du plus court au plus long, par exemple « Thomas Genoud → GG_IT_Helpdesk → GG_IT → GG_IT_Admins → SG_PIM_Exchange_Admin → Exchange Administrator (Éligible) » ;
- **Effacer le chemin** revient à la carte normale.

Plusieurs chemins vers la même cible signalent souvent une imbrication redondante, voire une boucle.

### Exporter des listes

**Outils > Exporter des listes** (barre du haut) ouvre un écran pour extraire n'importe quelle liste des données chargées, en CSV ou par copier-coller dans Excel.

À gauche, des **listes prêtes** : utilisateurs avec MFA, comptes activés sans MFA, membres d'un groupe, groupes d'un utilisateur, détenteurs d'un rôle, comptes inactifs ou jamais connectés, comptes désactivés, invités, comptes avec licence, groupes sans propriétaire, tous les utilisateurs, groupes ou rôles. Celles qui demandent une donnée absente de l'export sont grisées.

À droite, les **filtres**, qui se combinent :

- **Utilisateurs** : statut (plusieurs possibles), MFA, membres ou invités, profil, licence, membre d'un groupe (membres directs seulement, ou tous niveaux avec l'imbrication), détenteur d'un rôle, recherche par nom ou identifiant ;
- **Groupes** : catégorie, groupes d'un utilisateur (appartenance directe ou héritée), sans propriétaire, sans membre, recherche ;
- **Rôles** : tous, avec le nombre de détenteurs et les attributions directes.
- **Partages** (si chargés) : type de partage, destinataire (externes, liens « tout le monde »), personne (a partagé ou a reçu), site, recherche. La colonne « Statut du destinataire » indique si l'invité est actif, inactif ou désactivé dans l'annuaire ;
- **Sites** (si chargés) : niveau de partage externe.

Listes prêtes supplémentaires avec les partages : liens « Tout le monde » créés, partages vers l'extérieur, partages d'une personne, sites ouverts aux liens « Tout le monde », tous les sites. Le panneau de détails d'un utilisateur ou d'un invité affiche ses **partages reçus et effectués**, avec un bouton **Exporter ses partages**.

Le menu **Colonnes** choisit les colonnes exportées : toutes celles du fichier d'attributs, plus des colonnes calculées (groupes directs, groupes tous niveaux, rôles d'administration, nombre d'utilisateurs tous niveaux, sous-groupes, parents). Quand un filtre porte sur un groupe ou un rôle, la colonne **Appartenance** indique si le lien est direct ou « via » quel sous-groupe, et **Attribution du rôle** donne le type (permanente, éligible…).

L'aperçu montre les 100 premières lignes ; **Exporter en CSV** produit la liste complète, **Copier pour Excel** la place dans le presse-papiers, prête à coller.

Raccourci : dans le panneau de détails, **Exporter les membres** (groupe), **Exporter ses groupes** (utilisateur) ou **Exporter les détenteurs** (rôle) ouvre directement la bonne liste.

### Commandes de la barre du haut et du volet Affichage

| Commande | Effet |
|---|---|
| **Vue rapide** | Applique un ensemble de réglages prédéfini (voir ci-dessus) |
| Vue **Arborescence / Réseau** | Disposition hiérarchique ou libre |
| **Niveaux** | Limite la profondeur d'imbrication affichée (1, 2, 3… niveaux de groupes). Un groupe dont les sous-groupes sont masqués porte « +n » |
| **Espacement** Compact / Normal / Aéré | Resserre ou aère la carte ; Compact réduit aussi la taille du texte et la longueur des libellés |
| Interrupteur **Groupes sans imbrication** | Affiche ou masque les groupes qui n'ont ni parent ni sous-groupe (ils ne contiennent que des utilisateurs). Au chargement, ils sont masqués d'office s'ils sont plus de 100, sinon l'arborescence devient une colonne interminable ; un message l'indique |
| Utilisateurs **Sélection / Tous / Masqués** | Membres du groupe sélectionné (ou de toute la branche isolée), tous les utilisateurs, ou aucun |
| Interrupteur **Animation de la carte** | Vue Réseau uniquement : la disposition est figée par défaut une fois calculée ; activer pour remettre les éléments en mouvement |
| Interrupteur **Isoler la branche sélectionnée** | N'affiche que l'élément, ses parents et ses descendants (aussi par double-clic) |
| Interrupteur **Éléments sans lien** | Ajoute les groupes vides et les utilisateurs membres d'aucun groupe |
| Interrupteur **Mémoriser sur ce poste** (volet Données) | Voir section 9 |
| **Recentrer** | Affiche toute la carte |
| **Outils > Exporter la vue en PNG** | Enregistre la vue affichée, avec la mention MapTools, l'auteur et la date |
| **Rapport décisionnel** | Voir section 6 |
| **Outils > Exporter des listes** | Voir ci-dessus |
| **Outils > Avant / après** | Voir section 7 |
| **Outils > Profil client** | Voir section 8 |

La **légende** est cliquable : chaque catégorie se masque ou se réaffiche d'un clic.

### Zoom

Molette (centré sur le pointeur), boutons − et + en bas à gauche, touches `+` et `-`. Le bouton ⤢ ou la touche `0` affiche tout. Le zoom est limité entre 8 % et 250 %.

Au-delà de 800 éléments affichés, un avertissement apparaît : masquer les utilisateurs ou isoler une branche pour garder une carte lisible.

## 6. Rapport décisionnel

Le bouton **Rapport décisionnel** analyse les données chargées et produit un rapport prêt à remettre. Il s'ouvre sur une **synthèse pour la direction**, conçue pour tenir sur une page imprimée et se lire sans connaissance technique :

- un **indice d'hygiène des identités** de 0 à 100 avec une note de A à E, et le détail de ses composantes ;
- quatre **indicateurs clés** colorés (vert, jaune, rouge) : part des comptes activés utilisés, part des comptes avec MFA, comptes à privilèges à risque, licences récupérables avec l'économie estimée ;
- un **graphique** de répartition des comptes par statut, et les **points à traiter par priorité**, en distinguant comptes utilisateurs et groupes ;
- les **trois risques prioritaires**, formulés en langage courant avec l'action à mener.

L'indice est une moyenne pondérée des composantes disponibles : activité des comptes (25 %), MFA (25 %), part des comptes à privilèges sans constat (30 %), hygiène des groupes (20 %, moins 10 points par boucle). A ≥ 85, B ≥ 70, C ≥ 55, D ≥ 40, E en dessous. Il est **indicatif** : il sert à suivre l'évolution d'un même client, pas à comparer des organisations entre elles.

La suite du rapport est séparée en parties : **1. Comptes utilisateurs** (indicateurs, répartition, licences, constats), **2. Groupes** (indicateurs de structure, constats) et, si les fichiers de partages ou de sites sont chargés, **3. Partages et sites**. Chaque partie a ses propres contrôles sans anomalie.

- **Synthèse** rédigée : nombre de comptes, part des comptes activés réellement utilisés, nombre de constats par priorité, et les trois actions prioritaires ;
- **Indicateurs** et répartition des comptes par statut ;
- **Constats**, classés par priorité (Critique, Élevée, Moyenne, Faible). Chacun indique le risque, l'action recommandée et la liste des comptes ou groupes concernés (25 premiers affichés) ;
- **Contrôles sans anomalie**, puis le **périmètre et les limites** de l'analyse.

| Contrôle | Priorité |
|---|---|
| Accès privilégiés portés par des comptes non actifs (inactifs, jamais connectés ou désactivés, membres directs ou par imbrication d'un groupe sensible) | Critique |
| Comptes d'administration dormants | Critique |
| Rôles d'administration détenus par des comptes non actifs (directement ou par un groupe) | Critique |
| Rôles privilégiés attribués en permanence (au lieu d'éligible via PIM) | Élevée |
| Trop d'administrateurs généraux (plus de 4) · Un seul administrateur général | Élevée · Moyenne |
| Comptes à privilèges : mot de passe ancien ou sans expiration (export AD) | Élevée |
| Comptes à privilèges sans MFA enregistrée (profil admin, rôle d'administration Entra ou groupe sensible) | Critique |
| Comptes activés sans MFA enregistrée | Élevée |
| Licences attribuées à des comptes inactifs ou désactivés | Moyenne |
| Liens « Tout le monde » créés pendant la période d'audit (hors liens supprimés ensuite) | Élevée |
| Partages vers des comptes externes inactifs ou désactivés | Élevée |
| Éléments partagés avec des personnes externes | Moyenne |
| Sites autorisant les liens « Tout le monde » | Moyenne |
| Liens « Tout le monde » sans date d'expiration (paramètre du tenant) | Moyenne |
| Groupes Microsoft 365 sans propriétaire | Moyenne |
| Groupes dont tous les propriétaires sont inactifs ou désactivés | Moyenne |
| Comptes activés dont le mot de passe n'expire jamais (export AD) | Moyenne |
| Mots de passe inchangés depuis plus de 365 jours (export AD) | Faible |
| Comptes désactivés encore membres de groupes | Élevée |
| Comptes utilisateurs inactifs (hors admin, service, salles et invités) | Élevée |
| Invités inactifs ou jamais connectés | Moyenne |
| Comptes de service inactifs | Moyenne |
| Comptes activés jamais connectés (souvent des boîtes partagées ou des salles) | Moyenne |
| Boucles d'imbrication | Moyenne |
| Groupes sans aucun membre actif | Faible |
| Imbrication profonde (5 niveaux ou plus) | Faible |
| Groupes vides | Faible |

L'en-tête du rapport indique la source reconnue (Entra ID, Active Directory ou CSV). La section « Périmètre et limites » liste les **contrôles non réalisés** avec ces données et pourquoi : par exemple la MFA avec un export AD, ou les mots de passe avec un export Entra.

Quand l'export contient les licences, une section **Licences récupérables** détaille, par licence, le nombre de comptes inactifs ou désactivés concernés. Saisir le prix mensuel de chaque licence (tarif du contrat du client, dans la devise indiquée) : les totaux mensuel et annuel se calculent et la synthèse cite l'économie potentielle. Les prix sont mémorisés sur le poste pour les rapports suivants ; ce ne sont pas des données client. Avant de retirer la licence d'une boîte aux lettres à conserver, la convertir en boîte partagée : sans licence Exchange, elle est supprimée au bout de 30 jours.

Le rapport s'ouvre sur une **page de garde** : logos du prestataire et du client, titre, nom du client, source, date des données et du rapport, intervenant, sommaire et mention de confidentialité (tous issus du profil client, section 8). À l'impression, elle occupe la première page, et les pages suivantes portent en bas le client, la mention de confidentialité et le numéro de page (Chrome et Edge récents).

Après la synthèse pour la direction vient le **Plan d'action** : pour chaque constat, un **responsable** (client, prestataire…), une **échéance** et un **statut** (à faire, en cours, fait, reporté), saisis directement dans le tableau. Ils sont gardés dans le profil client, imprimés tels quels et repris dans le plan d'action CSV et dans l'avant / après.

Chaque élément d'un constat a un bouton **Accepter** (non imprimé) pour enregistrer un **risque accepté** : justification obligatoire, validateur côté client, date de revue (un an par défaut), pour ce constat ou pour tous. L'élément sort des constats et de l'indice, et apparaît dans la section **Risques acceptés** avec sa justification ; à sa date de revue, il redevient un constat et la section le signale en rouge. **Retirer** annule l'acceptation. Usage typique : comptes d'urgence (break glass), comptes de service justifiés.

Dans le rapport :

- le **nom du client** vient du profil client (ou du préfixe des fichiers) et reste modifiable avant impression ;
- **Imprimer / PDF** ouvre l'impression du navigateur ; choisir « Enregistrer au format PDF ». La mise en page est prévue pour du A4 ;
- **Exporter le plan d'action (CSV)** produit la liste complète des éléments concernés, une ligne par élément et par constat, avec la priorité et l'action recommandée. Elle est prête à importer dans un outil de tickets ou à filtrer dans Excel ;
- **Fermer** ou Échap revient à la carte.

Les groupes « sensibles » et les profils (admin, service, salle…) sont déduits des noms. Relire les constats critiques avant de les présenter comme établis. Le rapport ne couvre ni les rôles d'annuaire et PIM, ni l'authentification et l'accès conditionnel, ni les applications, ni la configuration de l'AD : il le précise lui-même en fin de document.

## 7. Comparer avant / après

Pour montrer l'effet d'une remédiation, MapTools compare deux exports du même annuaire pris à deux dates. Faire un export avant les travaux, un autre après, puis **Outils > Avant / après** dans la barre du haut.

1. Déposer les CSV du premier export dans **Avant** et ceux du second dans **Après** : liens et attributs de chaque côté. **Utiliser les fichiers chargés** reprend les fichiers déjà ouverts sur la carte, pour l'un ou l'autre côté.
2. Vérifier les dates : elles sont lues dans le nom des fichiers produits par les scripts (`_AAAAMMJJ_HHmm_`) et restent modifiables. Le bouton ⇄ inverse les deux côtés.
3. Cliquer sur **Comparer**.

Le rapport produit contient :

- une **synthèse pour la direction** : indice avant → après avec les deux notes, taux de traitement en barre de progression, quatre indicateurs avant → après colorés selon le sens de l'évolution, et les points à traiter par priorité avant / après ;
- une **synthèse détaillée** : éléments traités sur le total relevé, éléments restants et nouveaux, évolution des comptes dormants et de la part des comptes utilisés, appartenances retirées ;
- deux parties, **comptes utilisateurs** et **groupes**, chacune avec son tableau d'**indicateurs** avant / après, avec l'évolution en vert quand elle va dans le bon sens et en rouge sinon (en gris pour les chiffres neutres, comme le nombre de comptes) ;
- le tableau des **constats** du rapport décisionnel, avec pour chacun le nombre avant, après, traités et nouveaux, puis le détail des éléments traités et nouveaux ;
- une partie **partages et sites** quand les deux exports en contiennent : liens anonymes, partages externes, destinataires externes, sites ouverts, et les constats correspondants ;
- les **attributions de rôles** retirées, ajoutées ou modifiées (par exemple « Permanente → Éligible » après un passage à PIM) ;
- les **changements dans l'annuaire** : changements de statut (avec le décompte par transition, par exemple « Jamais connecté → Désactivé »), appartenances retirées et ajoutées, comptes supprimés et créés, groupes et imbrications supprimés ou ajoutés.

Si un plan d'action a été saisi dans le profil client, une section **Suivi du plan d'action** compare, pour chaque constat planifié, le responsable, l'échéance et le statut déclaré à l'état réel du second export : traité, partiel, en attente, ou échéance dépassée.

**Imprimer / PDF** produit le livrable client, avec la même page de garde que le rapport ; **Exporter le journal (CSV)** la liste complète des changements, une ligne par changement.

À savoir pour la lecture :

- la comparaison porte sur deux photographies : un changement fait puis annulé entre les deux dates n'apparaît pas ;
- les statuts dépendent de la date d'export : un compte actif peut devenir « inactif » par simple écoulement du temps et apparaître comme nouveau constat ;
- un élément « traité » a disparu du constat, qu'il ait été corrigé ou supprimé ;
- utiliser le même seuil d'inactivité (`-MoisInactivite`) pour les deux exports ; l'outil prévient si les seuils diffèrent ou si les dates semblent inversées.

Le dossier `exemplevant` et `exemplepres` contient deux jeux fictifs, également chargeables par le bouton **Charger l'exemple** de l'écran de comparaison.

## 8. Profil client

**Outils > Profil client** regroupe ce qui est propre à un client et doit être gardé d'un passage à l'autre :

- **nom du client**, **intervenant**, **mention de confidentialité** et **logos** (prestataire et client, réduits automatiquement) pour la page de garde ;
- **groupes sensibles supplémentaires** : mots ou noms de groupes ajoutés à la détection par défaut (rapport, vue rapide « Périmètre d'un groupe à privilèges », chemins d'héritage) ;
- les **risques acceptés** et le **plan d'action** saisis dans le rapport.

Le profil est rattaché au préfixe des fichiers exportés (par exemple tous les exports `Contoso_…`) et rechargé automatiquement avec eux. Il est mémorisé sur le poste si « Mémoriser » est actif. **Exporter** produit un fichier `MapTools_Profil_<client>.json` à ranger avec les exports du client : déposé avec les CSV (ou via **Importer…**), il restaure tout sur un autre poste ou pour un collègue.

Le profil contient les noms des comptes acceptés et les justifications : il relève de la même confidentialité que les exports.

## 9. Mémorisation

Par défaut, l'outil enregistre dans le navigateur les fichiers chargés, la correspondance des colonnes et l'état de l'affichage. Un rafraîchissement ou une réouverture restaure tout.

- Un seul jeu de données est conservé : en charger un nouveau remplace le précédent.
- Désactiver **Mémoriser sur ce poste** efface immédiatement les données enregistrées ; **Tout effacer** vide aussi l'affichage.
- En navigation privée, ou si le navigateur bloque le stockage, rien n'est mémorisé.

Les exports contiennent des données personnelles (nLPD, RGPD) : effacer avant de changer de client et en fin d'intervention sur un poste partagé ou un serveur de rebond. Pour un livrable client, préférer un export PNG d'une branche ciblée à la transmission des CSV.

## 10. Dépannage

| Symptôme | Cause probable | Solution |
|---|---|---|
| Script « non signé numériquement » | Marque « provenant d'Internet » | `Unblock-File` (section 1) |
| `Connect-MgGraph` non reconnu | Module absent dans cette version de PowerShell | L'installer dans la même version |
| « Could not load file or assembly » | Modules Microsoft.Graph de versions différentes | Aligner les versions (`Get-Module Microsoft.Graph.* -ListAvailable`) |
| Tous les utilisateurs « Activité inconnue » ou « Jamais connecté » | Scope, rôle Entra ou licence P1/P2 manquant | Voir les trois conditions (section 2) ; sans P1, ajouter `Reports.Read.All` pour le secours, ou faire l'export AD |
| Pas de partie « Partages et sites » | Export sans `-AvecPartages`, ou rôle d'audit / SharePoint Administrator manquant | Relire les avertissements de `Export-MapToolsPartages.ps1` |
| « Connexion SharePoint impossible » sous PowerShell 7 | Module SharePoint Online non compatible en natif | Le script l'importe avec `-UseWindowsPowerShell` ; sinon lancer `Export-MapToolsPartages.ps1` sous Windows PowerShell 5.1 |
| « Rapport d'activité anonymisé » | Noms masqués dans les rapports Microsoft 365 | Désactiver l'anonymisation (Paramètres de l'organisation > Rapports), puis relancer |
| Rien ne s'affiche, même l'exemple | Dossier `lib\` absent ou déplacé | Le garder à côté de `MapTools.html` |
| « Aucun fichier de liens reconnu » | Colonnes non standards | Choisir le rôle et les colonnes dans la fiche |
| Beaucoup d'éléments « Absent des attributs » | Mauvaise colonne clé, ou noms différents | Vérifier la colonne « Clé » de la fiche |
| Utilisateurs dessinés en rectangles | Colonne « Type d'élément » non choisie | La sélectionner dans la fiche |
| Aucun utilisateur visible | Mode « Sélection » sans groupe sélectionné | Cliquer un groupe ou passer sur « Tous » |
| Accents illisibles | CSV enregistré en ANSI | Réexporter en UTF-8 |
| Les éléments bougent sans arrêt | Animation activée en vue Réseau | Désactiver **Animation de la carte** |
| Zoom trop rapide ou trop lent | Sensibilité de la souris | Réglage `WHEEL_K` (`REFERENCE.md`) |
| Arborescence très haute, texte illisible | Beaucoup de groupes au même niveau | Masquer les groupes sans imbrication, limiter les **Niveaux**, passer en **Compact** |
| Un groupe recherché n'apparaît pas | Masqué par la limite de niveaux ou par « Groupes sans imbrication » | La recherche le sélectionne et l'affiche quand même ; sinon repasser Niveaux sur « Tous » |
| Données d'un autre client au démarrage | Mémorisation active | **Tout effacer** |
| Rapport sans statuts ni comptes | Fichier d'attributs absent, ou colonne « Type d'élément » / « Catégorie » non choisie | Charger les deux fichiers et vérifier la fiche d'attributs |
| Comparaison : « aucun fichier de liens reconnu » | CSV d'une autre source avec des colonnes non standards | Les charger d'abord sur la carte, corriger la fiche, puis **Utiliser les fichiers chargés** |
| Comparaison : beaucoup de « nouveaux » constats d'inactivité | Temps écoulé entre les exports | Normal : ce sont des comptes devenus inactifs entre-temps |
| Pas de losanges ni de constats MFA | Export sans MFA (`-SansMFA`, module Microsoft.Graph.Reports absent, droits ou licence P1/P2 manquants), ou colonne « MFA » non choisie | Relire les avertissements du script ; vérifier la fiche d'attributs |
| Pas de rôles sur la carte (export Entra) | `-SansRoles`, module Microsoft.Graph.Identity.Governance absent ou scope `RoleManagement.Read.Directory` manquant | Relire les avertissements du script |
| Toutes les attributions en « Active », aucune « Éligible » | Pas de licence Entra ID P2, ou PIM non utilisé | Normal sans PIM ; le rapport le précise |
| Vues rapides « Rôles » ou « MFA » grisées | Donnée absente de l'export (par exemple export AD) | Faire l'export Entra |
| Licences affichées sous forme d'identifiants | Module Microsoft.Graph.Identity.DirectoryManagement absent ou scope `LicenseAssignment.Read.All` manquant | Installer le module, se reconnecter avec le scope |
| Trop ou trop peu de groupes « sensibles » | Détection par le nom | Adapter `SENSITIVE` (`REFERENCE.md`) |
