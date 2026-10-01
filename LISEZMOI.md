# MapTools

**Version 1.1** · © 2026 Jimmy Paul · conditions d'utilisation dans `LICENCE.txt`

MapTools dessine une carte interactive des **groupes, sous-groupes et utilisateurs** d'un annuaire Entra ID ou Active Directory, avec le **statut d'activité** de chaque compte. Il sert à repérer d'un coup d'œil les comptes dormants, les comptes désactivés encore membres de groupes, les invités oubliés et les imbrications problématiques.

L'outil fonctionne hors ligne dans **Edge ou Chrome** (récents), par double-clic, sans installation. Firefox n'est pas pris en charge pour l'impression (numérotation des pages). Aucune donnée ne quitte le poste.

## Démarrer en 5 étapes (Entra ID)

1. **Extraire le zip**, puis débloquer les scripts dans PowerShell, depuis le dossier `MapTools` :

   ```powershell
   Get-ChildItem .\scripts\*.ps1 | Unblock-File
   ```

2. **Installer les modules** (une seule fois, PowerShell 7 recommandé) :

   ```powershell
   Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Groups, Microsoft.Graph.Reports, Microsoft.Graph.Identity.DirectoryManagement, Microsoft.Graph.Identity.Governance -Scope CurrentUser
   ```

3. **Se connecter au tenant** :

   ```powershell
   Connect-MgGraph -Scopes "GroupMember.Read.All","User.Read.All","AuditLog.Read.All","LicenseAssignment.Read.All","RoleManagement.Read.Directory","Reports.Read.All" -TenantId "<domaine du client>" -NoWelcome
   ```

4. **Exporter** :

   ```powershell
   .\scripts\Export-MapToolsEntra.ps1 -Prefixe <Client> -DossierExport C:\Temp
   ```

   Ajouter `-AvecPartages` pour exporter aussi les partages SharePoint / OneDrive et les paramètres de partage des sites (modules et rôles supplémentaires, voir le guide).

5. **Ouvrir `MapTools.html`** et y déposer tous les CSV produits en même temps (deux, ou quatre avec les partages).

Pour un Active Directory on-prem, remplacer les étapes 2 à 4 par `.\scripts\Export-MapToolsAD.ps1 -Prefixe <Client> -DossierExport C:\Temp` (module ActiveDirectory requis). L'outil reconnaît la source : avec un export AD, les contrôles propres à Entra (MFA, licences, rôles, propriétaires) sont signalés comme non réalisés, et les contrôles de mots de passe AD s'ajoutent. Pour découvrir l'outil sans données réelles : bouton **Voir un exemple**.

## Lire la carte en bref

- **Rectangles** = groupes, colorés par type. **Points** = utilisateurs, colorés par statut : vert actif, turquoise nouveau, jaune jamais connecté, orange inactif, rouge désactivé. **Losanges** = comptes activés sans MFA enregistrée. **Rectangles foncés** = rôles d'administration Entra (trait plein : attribution active ou permanente, pointillés : éligible via PIM).
- Les racines sont à gauche, chaque colonne vers la droite est un niveau d'imbrication.
- **Écran** : en haut, la recherche, la vue, les vues rapides, le rapport et le menu **Outils** (avant / après, exports, profil client, PNG) ; à gauche, les panneaux repliables Données, Affichage et Légende ; sur la carte, le bandeau des filtres actifs avec **Tout réinitialiser** ; en bas, les chiffres clés, cliquables. Le survol d'un élément affiche une infobulle.
- **Clic** sur un élément : ses détails s'ouvrent dans un panneau à droite (bouton × pour le fermer). **Double-clic** : n'affiche que sa branche.
- Par défaut, seuls les utilisateurs du groupe sélectionné sont affichés. Le sélecteur **Utilisateurs** permet de tous les afficher ou de les masquer.
- Un clic sur une ligne de la **légende** masque ou réaffiche cette catégorie.
- **Vue rapide** : cinq réglages prêts à l'emploi (structure seule, comptes à traiter, désactivés encore membres, périmètre d'un groupe à privilèges, groupes isolés).
- **Partages et sites** : liens « Tout le monde », partages vers des externes, sites ouverts au partage anonyme, partages reçus par chaque invité (journal d'audit, 180 jours au plus).
- **Pourquoi a-t-il accès ?** : depuis un compte ou un groupe, la carte surligne les chemins d'héritage exacts jusqu'à un rôle ou un groupe sensible.
- **Profil client** : logos, intervenant, groupes sensibles, risques acceptés et plan d'action, gardés d'un passage à l'autre et exportables en JSON.
- **Exporter des listes** : utilisateurs avec ou sans MFA, membres d'un groupe (directs ou tous niveaux), groupes d'un utilisateur, détenteurs d'un rôle… en CSV ou copiés pour Excel, avec des filtres combinables.
- **Rapport décisionnel** : page de garde, plan d'action pilotable (responsable, échéance, statut), risques acceptés justifiés, synthèse chiffrée, constats priorisés avec l'action recommandée (dont MFA et licences inutilisées, avec l'économie estimée), impression en PDF et plan d'action en CSV.
- **Comparer avant / après** : deux exports datés, et l'outil montre ce qui a été traité, ce qui reste et ce qui est nouveau.

## Trois usages courants en audit

- **Comptes désactivés encore membres de groupes** : dans la légende, masquer tous les statuts sauf « Désactivé », puis passer les utilisateurs sur « Tous ».
- **Comptes d'administration dormants** : rechercher le groupe d'administration, double-cliquer dessus et lire la répartition des statuts dans le panneau.
- **Invités inactifs** : sélectionner un groupe M365 ou une équipe et lire la liste de ses utilisateurs directs.

## Confidentialité

Les exports contiennent des **données personnelles** (noms, UPN, dates de connexion). Les conserver dans le dossier du client et les supprimer en fin de mission. L'outil mémorise les données dans le navigateur : cliquer sur **Tout effacer** avant de changer de client, et en fin d'intervention sur un poste partagé.

## Aller plus loin

- `docs\GUIDE.md` : utilisation complète (scripts et prérequis, statuts, affichage, mémorisation, dépannage).
- `docs\REFERENCE.md` : adapter l'outil (paramètres, formats CSV, colonnes, réglages, pièges connus, intégrité du paquet).

## Contenu du dossier

| Élément | Rôle |
|---|---|
| `MapTools.html` | L'outil |
| `scripts\Export-MapToolsEntra.ps1` | Export Entra ID : groupes, imbrications, utilisateurs et statut |
| `scripts\Export-MapToolsAD.ps1` | Export Active Directory : même contenu |
| `scripts\Export-MapToolsPartages.ps1` | Partages SharePoint / OneDrive et paramètres de partage des sites (seul, ou via `-AvecPartages`) |
| `scripts\Get-EntraIDComptesInactifs.ps1` | Audit seul des comptes inactifs et désactivés |
| `docs\` | Guide et référence |
| `exemple\` | CSV fictifs, dont un jeu avant / après remédiation |
| `lib\` | Bibliothèques embarquées (à garder à côté du HTML) |
| `dev\` | Sources, script d'assemblage et tests, pour faire évoluer l'outil (`dev\LISEZMOI-dev.md`) |
| `LICENCE.txt` · `SHA256SUMS.txt` | Conditions d'utilisation · empreintes du paquet original |
