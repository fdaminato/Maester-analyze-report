# Tableau de bord d’évaluation de la sécurité

Les chemins du projet indiqués ci-dessous sont relatifs au dossier `Script` contenant ce README.

## Lancer l’évaluation

Depuis ce dossier, dans PowerShell :

```powershell
./launch.ps1
```

Ou lancer directement le script principal :

```powershell
./Invoke-MaesterModernDashboard.ps1 -PromptModuleUpdate -UpdateTests -IncludePreview -OpenReport
```

Le script principal (`# V4`) demande le **UserPrincipalName** de l’administrateur et le **TenantId** cible lorsque ces paramètres sont omis. Une valeur vide ou invalide saisie à l’invite entraîne une nouvelle demande. Une valeur invalide fournie en paramètre interrompt l’exécution avant l’authentification.

Le lanceur (`# V3`) ne contient aucun compte ou identifiant de tenant prédéfini. Il inclut les tests en aperçu et ouvre le rapport généré. Les paramètres `-UserPrincipalName`, `-TenantId`, `-TestsPath` et `-OutputRoot` restent disponibles pour les lancements automatisés.

## Trouver le dossier de tests

- **Windows** : `Join-Path $env:USERPROFILE 'maester-tests'`, pour l’utilisateur Windows actuel.
- **macOS** : `Join-Path (Get-Location).Path 'maester-tests'`, équivalent à `$PWD/maester-tests`.
- Si le dossier attendu ou fourni est absent, inaccessible ou ne contient aucun fichier `*.Tests.ps1`, saisir un autre chemin complet à l’invite. Saisir `Q` pour annuler avant l’authentification et la collecte des résultats.
- Le lanceur utilise son propre dossier comme répertoire de travail. Sur macOS, il cherche donc le sous-dossier `maester-tests` à cet emplacement. Un lancement direct utilise le répertoire PowerShell actuel; ce chemin est conservé lors du redémarrage dans un processus propre.

## Mise à jour facultative du module

Après la validation du dossier de tests, `launch.ps1` active l’option `-PromptModuleUpdate` du script principal. L’invite affichée est :

```text
Run Update-Module Maester -Force before continuing? (Y/N)
```

- **`Y`** : exécuter `Update-Module Maester -Force -ErrorAction Stop`, puis continuer.
- **`N`** : conserver le module installé et poursuivre la mise à jour des tests et l’évaluation.
- **Autre réponse** : répéter la demande. Les majuscules, les minuscules et les espaces en début ou fin de réponse sont acceptés.

Ce choix précède l’importation de Maester et la mise à jour des fichiers de tests. La version installée est vérifiée de nouveau après le choix afin d’importer la plus récente. Une erreur de mise à jour interrompt le traitement avant la mise à jour des tests, l’authentification et la création du rapport.

Omettre `-PromptModuleUpdate` lors d’un lancement direct ou automatisé pour éviter cette invite. La reconstruction de résultats enregistrés ne l’affiche jamais.

## Mettre les tests à jour avant l’évaluation

`launch.ps1` active l’option `-UpdateTests`. Après avoir validé le dossier sélectionné, le processus PowerShell entre dans ce dossier et exécute :

```powershell
Update-MaesterTests -Path '.' -ErrorAction Stop
```

Le répertoire de travail précédent est rétabli avec `finally` et `Pop-Location`, même en cas d’erreur. Les fichiers de tests sont ensuite recensés de nouveau pour tenir compte des ajouts et suppressions.

Une erreur interrompt l’évaluation avant l’authentification ou la création du rapport. La confirmation d’écrasement prévue par Maester demeure active : aucun `-Force` n’est fourni pour cette commande. Omettre `-UpdateTests` lors d’un lancement direct pour conserver les tests existants. La reconstruction d’un rapport ne met jamais les tests installés à jour.

## Nom des rapports clients

Les nouveaux résultats sont enregistrés par défaut dans le dossier du script principal. Le paramètre `-OutputRoot` permet de choisir un autre emplacement.

Après l’authentification, le domaine initial vérifié du tenant, en `onmicrosoft.com`, fournit le préfixe court. S’il n’est pas disponible, le script utilise le domaine du compte de connexion en onmicrosoft, puis le nom du tenant ou son identifiant, adaptés aux noms de fichiers Windows. Les caractères interdits sont retirés. Un dossier déjà présent pour la même seconde n’est jamais réutilisé.

Exemple, avec la date et l’heure locales au début des tests :

```text
itimwlabc-20260917-105400/
  itimwlabc-20260917-105400.html
  Maester-Raw.html
  Maester-Raw.json
  ...autres rapports internes et journal de génération...
```

L’export CSV du tableau de bord porte le nom `itimwlabc-20260917-105400.csv`.

Le HTML autonome et son export CSV ne contiennent aucune mention, URL ou étiquette Maester, aucun chemin source local de l’opérateur, aucune version du module ni lien vers le rapport brut. Le titre ITI365, les identifiants des contrôles, les références aux normes, les constats, les statuts, les niveaux de gravité, les durées, les scores et les explications des tests ignorés ou en erreur sont conservés. Les scripts et rapports internes conservent leurs références originales.

Les données destinées au client sont copiées avant le nettoyage; les résultats bruts ne sont pas modifiés. Une vérification finale bloque l’écriture du HTML s’il reste une mention Maester. En cas d’échec de génération, le HTML affiche une indication générale destinée au client; les détails techniques restent dans les diagnostics internes.

## Reconstruire un rapport à partir des résultats enregistrés

```powershell
./Invoke-MaesterModernDashboard.ps1 -ExistingRunFolder './itimwlabc-20260917-105400'
```

Cette opération ne demande ni compte ni emplacement des tests et ne déclenche aucune authentification. Les dossiers portant déjà un nom de tenant conservent leur nom. Les anciens dossiers contenant seulement une date restent en place; le HTML reconstruit reçoit le préfixe du tenant et l’horodatage du dossier d’origine.

Les rapports existants n’ont pas été renommés lors des modifications du script. Le rapport enregistré précédemment demeure accessible à `2026-09-17_104954/Security-Assessment.html`.

## Sauvegardes et validation

Les sauvegardes ci-dessous restent locales et ne sont pas incluses dans le dépôt Git :

- `Backup-Before-Module-Prompt-2026-09-17_111514.zip` : quatre fichiers avant le lanceur V3 et le script principal V4; contrôles CRC et SHA-256 réussis avant les modifications.
- `Backup-Before-Test-Update-2026-09-17_111238.zip` : lanceur, script principal, README et journal avant le lanceur V2 et le script principal V3; contrôles réussis.
- `Backup-Before-Interactive-Launch-2026-09-17_105921.zip` : script principal, lanceur, README et journal avant les modifications V2; contrôles réussis pour les quatre fichiers.
- `Backup-Before-Client-Report-2026-09-17_104300.zip` : ancien script et six fichiers du rapport initial; archive conservée sans modification.

**Vérifié sur macOS :** analyse syntaxique des deux scripts; sélection du chemin Windows de l’utilisateur actuel et du chemin macOS réel; détection des dossiers, nouvelles demandes, chemins entre guillemets et annulation; validation des comptes et tenants saisis ou fournis; noms basés sur le domaine initial, domaines de connexion personnalisés et noms de remplacement compatibles Windows.

Les vérifications comprenaient aussi l’enchaînement complet d’une nouvelle exécution avec appels cloud simulés, les noms correspondants des dossiers et fichiers HTML/CSV, la prévention des collisions et le lanceur réel avec un exécutable simulé : choix du processus PowerShell, options, restauration du répertoire et transmission des erreurs. Le redémarrage réel dans un processus propre et l’annulation en cas de dossier absent ont été vérifiés.

Deux reconstructions réelles ont utilisé des copies temporaires des résultats enregistrés. Les 415 résultats, les décomptes, les scores et les fichiers bruts ont été préservés. Le HTML client ne contenait aucune mention Maester ou chemin de profil local. Le nettoyage, la vérification finale, la syntaxe JavaScript, son exécution avec un DOM simulé et le nom/contenu CSV ont été vérifiés.

**V3 — mise à jour des tests :** avec appels de mise à jour et cloud simulés, vérification du dossier de travail exact, du recensement des nouveaux tests, de la restauration du répertoire après succès ou erreur, de l’arrêt avant authentification en cas d’erreur et de l’absence de mise à jour sans l’option ou après annulation. Le lanceur renommé transmet correctement les options.

**V4 — mise à jour facultative du module :** avec appels simulés, vérification de `Y`, de `-Force`, de l’ordre des opérations et de la sélection de la version installée; `N` poursuit les tests et l’évaluation. Les réponses invalides sont redemandées et une erreur bloque les étapes suivantes. Les deux options sont transmises au processus enfant.

**Non vérifié :** exécution native sous Windows PowerShell 5.1/7 ou ISE, affichage visuel dans un navigateur, mises à jour réelles du module/des tests ou nouvelle évaluation d’un tenant. Aucune authentification cloud ni collecte réelle n’a été effectuée pendant ces validations.

PowerShell 7 demeure recommandé. Le lanceur privilégie `pwsh` et utilise `powershell.exe` comme solution de remplacement sous Windows. Depuis ISE, le redémarrage utilise un processus PowerShell en ligne de commande. Les fonctions ajoutées utilisent une syntaxe compatible avec Windows PowerShell 5.1; le paramètre JSON Depth est fourni seulement s’il est pris en charge.

Journal local détaillé : `Logs/2026-09-17-Client-Report.md`, exclu de l’historique Git.

Références pour les noms de tenants : Microsoft Graph [organisation](https://learn.microsoft.com/en-us/graph/api/resources/organization?view=graph-rest-1.0) et [domaine vérifié](https://learn.microsoft.com/en-us/graph/api/resources/verifieddomain?view=graph-rest-1.0).

## Dépôt GitHub privé

Le dépôt [pierreluc102/MaesterITI](https://github.com/pierreluc102/MaesterITI) est privé. La branche principale est `main`.

Seuls `Invoke-MaesterModernDashboard.ps1`, `launch.ps1`, `README.md` et `.gitignore` sont versionnés. La liste explicite de fichiers autorisés dans `.gitignore` conserve les rapports clients, le dossier installé `maester-tests`, les sauvegardes ZIP et les journaux locaux sur disque, hors des commits.

Ajouter tout nouveau fichier source à cette liste avant de le versionner. Installer les tests Maester séparément sur chaque ordinateur.

Références GitHub : [créer un dépôt](https://docs.github.com/en/repositories/creating-and-managing-repositories/creating-a-new-repository) et [ignorer des fichiers](https://docs.github.com/en/get-started/git-basics/ignoring-files).

## Donner accès aux utilisateurs

### Dépôt appartenant à un compte personnel

Pour inviter une personne : ouvrir le dépôt, puis **Settings → Collaborators → Add people**. L’invitation peut cibler son compte GitHub ou son adresse courriel.

Sur un dépôt privé personnel, les collaborateurs reçoivent un accès en lecture et en écriture. Ce type de dépôt ne permet pas d’attribuer un accès en lecture seule. Voir les [permissions des dépôts personnels](https://docs.github.com/en/repositories/managing-your-repositorys-settings-and-features/repository-access-and-collaboration/permission-levels-for-a-personal-account-repository) et les [étapes d’invitation](https://docs.github.com/en/repositories/managing-your-repositorys-settings-and-features/repository-access-and-collaboration/inviting-collaborators-to-a-personal-repository).

### Organisation GitHub pour des rôles distincts

Pour distinguer les utilisateurs, contributeurs et administrateurs, utiliser un dépôt privé appartenant à une organisation. Les équipes regroupent les personnes ayant besoin du même rôle.

| Rôle GitHub | Utilisation |
| --- | --- |
| Read | Consulter et télécharger les scripts. |
| Triage | Gérer les demandes et discussions sans modifier le code. |
| Write | Modifier le code et envoyer des changements. |
| Maintain | Gérer le dépôt sans certaines opérations sensibles réservées aux administrateurs. |
| Admin | Administrer le dépôt et les accès. |

Étapes proposées :

1. Créer une organisation GitHub ou utiliser celle de l’entreprise.
2. Transférer le dépôt via **Settings → General → Danger Zone → Transfer ownership**, en conservant sa visibilité privée.
3. Dans l’organisation, choisir **Settings → Member privileges → Base permissions → None** pour attribuer explicitement l’accès à chaque dépôt aux membres concernés.
4. Dans le dépôt, utiliser **Settings → Collaborators & teams → Add people / Add teams**, puis choisir le rôle.

Les permissions de lecture portent sur le dépôt entier. Pour limiter la visibilité de projets à différents groupes, utiliser des dépôts distincts. Les responsables d’organisation conservent l’accès administratif à tous ses dépôts.

Références : [rôles d’un dépôt d’organisation](https://docs.github.com/en/organizations/managing-user-access-to-your-organizations-repositories/managing-repository-roles/repository-roles-for-an-organization), [permissions de base](https://docs.github.com/en/organizations/managing-user-access-to-your-organizations-repositories/managing-repository-roles/setting-base-permissions-for-an-organization), [gestion des personnes et équipes](https://docs.github.com/en/repositories/managing-your-repositorys-settings-and-features/managing-repository-settings/managing-teams-and-people-with-access-to-your-repository) et [transfert d’un dépôt](https://docs.github.com/en/repositories/creating-and-managing-repositories/transferring-a-repository).
