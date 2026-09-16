# 🛠️ Scripts d'Administration & Audit Atlassian Cloud (Jira & Confluence)

Ce référentiel regroupe les outils d'automatisation, d'audit et d'administration développés en **PowerShell** pour la gestion de la gouvernance, des habilitations et des migrations de groupes sur **Jira Cloud** et **Confluence Cloud** (*instance jiradot*).

---
👤 Auteur & Maintenance
Auteur : Frédéric GUEDJ
Équipe : HM_DSIM_PACT

Organisation : Harmonie Mutuelle / Groupe VYV

## 📋 Sommaire

- [Vue d'ensemble des scripts](#-vue-densemble-des-scripts)
- [Architecture du projet](#-architecture-du-projet)
- [Prérequis techniques](#-prérequis-techniques)
- [Sécurité & Données locales (.gitignore)](#-sécurité--données-locales-gitignore)
- [Guide d'utilisation](#-guide-dutilisation)
  - [1. Comparaison nominale des groupes (Compare-JiraGroups.ps1)](#1-comparaison-nominale-des-groupes-compare-jiragroupsps1)
  - [2. Audit différentiel des habilitations (Audit-GroupPermissions-DryRun.ps1)](#2-audit-différentiel-des-habilitations-audit-grouppermissions-dryrunps1)
  - [3. Diagnostic des espaces Confluence (Test-ConfluenceSpacePermissionsDiagnostic.ps1)](#3-diagnostic-des-espaces-confluence-test-confluencespacepermissionsdiagnosticps1)
  - [4. Outil de diagnostic Proxy & Réseau (Get-ProxyConfiguration.ps1)](#4-outil-de-diagnostic-proxy--réseau-get-proxyconfigurationps1)
- [Endpoints REST Atlassian utilisés](#-endpoints-rest-atlassian-utilisés)

---

## 🔍 Vue d'ensemble des scripts

```text
Scripts Atlassian/
+--- backup/
+--- sheets/
|    +--- sheets/
|    +--- Common.ps1
|    +--- Sheet1-WorklogsIssues.ps1
|    +--- Sheet2-SaisiTempsTempo.ps1
|    +--- Sheet3-SaisiTempsAsset.ps1
|    +--- Sheet4-Anomalies.ps1
|    +--- Sheet5-WorklogsSansDate.ps1
+--- .mcp.json
+--- Ajout-Groupe-Espaces-Confluence.ps1
+--- analyse-budget.ps1
+--- Analyse-Comptes-Desactives-Suppression.ps1
+--- Analyse-Comptes-Jamais-Utilises.ps1
+--- Analyse-InactifsAssets.ps1
+--- Analyze-AuditLog-UsersSansApps.ps1
+--- Apply-DSIMRoleMatrix-ToProjects.ps1
+--- Audit-AutomationsEtAuteursAssets.ps1
+--- Audit-GroupPermissions-DryRun.ps1
+--- Audit-ObjetsAssetsObsoletes.ps1
+--- Audit-ParametrageJira.ps1
+--- Audit-SchemaExportMensuel.ps1
+--- Audit-ShadowITInstances.ps1
+--- Audit-TypesTickets-Jira.ps1
+--- audit-users-assets.ps1
+--- Audit-VolumeHistoriqueAssets.ps1
+--- Check-AuditInvite-Overlap.ps1
+--- Common.ps1
+--- Compare-JiraGroups.ps1
+--- Debug-AssetsSchema.ps1
+--- Desactivation-Comptes-Sans-Produit.ps1
+--- Desactivation-NoAccess.ps1
+--- Desactivation-Suppression-Users-sav.ps1
+--- Desactivation-Suppression-Users-v2.ps1
+--- Desactivation-Suppression-Users.ps1
+--- detachInactiveUsersFromProjects - Copie.ps1
+--- detachInactiveUsersFromProjects.ps1
+--- Detection-DonneesSensibles-Jira.ps1
+--- Diag-Projet-Jira.ps1
+--- Diagnostic-AttributsAssetsInactifs.ps1
+--- Disable-InactiveJiraUsers.ps1
+--- Doublons-Groupes-Sites.ps1
+--- dsimGroupesVsCategorieParProjet.ps1
+--- dsimNominatifsParProjet.ps1
+--- dsimRoleHierarchyCleanup.ps1
+--- dsimUsersPerProject.ps1
+--- Ensure-SiteAdmin-ProjectManagerRole.ps1
+--- Export-DSIMGroups_x_ProjectCategories.ps1
+--- Export-Scan-OrgAuditTokens.ps1
+--- Export-Users-Groups-Teams.ps1
+--- Gestion-Datasources-Analytics.ps1
+--- Get-AddedToOrg-ForUserSansApps.ps1
+--- Get-AssetsAccessUsers.ps1
+--- Get-AssetsAndJiraGroupsByEmail.ps1
+--- Get-AssetsJiraUsers.ps1
+--- Get-InviteDate-ForUserSansApps.ps1
+--- Get-JiraSuperAdmins.ps1
+--- Get-ProjectTree.ps1
+--- Get-ShadowITUsersHM.ps1
+--- ImportExcel.ps1
+--- Inspect-AuditLogCsv.ps1
+--- Install-ImportExcel.ps1
+--- Invoke-EtatsFinanciers_v2.ps1
+--- Invoke-EtatsFinanciers_v3.ps1
+--- Invoke-EtatsFinanciers.ps1
+--- journalAuditConfluence.ps1
+--- lecture-directory-users.ps1
+--- List-ProjectManagers-ByCategory.ps1
+--- List-ProjectRoles.ps1
+--- listeGroupesDsimMembres.ps1
+--- listeGroupesDsimMembresInactifs.ps1
+--- listeHabilitationsJira.ps1
+--- listeprojetsJira5.ps1
+--- Listing-AnomaliesAccountBudget.ps1
+--- Listing-AnomaliesComptesAssets - Copie.ps1
+--- Listing-AnomaliesComptesAssets.ps1
+--- Listing-Hab-Confluence2.ps1
+--- Listing-Habilitations-Confluence.ps1
+--- Manage-IssueLinkTypes.ps1
+--- Migrate-ProjectRoles.ps1
+--- mvp-reporting_v2.ps1
+--- Nettoyage-TypesProjetScoped-Jira.ps1
+--- Nettoyage-TypesTickets-Jira.ps1
+--- NoApps-Cutoff-VerifyAndSuspend.ps1
+--- NoApps-LastActive-Cutoff.ps1
+--- Purge-ObjetsAssetsObsoletes.ps1
+--- Rationalisation-TypesTickets-Jira.ps1
+--- Rattachement-Equipe-Jira.ps1
+--- README.md
+--- Recherche-Filtres-Jira.ps1
+--- Remove-DSIM-NominativeAssignments.ps1
+--- removeInactiveUsersFromDsimGroups.ps1
+--- Restore-ObjetsAssets.ps1
+--- Save-AtlassianOrgAdminKey.ps1
+--- Save-JiraCredential.ps1
+--- Scan-RGPD-Estimation.ps1
+--- Scan-RGPD-Jira-Part1.ps1
+--- Scan-RGPD-Jira-Part2-PJ.ps1
+--- scanTokensAudit-deep.ps1
+--- scanTokensAudit.ps1
+--- Sheet1-WorklogsIssues.ps1
+--- Sheet2-SaisiTempsTempo.ps1
+--- Sheet3-SaisiTempsAsset.ps1
+--- Sheet4-Anomalies.ps1
+--- Sheet5-WorklogsSansDate.ps1
+--- Sheet6-AnalyseParLigneBudgetaire.ps1
+--- Sheet7-AnalyseParInitiativeLotEpic.ps1
+--- Suppression-Comptes-Desactives.ps1
+--- Suspend-OldUsers-WithNoApps.ps1
+--- Test-ConfluenceSpacePermissionsDiagnostic.ps1
```


---

## 📂 Architecture du projet

Structure versionnée du dépôt GitHub (conforme aux exclusions `.gitignore`) :

```text
Scripts-Atlassian/
├── .gitignore                                      # Définition des dossiers et fichiers exclus du versionnement
... scripts ps1
└── README.md                                       # Documentation du référentiel

Les dossiers secrets/, exports/, cache/, logs/, output/ et scripts/ sont réservés à l'exécution locale sur le poste de travail et sont strictement exclus du dépôt distant.

⚙️ Prérequis techniques
PowerShell 5.1 (Windows PowerShell) ou PowerShell 7+ (Core).

Droits d'administration Jira et Confluence sur l'instance cible.

Support du Proxy d'entreprise avec authentification intégrée Windows (Kerberos/NTLM).

Protocole TLS 1.2 / 1.3 activé.

# Credentials et identifiants
secrets/

# Fichiers de travail temporaires
cache/
logs/
output/
scripts/

# Résultats générés
exports/

# À exécuter une seule fois sur votre machine
$cred = Get-Credential # Saisir l'e-mail en Utilisateur et l'API Token en Mot de passe
$jiraConfig = [pscustomobject]@{
    JiraBaseUrl = "https://jiradot.atlassian.net"
    Credential  = $cred
}
if (-not (Test-Path ".\secrets")) { New-Item -ItemType Directory -Path ".\secrets" }
$jiraConfig | Export-Clixml -Path ".\secrets\jira-jiradot.cred.xml"

