# ============================================================
# PHASE B : MARQUAGE DES HOMONYMES GLOBAUX
# Ajoute un commentaire en debut de description avec le type cible
# Aucune migration automatique — l'admin migre manuellement apres verification
# ============================================================

if ($doPhaseB) {
  Write-Host ""
  Write-Host "========================================" -ForegroundColor DarkYellow
  Write-Host "  PHASE B : MARQUAGE HOMONYMES GLOBAUX" -ForegroundColor DarkYellow
  Write-Host "  (ajout commentaire en description)" -ForegroundColor DarkYellow
  Write-Host "========================================" -ForegroundColor DarkYellow
  Log "=== PHASE B : Marquage homonymes globaux (description) ==="

  $cFusionGroups = $fusionGroups.Count

  if ($cFusionGroups -eq 0) {
    Write-Host "  Aucun homonyme global a traiter." -ForegroundColor DarkGray
    Log "  Aucun homonyme global a traiter."
  } else {
    Write-Host ("  {0} groupes d'homonymes :" -f $cFusionGroups) -ForegroundColor DarkYellow

    foreach ($k in ($fusionGroups.Keys | Sort-Object)) {
      $members = $fusionGroups[$k] | Sort-Object { $_.IssueCount } -Descending
      $target  = $members[0]
      $sources = $members | Select-Object -Skip 1

      $targetLabel = "'{0}' (ID:{1}, {2} issues)" -f $target.Name, $target.Id, $target.IssueCount
      Write-Host ("    Groupe '{0}' -> type cible {1}" -f $k, $targetLabel) -ForegroundColor DarkYellow

      foreach ($src in $sources) {
        $srcLabel = "'{0}' (ID:{1}, {2} issues)" -f $src.Name, $src.Id, $src.IssueCount
        if ($src.IssueCount -eq 0) {
          Write-Host ("      SUPPRIMER {0} (0 issue)" -f $srcLabel) -ForegroundColor Red
        } else {
          Write-Host ("      MARQUER   {0} issues" -f $src.IssueCount) -ForegroundColor DarkYellow
        }
      }
    }

    $confirmB = [System.Windows.Forms.MessageBox]::Show(
      ("Traiter {0} groupes d'homonymes ?`n`nPour chaque issue du type doublon :`n- Ajout d'un marqueur [FUSION] en debut de description`n- Aucune migration automatique`n`nLes types doublons avec 0 issue seront supprimes." -f $cFusionGroups),
      "Phase B - Confirmation",
      [System.Windows.Forms.MessageBoxButtons]::YesNo,
      [System.Windows.Forms.MessageBoxIcon]::Warning)

    if ($confirmB -eq [System.Windows.Forms.DialogResult]::Yes) {
      foreach ($k in ($fusionGroups.Keys | Sort-Object)) {
        $members = $fusionGroups[$k] | Sort-Object { $_.IssueCount } -Descending
        $target  = $members[0]
        $sources = $members | Select-Object -Skip 1
        $normalizedName = $k

        foreach ($src in $sources) {
          if ($src.IssueCount -gt 0) {
            Log ("  Marquage {0} (ID:{1}, {2} issues)..." -f $src.Name, $src.Id, $src.IssueCount)
            Write-Host ("    Marquage {0} (ID:{1}, {2} issues)..." -f $src.Name, $src.Id, $src.IssueCount) -ForegroundColor DarkYellow

            $marked = 0; $markErrors = 0; $alreadyMarked = 0
            $searchUrl = "{0}/rest/api/3/search/jql" -f $site.BaseUrl
            $moreIssues = $true
            $nextPageToken = $null

            while ($moreIssues) {
              $jqlText = "issuetype = {0} ORDER BY key ASC" -f $src.Id
              $bodyObj = @{ jql = $jqlText; maxResults = 50; fields = @("key","description","project") }
              if ($nextPageToken) { $bodyObj["nextPageToken"] = $nextPageToken }
              $bodyJson = $bodyObj | ConvertTo-Json -Depth 5 -Compress
              $searchResp = Invoke-ApiCall -Method "POST" -Url $searchUrl -Headers $site.Headers -Body $bodyJson
              if (-not $searchResp.ok) { Log "    Erreur recherche" "ERROR"; break }
              $searchJson = $searchResp.content | ConvertFrom-Json
              $issues = $searchJson.issues
              if (($issues | Measure-Object).Count -eq 0) { $moreIssues = $false; break }

              foreach ($issue in $issues) {
                $issueKey = [string]$issue.key
                $issueProject = $issueKey -replace '-.*',''

                # --- DETERMINER LE TYPE CIBLE POUR CE PROJET ---
                $acceptedTypes = Get-ProjectAcceptedTypes $issueProject
                $localTargetId = $null
                $localTargetName = ""

                # Chercher un type homonyme dans le scheme du projet
                if ($acceptedTypes.ContainsKey($normalizedName)) {
                  $localTarget = $acceptedTypes[$normalizedName]
                  if ($localTarget.Id -ne $src.Id) {
                    $localTargetId = $localTarget.Id
                    $localTargetName = $localTarget.Name
                  }
                }

                # Fallback : target global
                if (-not $localTargetId) {
                  $targetNorm = Normalize-TypeName $target.Name
                  if ($acceptedTypes.ContainsKey($targetNorm) -and $acceptedTypes[$targetNorm].Id -ne $src.Id) {
                    $localTargetId = $acceptedTypes[$targetNorm].Id
                    $localTargetName = $acceptedTypes[$targetNorm].Name
                  }
                }

                # Si aucun type cible trouve, marquer quand meme avec le target global
                if (-not $localTargetId) {
                  $localTargetId = $target.Id
                  $localTargetName = $target.Name
                }

                # --- CONSTRUIRE LE MARQUEUR ---
                $marker = "[FUSION] Type cible: {0} (ID:{1}) | Type actuel: {2} (ID:{3}) | Projet: {4}" -f `
                  $localTargetName, $localTargetId, $src.Name, $src.Id, $issueProject

                # --- VERIFIER SI DEJA MARQUE ---
                $currentDesc = ""
                if ($issue.fields -and $issue.fields.description) {
                  # La description est en ADF (Atlassian Document Format)
                  # On la convertit en texte pour verifier
                  try {
                    $descJson = $issue.fields.description
                    if ($descJson.content) {
                      foreach ($block in $descJson.content) {
                        if ($block.content) {
                          foreach ($inline in $block.content) {
                            if ($inline.text) { $currentDesc += $inline.text }
                          }
                        }
                      }
                    }
                  } catch {
                    $currentDesc = [string]$issue.fields.description
                  }
                }

                if ($currentDesc -match "\[FUSION\]") {
                  $alreadyMarked++
                  continue
                }

                # --- PREPARER LA NOUVELLE DESCRIPTION (ADF) ---
                # Ajouter le marqueur comme premier paragraphe
                $markerNode = @{
                  type = "paragraph"
                  content = @(
                    @{
                      type = "text"
                      text = $marker
                      marks = @(
                        @{ type = "strong" }
                        @{
                          type = "textColor"
                          attrs = @{ color = "#FF8B00" }
                        }
                      )
                    }
                  )
                }

                $separatorNode = @{
                  type = "rule"
                }

                # Construire la nouvelle description : marqueur + separateur + description existante
                $newDescContent = @($markerNode, $separatorNode)

                if ($issue.fields -and $issue.fields.description -and $issue.fields.description.content) {
                  foreach ($existingBlock in $issue.fields.description.content) {
                    $newDescContent += $existingBlock
                  }
                }

                $newDescription = @{
                  version = 1
                  type    = "doc"
                  content = $newDescContent
                }

                # --- METTRE A JOUR LA DESCRIPTION ---
                $updateBody = @{
                  fields = @{
                    description = $newDescription
                  }
                } | ConvertTo-Json -Depth 20 -Compress

                $updateUrl = "{0}/rest/api/3/issue/{1}" -f $site.BaseUrl, $issueKey
                $updateResp = Invoke-ApiCall -Method "PUT" -Url $updateUrl -Headers $site.Headers -Body $updateBody

                if ($updateResp.ok -or $updateResp.status -eq 204) {
                  $marked++
                } else {
                  $markErrors++
                  $errMsg = Get-JiraErrorMessage $updateResp.body
                  $failDetail = "status={0}, issue={1}, erreur={2}" -f $updateResp.status, $issueKey, $errMsg
                  Log ("    ERREUR marquage {0} : {1}" -f $issueKey, $failDetail) "ERROR"
                  Write-ActionCsv "B" "MARQUER-ECHEC" $src.Name $src.Id "Global" 1 $localTargetName $localTargetId $issueProject $issueKey "ERREUR" $failDetail
                }

                if (($marked + $markErrors + $alreadyMarked) % 50 -eq 0) {
                  Write-Progress -Activity "Marquage" -Status ("{0} marques, {1} deja marques, {2} erreurs" -f $marked, $alreadyMarked, $markErrors)
                }
              }

              # Pagination
              if ($searchJson.nextPageToken) {
                $nextPageToken = [string]$searchJson.nextPageToken
              } else {
                $moreIssues = $false
              }

              Start-Sleep -Milliseconds 200
            }
            Write-Progress -Activity "Marquage" -Completed

            Log ("    Marquage termine : {0} marques, {1} deja marques, {2} erreurs" -f $marked, $alreadyMarked, $markErrors)
            Write-Host ("      {0} marques, {1} deja marques, {2} erreurs" -f $marked, $alreadyMarked, $markErrors) -ForegroundColor Cyan

            $markDetail = "{0} marques, {1} deja marques, {2} erreurs" -f $marked, $alreadyMarked, $markErrors
            $markResult = if ($markErrors -eq 0) { "OK" } elseif ($marked -gt 0) { "PARTIEL" } else { "ECHEC" }
            Write-ActionCsv "B" "MARQUER" $src.Name $src.Id "Global" $src.IssueCount $target.Name $target.Id "" "" $markResult $markDetail
          }

          # Supprimer les doublons avec 0 issue (ceux-la n'ont pas besoin de marquage)
          if ($src.IssueCount -eq 0) {
            $delUrl = "{0}/rest/api/3/issuetype/{1}" -f $site.BaseUrl, $src.Id
            $delResp = Invoke-ApiCall -Method "DELETE" -Url $delUrl -Headers $site.Headers
            if ($delResp.ok -or $delResp.status -eq 204) {
              Log ("    SUPPRIME : {0} (ID:{1}) [0 issue]" -f $src.Name, $src.Id)
              Write-Host ("      SUPPRIME : {0} (0 issue)" -f $src.Name) -ForegroundColor Green
              Write-ActionCsv "B" "SUPPRIMER" $src.Name $src.Id "Global" 0 $target.Name $target.Id "" "" "OK" "Supprime (0 issue)"
            } else {
              $errMsg = Get-JiraErrorMessage $delResp.body
              $errDetail = "status={0}, erreur={1}" -f $delResp.status, $errMsg
              Log ("    ERREUR suppression {0} : {1}" -f $src.Name, $errDetail) "ERROR"
              Write-Host ("      ERREUR suppression : {0}" -f $src.Name) -ForegroundColor Red
              Write-ActionCsv "B" "SUPPRIMER" $src.Name $src.Id "Global" 0 $target.Name $target.Id "" "" "ERREUR" $errDetail
            }
          }

          Start-Sleep -Milliseconds 200
        }
      }

      # Resume Phase B
      Write-Host ""
      Write-Host "  Phase B terminee." -ForegroundColor Cyan
      Write-Host "  Pour retrouver les issues marquees dans Jira :" -ForegroundColor Yellow
      Write-Host '    JQL : description ~ "[FUSION]"' -ForegroundColor Yellow
      Write-Host "  Migrez manuellement par projet, puis relancez la Phase B" -ForegroundColor Yellow
      Write-Host "  pour supprimer les types doublons une fois vides." -ForegroundColor Yellow
      Write-Host ""
      Log "  Phase B terminee."
    } else {
      Log "  Phase B annulee par l'utilisateur."
      Write-Host "  Phase B annulee." -ForegroundColor Yellow
    }
  }
}