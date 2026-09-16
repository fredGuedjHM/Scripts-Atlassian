<#
Sheet5-WorklogsSansDate.ps1
Feuille 5 : Worklogs sans Start Date.
Identifie les worklogs Tempo dont startDate est vide ou null.
1 ligne par worklog × par équipe de l'auteur.
#>

function Build-WorklogsSansDate {
    param(
        [Parameter(Mandatory)]$Worklogs,
        [Parameter(Mandatory)][hashtable]$IssuesById,
        [Parameter(Mandatory)][hashtable]$UserTeamsMap,
        [Parameter(Mandatory)][hashtable]$JiraHeaders
    )

    $headers = @(
        "Start Date", "Author ID", "Nom d'Équipe",
        "Time Spent (Hours)", "Issue (Key - Summary)",
        "Issue Type", "Account", "Parent Issue"
    )

    $rows = New-Object System.Collections.Generic.List[object]

    foreach ($wl in $Worklogs) {
        $sd = [string]$wl.startDate
        if (-not [string]::IsNullOrWhiteSpace($sd)) { continue }

        $authorId = ""
        try { $authorId = [string]$wl.author.accountId } catch {}
        if ([string]::IsNullOrWhiteSpace($authorId)) { continue }

        $timeH = [math]::Round(([double]$wl.timeSpentSeconds) / 3600.0, 2)

        # Issue
        $issueKeySummary = ""; $issueType = ""; $account = ""; $parentIssue = ""
        if ($wl.issue -and $wl.issue.id) {
            $issueId = [string]$wl.issue.id
            $issue = $null
            if ($IssuesById.ContainsKey($issueId)) { $issue = $IssuesById[$issueId] }
            if ($issue) {
                $key = [string]$issue.key
                $summary = ""
                if ($issue.fields -and $issue.fields.summary) { $summary = [string]$issue.fields.summary }
                $issueKeySummary = "$key - $summary"

                if ($issue.fields -and $issue.fields.issuetype) {
                    $issueType = [string]$issue.fields.issuetype.name
                }

                # Account (customfield_10032)
                if ($issue.fields -and $issue.fields.customfield_10032) {
                    $acctField = $issue.fields.customfield_10032
                    if ($acctField.value) { $account = [string]$acctField.value }
                    elseif ($acctField.name) { $account = [string]$acctField.name }
                    else { $account = [string]$acctField }
                }

                # Remontée parent (Initiative) — CORRIGÉ : était Get-ParentInitiative
                $parentIssue = Get-InitiativeParentString -Issue $issue -IssuesById $IssuesById -JiraHeaders $JiraHeaders
            }
        }

        # Équipes
        $teams = @("Non attribué")
        if ($UserTeamsMap.ContainsKey($authorId) -and $UserTeamsMap[$authorId].Count -gt 0) {
            $teams = $UserTeamsMap[$authorId]
        }

        foreach ($teamName in $teams) {
            $rows.Add([pscustomobject]@{
                "Start Date"            = ""
                "Author ID"             = $authorId
                "Nom d'Équipe"          = $teamName
                "Time Spent (Hours)"    = $timeH
                "Issue (Key - Summary)" = $issueKeySummary
                "Issue Type"            = $issueType
                "Account"               = $account
                "Parent Issue"          = $parentIssue
            }) | Out-Null
        }
    }

    return @{ Headers = $headers; Rows = $rows }
}