<#
Sheet2-SaisiTempsTempo.ps1 v2
Feuille 2 : Synthèse de saisie par équipe Tempo.
Utilise Get-ExpectedHoursFromWorkload (Common.ps1) pour le prorata
basé sur le Workload Scheme de chaque personne.
Détecte la saisie excessive (> heures attendues, sans tolérance).
#>

function Build-SaisiTempsTempo {
    param(
        [Parameter(Mandatory)]$TeamsData,
        [Parameter(Mandatory)][hashtable]$TimeLoggedMap,
        [Parameter(Mandatory)][double]$HeuresAttendue,
        [Parameter(Mandatory)][hashtable]$AssetByAccountId,
        [Parameter(Mandatory)][datetime]$PeriodeFrom,
        [Parameter(Mandatory)][datetime]$PeriodeTo,
        [Parameter(Mandatory)]$JoursFeries,
        [Parameter(Mandatory)][hashtable]$UserWorkloadDays
    )

    $headers = @(
        "Nom d'Équipe", "Programme", "Nombre de Membres", "Membres (ID)",
        "Aucune saisie", "Saisie incomplète", "Saisie complète",
        "Saisie excessive", "Saisie Incorrecte"
    )

    $rows = New-Object System.Collections.Generic.List[object]

    foreach ($t in $TeamsData) {
        $memberIdsStr = [string]$t."Membres ID"
        $members = @()
        if ($memberIdsStr) { $members = $memberIdsStr -split ",\s*" }

        $memberDates = $t.MemberDates

        $aucuneSaisie = 0
        $incomplete   = 0
        $complete     = 0
        $excessive    = 0
        $horsperiode  = 0

        foreach ($m in $members) {
            $heuresPerso = Get-ExpectedHoursFromWorkload `
                -AccountId $m `
                -UserWorkloadDays $UserWorkloadDays `
                -AssetByAccountId $AssetByAccountId `
                -MemberDates $memberDates `
                -PeriodeFrom $PeriodeFrom `
                -PeriodeTo $PeriodeTo `
                -JoursFeries $JoursFeries

            if ($heuresPerso -eq -1) {
                $horsperiode++
                Write-Log "F2 prorata: $m hors période dans équipe $($t.Nom)" "DEBUG"
                continue
            }

            $hours = 0.0
            if ($TimeLoggedMap.ContainsKey($m)) { $hours = [double]$TimeLoggedMap[$m] }

            if ($hours -eq 0) {
                if ($heuresPerso -eq 0) {
                    $horsperiode++
                } else {
                    $aucuneSaisie++
                }
            }
            elseif ($hours -gt $heuresPerso) {
                $excessive++
                Write-Log ("F2 excessive: {0} dans {1} — {2:N1}h saisies / {3:N1}h attendues" -f `
                    $m, $t.Nom, $hours, $heuresPerso) "WARN"
            }
            elseif ($hours -lt $heuresPerso) {
                $incomplete++
            }
            else {
                $complete++
            }
        }

        $membresActifs = $members.Count - $horsperiode

        $rows.Add([pscustomobject]@{
            "Nom d'Équipe"      = $t.Nom
            "Programme"          = (Transform-Programme ([string]$t.Programme))
            "Nombre de Membres" = [string]$membresActifs
            "Membres (ID)"      = $memberIdsStr
            "Aucune saisie"     = $aucuneSaisie
            "Saisie incomplète" = $incomplete
            "Saisie complète"   = $complete
            "Saisie excessive"  = $excessive
            "Saisie Incorrecte" = ($aucuneSaisie + $incomplete + $excessive)
        }) | Out-Null

        if ($horsperiode -gt 0) {
            Write-Log "F2: équipe $($t.Nom) — $horsperiode membre(s) hors période exclus" "DEBUG"
        }
    }

    return @{ Headers = $headers; Rows = $rows }
}
