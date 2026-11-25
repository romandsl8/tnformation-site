# Lancé par la tâche planifiée « TN FORMATION - Disponibilités » (toutes les 15 min et à l'ouverture de session).
# Publie les disponibilités et garde un petit journal des 200 derniers passages dans journal.txt (non publié).
$journal = Join-Path $PSScriptRoot 'journal.txt'
$sortie = try { & (Join-Path $PSScriptRoot 'Publier-Disponibilites.ps1') -Publier *>&1 | Out-String } catch { "ERREUR : $($_.Exception.Message)" }
$ligne = "{0}  {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm'), (($sortie -split "`r?`n" | Where-Object { $_.Trim() }) -join ' | ')
$anciennes = @(if (Test-Path $journal) { Get-Content $journal -Encoding UTF8 | Select-Object -Last 199 })
Set-Content -Path $journal -Value ($anciennes + $ligne) -Encoding UTF8
