<#
  Publier-Disponibilites.ps1 — TN FORMATION
  Lit l'onglet « Interventions » de « Calendrier interventions.xlsx » et génère :
    - disponibilites.js  : données de la page web (index.html)
    - disponibilites.ics : flux d'agenda auquel les clients peuvent s'abonner
  Ne publie AUCUN nom de client, cours ou lieu : seulement des jours libres ou pris.

  Règles (les mêmes que l'onglet « Disponibilités » du classeur) :
    - lundi à vendredi uniquement, jours fériés exclus
    - statut « Confirmée » ou « Prévisionnelle » -> jour pris
    - statut « Annulée »                         -> jour libre
    - jours ajoutés dans jours-bloques.txt (congés, perso…) -> réservé

  Utilisation :
    .\Publier-Disponibilites.ps1            # génère les fichiers depuis le classeur local
    .\Publier-Disponibilites.ps1 -Publier   # génère puis met en ligne (git push) si quelque chose a changé
    Dans le cloud (GitHub Actions), -Source reçoit le lien de partage OneDrive du classeur.

  Période publiée : d'aujourd'hui au 31/12/2027 (paramètre -Jusqua). Une fois cette date
  à moins de 6 mois, la période s'allonge automatiquement jusqu'à la fin de l'année suivante.
#>
[CmdletBinding()]
param(
  [string]$Source = $(if ($env:TNF_SOURCE) { $env:TNF_SOURCE } else { Join-Path "$env:USERPROFILE" 'OneDrive\Desktop\Calendrier interventions.xlsx' }),
  [string]$Jusqua = '2027-12-31',
  [switch]$Publier
)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$site = Split-Path $PSScriptRoot -Parent
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Heure de Paris, que le script tourne sur ton PC ou sur un serveur en UTC
$tz = $null
foreach ($id in 'Europe/Paris', 'Romance Standard Time') { try { $tz = [TimeZoneInfo]::FindSystemTimeZoneById($id); break } catch {} }
$maintenant = if ($tz) { [TimeZoneInfo]::ConvertTime([DateTimeOffset]::UtcNow, $tz) } else { [DateTimeOffset]::Now }

# --- 1. Lecture du classeur (copie temporaire : fonctionne même si Excel l'a ouvert) ---
$tmp = Join-Path ([IO.Path]::GetTempPath()) ("tnf-" + [guid]::NewGuid().ToString('N') + '.xlsx')
if ($Source -match '^https?://') {
  # Lien de partage OneDrive. Plusieurs méthodes, car Microsoft a changé le format des liens.
  # Le journal n'affiche JAMAIS d'adresse : seulement le numéro de méthode, le code HTTP et le type reçu.
  $enc = { param($s) 'u!' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($s)).TrimEnd('=').Replace('/', '_').Replace('+', '-') }
  $ajout = { param($u, $q) $u + $(if ($u.Contains('?')) { '&' } else { '?' }) + $q }
  $essais = [System.Collections.Generic.List[string]]::new()
  $essais.Add("https://api.onedrive.com/v1.0/shares/$(& $enc $Source)/root/content")
  $essais.Add("https://api.onedrive.com/v1.0/shares/$(& $enc $Source)/driveItem/content")
  $essais.Add((& $ajout ($Source -replace '\?.*$', '') 'download=1'))
  $essais.Add((& $ajout $Source 'download=1'))
  # Lien court 1drv.ms : on récupère l'adresse de destination pour en déduire un lien de téléchargement
  try {
    $req = [Net.HttpWebRequest]::Create($Source); $req.AllowAutoRedirect = $false; $req.Method = 'GET'
    $rep = $req.GetResponse(); $loc = $rep.Headers['Location']; $rep.Close()
  } catch { $loc = $null; if ($_.Exception.InnerException.Response) { $loc = $_.Exception.InnerException.Response.Headers['Location'] } }
  if ($loc) {
    if ($loc -match '[?&]redeem=([^&]+)') { $essais.Add("https://api.onedrive.com/v1.0/shares/u!$($Matches[1])/root/content") }
    if ($loc -match '[?&]resid=([^&]+)' ) { $resid = $Matches[1]
      $auth = if ($loc -match '[?&]authkey=([^&]+)') { "&authkey=$($Matches[1])" } else { '' }
      $essais.Add("https://onedrive.live.com/download?resid=$resid$auth") }
    $essais.Add((& $ajout $loc 'download=1'))
  }
  $ok = $false; $n = 0
  foreach ($u in $essais) {
    $n++
    if ($env:GITHUB_ACTIONS) { Write-Host "::add-mask::$u" }   # masque toute adresse dérivée du lien
    try {
      $r = Invoke-WebRequest -Uri $u -OutFile $tmp -UseBasicParsing -MaximumRedirection 10 -PassThru
      $sig = [IO.File]::ReadAllBytes($tmp)[0..1]
      $estXlsx = ($sig[0] -eq 0x50 -and $sig[1] -eq 0x4B)    # « PK » = fichier xlsx
      Write-Host ("Méthode {0} : HTTP {1}, type {2}, xlsx={3}" -f $n, [int]$r.StatusCode, ($r.Headers['Content-Type'] -join ''), $estXlsx)
      if ($estXlsx) { $ok = $true; break }
    } catch {
      $code = try { [int]$_.Exception.Response.StatusCode } catch { 0 }
      Write-Host ("Méthode {0} : échec (HTTP {1})" -f $n, $code)
    }
  }
  if (-not $ok) { throw "Impossible de télécharger le classeur depuis le lien OneDrive (lien expiré ou non partagé ?)." }
} else {
  Copy-Item -LiteralPath $Source -Destination $tmp
}
$zip = [System.IO.Compression.ZipFile]::OpenRead($tmp)
function Read-ZipXml([string]$name) {
  $e = $zip.Entries | Where-Object FullName -eq $name
  if (-not $e) { return $null }
  $r = New-Object System.IO.StreamReader($e.Open(), [Text.Encoding]::UTF8)
  try { [xml]$r.ReadToEnd() } finally { $r.Close() }
}
try {
  $wb   = Read-ZipXml 'xl/workbook.xml'
  $rels = Read-ZipXml 'xl/_rels/workbook.xml.rels'
  $sst  = Read-ZipXml 'xl/sharedStrings.xml'
  $strings = @(); if ($sst) { $strings = @($sst.sst.si | ForEach-Object { $_.InnerText }) }
  function Read-Sheet([string]$nom) {
    $sheet = $wb.workbook.sheets.sheet | Where-Object name -eq $nom
    if (-not $sheet) { return $null }
    $rid = $sheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $target = ($rels.Relationships.Relationship | Where-Object Id -eq $rid).Target -replace '^/xl/', ''
    Read-ZipXml ("xl/" + $target)
  }
  $ws = Read-Sheet 'Interventions'
  if (-not $ws) { throw "Onglet « Interventions » introuvable dans le classeur" }
  $wsIndispo = Read-Sheet 'Indisponibilités'   # facultatif : A = du, B = au (vide = 1 jour)
} finally { $zip.Dispose(); Remove-Item $tmp -Force }

function Get-CellValue($c) {
  if ($null -eq $c) { return $null }
  switch ($c.t) {
    's'         { return $strings[[int]$c.v] }
    'inlineStr' { return $c.is.InnerText }
    default     { return $c.v }
  }
}

# --- 2. Jours occupés d'après les interventions ---
$etat = @{}   # 'yyyy-MM-dd' -> 'reserve'
foreach ($row in $ws.worksheet.sheetData.row) {
  if ([int]$row.r -lt 2) { continue }
  $cA = $row.c | Where-Object { $_.r -eq "A$($row.r)" }
  $cL = $row.c | Where-Object { $_.r -eq "L$($row.r)" }
  $v = Get-CellValue $cA
  if (-not $v -or $v -notmatch '^\d+(\.\d+)?$') { continue }
  $jour = [DateTime]::FromOADate([double]$v).ToString('yyyy-MM-dd')
  $statut = "$(Get-CellValue $cL)".Trim()
  if ($statut -match '^Annul') { continue }
  $etat[$jour] = 'reserve'   # Confirmée, Prévisionnelle ou vide : on bloque
}

# --- 3. Jours bloqués à la main (congés, perso…) ---
# a) onglet facultatif « Indisponibilités » du classeur (fonctionne aussi dans le cloud)
if ($wsIndispo) {
  foreach ($row in $wsIndispo.worksheet.sheetData.row) {
    if ([int]$row.r -lt 2) { continue }
    $vA = Get-CellValue ($row.c | Where-Object { $_.r -eq "A$($row.r)" })
    $vB = Get-CellValue ($row.c | Where-Object { $_.r -eq "B$($row.r)" })
    if (-not $vA -or $vA -notmatch '^\d+(\.\d+)?$') { continue }
    $d1 = [DateTime]::FromOADate([double]$vA).Date
    $d2 = if ($vB -and $vB -match '^\d+(\.\d+)?$') { [DateTime]::FromOADate([double]$vB).Date } else { $d1 }
    for ($d = $d1; $d -le $d2; $d = $d.AddDays(1)) { $etat[$d.ToString('yyyy-MM-dd')] = 'reserve' }
  }
}
# b) fichier jours-bloques.txt à côté du script (utilisation sur ton PC uniquement)
$bloques = Join-Path $PSScriptRoot 'jours-bloques.txt'
if (Test-Path $bloques) {
  foreach ($l in Get-Content $bloques -Encoding UTF8) {
    $l = ($l -replace '#.*$', '').Trim(); if (-not $l) { continue }
    $p = $l -split '\s+'
    $d1 = [DateTime]::ParseExact($p[0], 'yyyy-MM-dd', $null)
    $d2 = if ($p.Count -gt 1) { [DateTime]::ParseExact($p[1], 'yyyy-MM-dd', $null) } else { $d1 }
    for ($d = $d1; $d -le $d2; $d = $d.AddDays(1)) { $etat[$d.ToString('yyyy-MM-dd')] = 'reserve' }
  }
}

# --- 4. Jours fériés (France métropolitaine) ---
function Get-Paques([int]$y) {
  $a=$y%19; $b=[math]::Floor($y/100); $c=$y%100; $d=[math]::Floor($b/4); $e=$b%4
  $f=[math]::Floor(($b+8)/25); $g=[math]::Floor(($b-$f+1)/3); $h=(19*$a+$b-$d-$g+15)%30
  $i=[math]::Floor($c/4); $k=$c%4; $l=(32+2*$e+2*$i-$h-$k)%7; $m=[math]::Floor(($a+11*$h+22*$l)/451)
  $mo=[math]::Floor(($h+$l-7*$m+114)/31); $dy=(($h+$l-7*$m+114)%31)+1
  Get-Date -Year $y -Month $mo -Day $dy -Hour 0 -Minute 0 -Second 0 -Millisecond 0
}
$debut = $maintenant.Date
$fin = [DateTime]::ParseExact($Jusqua, 'yyyy-MM-dd', $null)
if ($fin -lt $debut.AddMonths(6)) { $fin = Get-Date -Year ($debut.Year + 1) -Month 12 -Day 31 -Hour 0 -Minute 0 -Second 0 -Millisecond 0 }
$feries = @{}
foreach ($y in $debut.Year..$fin.Year) {
  $pq = Get-Paques $y
  $liste = @(
    @("$y-01-01", 'Jour de l''An'), @("$y-05-01", 'Fête du Travail'), @("$y-05-08", 'Victoire 1945'),
    @("$y-07-14", 'Fête nationale'), @("$y-08-15", 'Assomption'), @("$y-11-01", 'Toussaint'),
    @("$y-11-11", 'Armistice'), @("$y-12-25", 'Noël'),
    @($pq.AddDays(1).ToString('yyyy-MM-dd'), 'Lundi de Pâques'),
    @($pq.AddDays(39).ToString('yyyy-MM-dd'), 'Ascension'),
    @($pq.AddDays(50).ToString('yyyy-MM-dd'), 'Lundi de Pentecôte'))
  foreach ($j in $liste) { $feries[$j[0]] = $j[1] }   # deux fêtes le même jour : pas d'erreur
}

# --- 5. Calcul jour par jour sur la période publiée ---
$jours = [ordered]@{}
for ($d = $debut; $d -le $fin; $d = $d.AddDays(1)) {
  $k = $d.ToString('yyyy-MM-dd')
  if ($d.DayOfWeek -in 'Saturday','Sunday') { continue }        # week-end : déduit par la page
  if ($feries.ContainsKey($k)) { $jours[$k] = 'ferie'; continue }
  $jours[$k] = if ($etat.ContainsKey($k)) { $etat[$k] } else { 'dispo' }
}
$maj = $maintenant
$data = [ordered]@{
  maj    = $maj.ToString('yyyy-MM-ddTHH:mm:sszzz')
  debut  = $debut.ToString('yyyy-MM-dd')
  fin    = $fin.ToString('yyyy-MM-dd')
  feries = $feries
  jours  = $jours
}
$utf8 = New-Object System.Text.UTF8Encoding($false)
# N'écrit le fichier que si son contenu change (en ignorant l'horodatage), pour éviter des publications inutiles
function Write-SiChange([string]$chemin, [string]$contenu, [string]$horodatage) {
  if (Test-Path $chemin) {
    $avant = [IO.File]::ReadAllText($chemin, $utf8)
    if (($avant -replace $horodatage, '') -eq ($contenu -replace $horodatage, '')) { return $false }
  }
  [IO.File]::WriteAllText($chemin, $contenu, $utf8); return $true
}
$js = "window.DISPO = " + ($data | ConvertTo-Json -Depth 4 -Compress) + ";`n"
$changeJs = Write-SiChange (Join-Path $site 'disponibilites.js') $js '"maj":"[^"]*"'

# --- 6. Flux iCalendar : un événement par période de jours libres consécutifs ---
$lines = @('BEGIN:VCALENDAR','VERSION:2.0','PRODID:-//TN FORMATION//Disponibilites//FR','CALSCALE:GREGORIAN','METHOD:PUBLISH',
           'X-WR-CALNAME:TN FORMATION - Disponibilités','X-WR-TIMEZONE:Europe/Paris','REFRESH-INTERVAL;VALUE=DURATION:PT1H','X-PUBLISHED-TTL:PT1H')
$stamp = $maj.ToUniversalTime().ToString("yyyyMMdd'T'HHmmss'Z'")
$libres = @($jours.Keys | Where-Object { $jours[$_] -eq 'dispo' } | ForEach-Object { [DateTime]::ParseExact($_, 'yyyy-MM-dd', $null) })
$i = 0
while ($i -lt $libres.Count) {
  $s = $libres[$i]; $e = $s
  # on fusionne les jours libres qui se suivent (en sautant le week-end)
  while ($i + 1 -lt $libres.Count) {
    $n = $libres[$i + 1]; $gap = $e.AddDays(1)
    while ($gap.DayOfWeek -in 'Saturday','Sunday') { $gap = $gap.AddDays(1) }
    if ($n -ne $gap) { break }
    $e = $n; $i++
  }
  $nb = @($libres | Where-Object { $_ -ge $s -and $_ -le $e }).Count
  $lines += 'BEGIN:VEVENT'
  $lines += "UID:dispo-$($s.ToString('yyyyMMdd'))-$($e.ToString('yyyyMMdd'))@tnformation"
  $lines += "DTSTAMP:$stamp"
  $lines += "DTSTART;VALUE=DATE:$($s.ToString('yyyyMMdd'))"
  $lines += "DTEND;VALUE=DATE:$($e.AddDays(1).ToString('yyyyMMdd'))"
  $lines += "SUMMARY:Tanguy NOTO disponible ($nb j)"
  $lines += 'DESCRIPTION:Disponible pour une intervention (présentiel IDF ou distanciel). Réserver : tanguy@tnformation.net'
  $lines += 'TRANSP:TRANSPARENT'
  $lines += 'END:VEVENT'
  $i++
}
$lines += 'END:VCALENDAR'
# RFC 5545 : lignes de 75 octets max, suite précédée d'un espace
$folded = foreach ($l in $lines) {
  $out = ''; $cur = ''
  foreach ($ch in $l.ToCharArray()) {
    if ($utf8.GetByteCount($cur + $ch) -gt 74) { $out += $cur + "`r`n "; $cur = '' }
    $cur += $ch
  }
  $out + $cur
}
$changeIcs = Write-SiChange (Join-Path $site 'disponibilites.ics') (($folded -join "`r`n") + "`r`n") 'DTSTAMP:\d{8}T\d{6}Z'
if (-not ($changeJs -or $changeIcs)) { Write-Host 'Disponibilités inchangées depuis la dernière publication.' }

$nbDispo = @($jours.Values | Where-Object { $_ -eq 'dispo' }).Count
$nbRes   = @($jours.Values | Where-Object { $_ -eq 'reserve' }).Count
Write-Host ("Période {0} -> {1} : {2} jours libres, {3} jours pris." -f $data.debut, $data.fin, $nbDispo, $nbRes)

# --- 7. Mise en ligne (GitHub Pages) ---
if ($Publier) {
  Push-Location $site
  # git écrit ses messages sur la sortie d'erreur : on ne doit pas les traiter comme des erreurs fatales
  $ErrorActionPreference = 'Continue'
  try {
    git add disponibilites.js disponibilites.ics 2>&1 | Out-Null
    git diff --cached --quiet
    if ($LASTEXITCODE -ne 0) {
      git commit -q -m ("Disponibilités mises à jour le " + $maj.ToString('dd/MM/yyyy HH:mm')) 2>&1 | Out-Null
      # --autostash : fonctionne même si un autre fichier du site est en cours de modification
      git pull -q --rebase --autostash 2>&1 | Out-Null
      git push -q 2>&1 | Out-Null
      if ($LASTEXITCODE -eq 0) { Write-Host 'Disponibilités publiées en ligne.' }
      else { Write-Host 'Échec de la mise en ligne (pas de connexion internet ?) : nouvel essai au prochain passage.' }
    } else {
      # rien de neuf, mais on renvoie un éventuel envoi resté en attente (PC hors ligne la fois précédente)
      $attente = git rev-list --count '@{u}..HEAD' 2>$null
      if ($attente -and [int]$attente -gt 0) { git push -q 2>&1 | Out-Null; Write-Host 'Envoi en attente publié.' }
      else { Write-Host 'Aucun changement : rien à publier.' }
    }
  } finally { Pop-Location }
}
