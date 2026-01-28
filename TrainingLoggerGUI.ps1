#requires -Version 7.0
using namespace System
using namespace System.IO
using namespace System.Drawing
using namespace System.Windows.Forms
using namespace System.Windows.Forms.DataVisualization.Charting

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms.DataVisualization

[Application]::EnableVisualStyles()

# ----------------------------
# Storage
# ----------------------------
$scriptDir = Split-Path -Parent $PSCommandPath
$BaseDir   = Join-Path $env:USERPROFILE "TrainingLoggerPro"
$LogPath   = Join-Path $BaseDir "traininglog.csv"
$ReportDir = Join-Path $BaseDir "Reports"
$BackupDir = Join-Path $BaseDir "Backups"
$SettingsPath = Join-Path $BaseDir "settings.json"

$LogoIco = Join-Path $scriptDir "logo.ico"   # optional: put logo.ico next to this script
$LogoPng = Join-Path $scriptDir "logo.png"   # optional: put logo.png next to this script

$Sports = @("Running","Gym","BJJ","Kickboxing","Cycling","Swimming","Other")

function Ensure-Storage {
    foreach ($p in @($BaseDir,$ReportDir,$BackupDir)) {
        if (-not (Test-Path $p)) { New-Item -ItemType Directory -Path $p | Out-Null }
    }
    if (-not (Test-Path $LogPath)) {
        "Id,Date,Sport,DurationMin,Calories,DistanceKm,RPE,AvgHR,Note" | Out-File -Encoding utf8 -FilePath $LogPath
    }
}

# ------------------------------
# Storage / paths (always set)
# ------------------------------
$script:AppName = "TrainingLoggerPro"
$script:DataDir = Join-Path $env:APPDATA $script:AppName
if (-not (Test-Path $script:DataDir)) { New-Item -ItemType Directory -Path $script:DataDir -Force | Out-Null }
$script:EntriesFile  = Join-Path $script:DataDir "entries.json"
$script:PrsFile      = Join-Path $script:DataDir "prs.json"
$script:SettingsFile = Join-Path $script:DataDir "settings.json"

# Simple in-memory cache
$script:EntriesCache = $null
function Get-EntriesCached {
    param([switch]$Force)
    if ($Force -or -not $script:EntriesCache) { $script:EntriesCache = Load-Entries }
    return $script:EntriesCache
}

# ------------------------------
# "AI" feedback (fast + local)
# ------------------------------
function New-OllamaStyleFeedback {
    param(
        [Parameter(Mandatory)]$Summary,
        [Parameter()]$Entries
    )

    $rng = New-Object System.Random

    $wins = @(
        "Consistency looks solid — keep the streak going.",
        "Good balance between endurance and strength this week.",
        "Nice volume — you handled it well overall.",
        "Good discipline on the basics: you showed up and got it done."
    )

    $risks = @(
        "Your average effort looks high — add 1–2 genuinely easy days next week.",
        "Watch recovery: prioritise sleep + hydration and keep at least one day very light.",
        "If you feel niggles, cut intensity first and keep the easy volume."
    )

    $focus = @(
        "1 long easy session + 1 quality session + 2 strength sessions (full body).",
        "2 easy aerobic sessions + 1 controlled tempo/interval + 2 strength sessions.",
        "Keep it simple: 1 long easy, 2 short easy, 1 quality, 2 strength."
    )

    $micro = @(
        "Warm-up: 8–12 min easy + 3 strides. Cool-down: 5–10 min easy.",
        "For strength, leave 1–2 reps in reserve on most sets (don’t grind).",
        "Keep most running in Z2; save the hard work for one key session."
    )

    $period = if ($Summary.Period) { $Summary.Period } else { (Get-Date).ToString("yyyy-MM-dd") }
    $s = @()
    $s += "Ollama Coach Feedback"
    $s += ""
    $s += "Period: $period"
    $s += ""
    $s += "Key wins:"
    $s += "- " + $wins[$rng.Next($wins.Count)]
    $s += "- " + $wins[$rng.Next($wins.Count)]
    $s += ""
    $s += "What to improve next week:"
    $s += "- " + $risks[$rng.Next($risks.Count)]
    $s += "- " + $micro[$rng.Next($micro.Count)]
    $s += ""
    $s += "Suggested focus:"
    $s += "- " + $focus[$rng.Next($focus.Count)]
    $s += ""
    $s += "Next step:"
    $s += "- Pick 2–3 priorities, keep everything else easy, and reassess in 7 days."

    return ($s -join "`r`n")
}
function Backup-Log {
    Ensure-Storage
    $stamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
    Copy-Item $LogPath (Join-Path $BackupDir "traininglog_backup_$stamp.csv") -Force
}

function Load-Settings {
    Ensure-Storage
    if (Test-Path $SettingsPath) {
        try { return (Get-Content $SettingsPath -Raw | ConvertFrom-Json) } catch {}
    }
    return [pscustomobject]@{
        DarkMode = $false
        DefaultRPEForLoad = 5
    }
}
function Save-Settings($s) {
    $s | ConvertTo-Json -Depth 6 | Out-File -Encoding utf8 -FilePath $SettingsPath
}
$script:Settings = Load-Settings

# ----------------------------
# Data
# ----------------------------
function Load-Entries {
    Ensure-Storage
    Import-Csv -Path $LogPath | ForEach-Object {
        $d = $null
        try { $d = [datetime]::Parse($_.Date) } catch { $d = Get-Date }
        [pscustomobject]@{
            Id          = $_.Id
            Date        = $d.ToString("yyyy-MM-dd")
            DateObj     = $d.Date
            Sport       = $_.Sport
            DurationMin = [int]$_.DurationMin
            Calories    = if ($_.Calories -and $_.Calories -match '^\d+$') { [int]$_.Calories } else { $null }
            DistanceKm  = if ($_.DistanceKm) { [double]($_.DistanceKm -replace ",",".") } else { $null }
            RPE         = if ($_.RPE -and $_.RPE -match '^\d+$') { [int]$_.RPE } else { $null }
            AvgHR       = if ($_.AvgHR -and $_.AvgHR -match '^\d+$') { [int]$_.AvgHR } else { $null }
            Note        = $_.Note
        }
    } | Sort-Object DateObj
}
function Append-Entry($e) {
    $note = ($e.Note ?? "") -replace '"','""'
    $dist = if ($null -ne $e.DistanceKm) { $e.DistanceKm } else { "" }
    $cal  = if ($null -ne $e.Calories)   { $e.Calories } else { "" }
    $rpe  = if ($null -ne $e.RPE)        { $e.RPE } else { "" }
    $hr   = if ($null -ne $e.AvgHR)      { $e.AvgHR } else { "" }

    $line = "{0},{1},{2},{3},{4},{5},{6},{7},""{8}""" -f `
        $e.Id, $e.Date, $e.Sport, $e.DurationMin, $cal, $dist, $rpe, $hr, $note
    Add-Content -Path $LogPath -Value $line -Encoding utf8
}
function Rewrite-All($entries) {
    Backup-Log
    "Id,Date,Sport,DurationMin,Calories,DistanceKm,RPE,AvgHR,Note" | Out-File -Encoding utf8 -FilePath $LogPath
    foreach ($e in $entries) { Append-Entry $e }
}

function Get-LoadForEntry($e) {
    $rpe = $e.RPE
    if (-not $rpe -or $rpe -le 0) { $rpe = [int]$script:Settings.DefaultRPEForLoad }
    return [int]($e.DurationMin * $rpe)
}

function Get-DailyLoad($entries) {
    $entries | Group-Object { $_.DateObj } | ForEach-Object {
        [pscustomobject]@{
            Date     = $_.Group[0].DateObj
            Load     = ($_.Group | ForEach-Object { Get-LoadForEntry $_ } | Measure-Object -Sum).Sum
            Minutes  = ($_.Group | Measure-Object DurationMin -Sum).Sum
            Sessions = $_.Count
        }
    } | Sort-Object Date
}

function Get-WeekStart([datetime]$d) {
    $day = [int]$d.DayOfWeek
    if ($day -eq 0) { $day = 7 } # Sunday -> 7
    return $d.Date.AddDays(1 - $day) # Monday
}

# ----------------------------
# PRs (manual + computed)
# ----------------------------
$script:PrsFile = Join-Path $BaseDir "prs.json"

function ConvertTo-DataTable {
    param(
        [Parameter(Mandatory=$true)][object[]]$Items
    )
    $dt = New-Object System.Data.DataTable
    if (-not $Items -or $Items.Count -eq 0) { return $dt }

    # Build columns from first object
    $first = $Items[0]
    $props = @()
    if ($first -is [hashtable]) {
        $props = $first.Keys
    } else {
        $props = ($first | Get-Member -MemberType NoteProperty,Property | Select-Object -ExpandProperty Name)
    }
    foreach ($p in $props) { [void]$dt.Columns.Add($p) }

    foreach ($it in $Items) {
        $row = $dt.NewRow()
        foreach ($p in $props) {
            try {
                $val = if ($it -is [hashtable]) { $it[$p] } else { $it.$p }
                $row[$p] = if ($null -eq $val) { "" } else { "$val" }
            } catch {
                $row[$p] = ""
            }
        }
        [void]$dt.Rows.Add($row)
    }
    return $dt
}

function Load-ManualPRs {
    Ensure-Storage
    if (-not (Test-Path $script:PrsFile)) { return @() }
    try {
        $j = Get-Content $script:PrsFile -Raw | ConvertFrom-Json
        if ($null -eq $j) { return @() }
        return @($j)
    } catch { return @() }
}

function Save-ManualPRs([object[]]$prs) {
    Ensure-Storage
    $prs | ConvertTo-Json -Depth 6 | Out-File -Encoding utf8 -FilePath $script:PrsFile
}

function Get-ComputedPRs {
    param([object[]]$Entries)
    if (-not $Entries -or $Entries.Count -eq 0) { return @() }

    $prs = New-Object System.Collections.Generic.List[object]

    # General
    $longest = $Entries | Sort-Object DurationMin -Descending | Select-Object -First 1
    if ($longest) {
        $prs.Add([pscustomobject]@{ Type='Longest session'; Sport=$longest.Sport; Value="{0} min" -f $longest.DurationMin; Date=$longest.Date; Note=($longest.Note ?? '') })
    }
    $highest = $Entries | Sort-Object { Get-LoadForEntry $_ } -Descending | Select-Object -First 1
    if ($highest) {
        $prs.Add([pscustomobject]@{ Type='Highest load'; Sport=$highest.Sport; Value="{0} (min*RPE)" -f (Get-LoadForEntry $highest); Date=$highest.Date; Note=($highest.Note ?? '') })
    }

    # Running (best pace by distance)
    $runs = $Entries | Where-Object { $_.Sport -eq 'Running' -and $_.DurationMin -gt 0 -and $_.DistanceKm -gt 0 }
    foreach ($target in @(5,10,21.1)) {
        $cand = $runs | Where-Object { [double]$_.DistanceKm -ge ($target - 0.15) -and [double]$_.DistanceKm -le ($target + 0.15) }
        if ($cand) {
            $best = $cand | Sort-Object { $_.DurationMin / $_.DistanceKm } | Select-Object -First 1
            $pace = [timespan]::FromMinutes($best.DurationMin / $best.DistanceKm)
            $paceStr = "{0}:{1:00} /km" -f [int]$pace.Minutes, [int]$pace.Seconds
            $prs.Add([pscustomobject]@{ Type="Best ~${target}k"; Sport='Running'; Value="$($best.DurationMin) min @ $paceStr"; Date=$best.Date; Note=($best.Note ?? '') })
        }
    }
    $bestRunDist = $runs | Sort-Object DistanceKm -Descending | Select-Object -First 1
    if ($bestRunDist) {
        $prs.Add([pscustomobject]@{ Type='Longest run'; Sport='Running'; Value="{0} km" -f [math]::Round($bestRunDist.DistanceKm,2); Date=$bestRunDist.Date; Note=($bestRunDist.Note ?? '') })
    }

    return $prs
}

# ----------------------------
# PDF Export (Edge headless)
# ----------------------------
function Get-EdgePath {
    # Try PATH / App Paths / common installs
    try {
        $cmd = Get-Command msedge.exe -ErrorAction SilentlyContinue
        if ($cmd -and $cmd.Source -and (Test-Path $cmd.Source)) { return $cmd.Source }
    } catch {}

    $candidates = @(
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles(x86)\Microsoft\Edge\Application\msedge.exe",
        "$env:LocalAppData\Microsoft\Edge\Application\msedge.exe"
    ) | Where-Object { $_ -and (Test-Path $_) }

    if ($candidates.Count -gt 0) { return $candidates[0] }

    foreach ($k in @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe"
    )) {
        try {
            $p = (Get-ItemProperty -Path $k -ErrorAction Stop).'(default)'
            if ($p -and (Test-Path $p)) { return $p }
        } catch {}
    }

    return $null
}
function Export-WeeklyReportToPDF([string]$pdfPath) {
    Ensure-Storage
    $entries = Load-Entries
    $today = (Get-Date).Date
    $start = Get-WeekStart $today
    $end = $start.AddDays(7)

    $week = $entries | Where-Object { $_.DateObj -ge $start -and $_.DateObj -lt $end } | Sort-Object DateObj
    $totalSessions = $week.Count
    $totalMin = ($week | Measure-Object DurationMin -Sum).Sum
    if (-not $totalMin) { $totalMin = 0 }

    $totalLoad = ($week | ForEach-Object { Get-LoadForEntry $_ } | Measure-Object -Sum).Sum
    if (-not $totalLoad) { $totalLoad = 0 }

    $totalDist = ($week | Where-Object { $null -ne $_.DistanceKm } | Measure-Object DistanceKm -Sum).Sum
    if (-not $totalDist) { $totalDist = 0 }

    $bySport = $week | Group-Object Sport | ForEach-Object {
        [pscustomobject]@{
            Sport = $_.Name
            Sessions = $_.Count
            Minutes = ($_.Group | Measure-Object DurationMin -Sum).Sum
            Load = ($_.Group | ForEach-Object { Get-LoadForEntry $_ } | Measure-Object -Sum).Sum
        }
    } | Sort-Object Minutes -Descending

    $htmlPath = Join-Path $ReportDir ("WeeklyReport_{0}.html" -f $start.ToString("yyyyMMdd"))

    $rows = ($week | Sort-Object DateObj -Descending | ForEach-Object {
        $note = [System.Web.HttpUtility]::HtmlEncode(($_.Note ?? ""))
        $dist = if ($null -ne $_.DistanceKm) { [math]::Round($_.DistanceKm,2) } else { "" }
        $kcal = if ($null -ne $_.Calories) { $_.Calories } else { "" }
        $rpe  = if ($null -ne $_.RPE) { $_.RPE } else { "" }
        $hr   = if ($null -ne $_.AvgHR) { $_.AvgHR } else { "" }
        "<tr><td>$($_.Date)</td><td>$($_.Sport)</td><td style='text-align:right'>$($_.DurationMin)</td><td style='text-align:right'>$dist</td><td style='text-align:right'>$kcal</td><td style='text-align:right'>$rpe</td><td style='text-align:right'>$hr</td><td>$note</td></tr>"
    }) -join "`n"

    $bySportRows = ($bySport | ForEach-Object {
        "<tr><td>$($_.Sport)</td><td style='text-align:right'>$($_.Sessions)</td><td style='text-align:right'>$($_.Minutes)</td><td style='text-align:right'>$($_.Load)</td></tr>"
    }) -join "`n"

    $css = @"
<style>
body { font-family: Segoe UI, Arial, sans-serif; margin: 28px; color:#111; }
h1 { margin: 0 0 6px 0; }
.small { color:#666; margin: 0 0 18px 0; }
.card { border:1px solid #ddd; border-radius:14px; padding:14px 16px; margin-bottom:14px; }
.grid { display:grid; grid-template-columns: repeat(4, 1fr); gap:10px; }
.kpi { background:#f6f7fb; border-radius:12px; padding:10px 12px; border:1px solid #e7e8ef; }
.kpi b { display:block; font-size:18px; margin-top:4px; }
table { width:100%; border-collapse: collapse; }
th, td { border-bottom:1px solid #eee; padding:8px 6px; font-size: 12.5px; }
th { text-align:left; background:#fafafa; }
.footer { margin-top: 18px; color:#666; font-size: 12px; }
</style>
"@

    $html = @"
<html><head><meta charset="utf-8" />$css</head>
<body>
  <h1>Training Logger Pro (Joe Witton) – Weekly Report</h1>
  <p class="small">$($start.ToString("yyyy-MM-dd")) → $($end.AddDays(-1).ToString("yyyy-MM-dd"))</p>

  <div class="card grid">
    <div class="kpi">Sessions<b>$totalSessions</b></div>
    <div class="kpi">Minutes<b>$totalMin</b></div>
    <div class="kpi">Distance (km)<b>$([math]::Round($totalDist,2))</b></div>
    <div class="kpi">Training Load<b>$totalLoad</b></div>
  </div>

  <div class="card">
    <h3 style="margin:0 0 10px 0;">By sport</h3>
    <table>
      <thead><tr><th>Sport</th><th style="text-align:right">Sessions</th><th style="text-align:right">Minutes</th><th style="text-align:right">Load</th></tr></thead>
      <tbody>$bySportRows</tbody>
    </table>
  </div>

  <div class="card">
    <h3 style="margin:0 0 10px 0;">Sessions</h3>
    <table>
      <thead><tr><th>Date</th><th>Sport</th><th style="text-align:right">Min</th><th style="text-align:right">Km</th><th style="text-align:right">Kcal</th><th style="text-align:right">RPE</th><th style="text-align:right">Avg HR</th><th>Note</th></tr></thead>
      <tbody>$rows</tbody>
    </table>
  </div>

  <div class="footer">Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm")</div>
</body></html>
"@

    $html | Out-File -Encoding utf8 -FilePath $htmlPath

    $edge = Get-EdgePath
    if (-not $edge) { return $null }

    $args = @("--headless","--disable-gpu","--no-first-run","--print-to-pdf=""$pdfPath""",$htmlPath)
    $p = Start-Process -FilePath $edge -ArgumentList $args -PassThru -WindowStyle Hidden
    $p.WaitForExit()

    for ($i=0; $i -lt 40 -and -not (Test-Path $pdfPath); $i++) { Start-Sleep -Milliseconds 250 }
    if (-not (Test-Path $pdfPath)) { return $null }
    return $pdfPath
}

# ----------------------------
# UI helpers
# ----------------------------
$FontUI    = New-Object Font("Segoe UI", 10)
$FontTitle = New-Object Font("Segoe UI", 11, [FontStyle]::Bold)

function MakeRoundedButton([string]$text, [System.Drawing.Icon]$icon) {
    $b = New-Object Button
    $b.Text = "  " + $text
    $b.Font = New-Object Font("Segoe UI", 9.5)
    $b.Height = 40
    $b.Width = 150
    $b.FlatStyle = "Flat"
    $b.FlatAppearance.BorderSize = 0
    $b.BackColor = [Color]::FromArgb(245,245,245)
    $b.ForeColor = [Color]::FromArgb(40,40,40)
    $b.Margin = "8,6,8,6"
    $b.Padding = "12,0,12,0"
    $b.TextAlign = "MiddleCenter"
    $b.TextImageRelation = "ImageBeforeText"

    if ($icon -is [System.Drawing.Icon]) {
        $b.Image = $icon.ToBitmap()
        $b.ImageAlign = "MiddleLeft"
    }

    $b.Add_Resize({
        $radius = 16
        $rect = $this.ClientRectangle
        $path = New-Object Drawing2D.GraphicsPath
        $d = $radius * 2
        $path.AddArc($rect.X, $rect.Y, $d, $d, 180, 90)
        $path.AddArc($rect.Right - $d, $rect.Y, $d, $d, 270, 90)
        $path.AddArc($rect.Right - $d, $rect.Bottom - $d, $d, $d, 0, 90)
        $path.AddArc($rect.X, $rect.Bottom - $d, $d, $d, 90, 90)
        $path.CloseFigure()
        $this.Region = New-Object Region($path)
        $path.Dispose()
    })

    $b.Add_MouseEnter({ if ($this.Tag -ne "active") { $this.BackColor = [Color]::FromArgb(235,240,255) } })
    $b.Add_MouseLeave({ if ($this.Tag -ne "active") { $this.BackColor = [Color]::FromArgb(245,245,245) } })

    return $b
}

function ApplyTheme([Control]$root, [bool]$dark) {
    $bg = if ($dark) { [Color]::FromArgb(28,28,30) } else { [SystemColors]::Window }
    $panel = if ($dark) { [Color]::FromArgb(40,40,42) } else { [SystemColors]::Control }
    $fg = if ($dark) { [Color]::Gainsboro } else { [SystemColors]::ControlText }

    $queue = New-Object System.Collections.Generic.Queue[Control]
    $queue.Enqueue($root)

    while ($queue.Count -gt 0) {
        $c = $queue.Dequeue()
        try {
            if ($c -is [Form]) { $c.BackColor=$bg; $c.ForeColor=$fg }
            elseif ($c -is [Panel] -or $c -is [TableLayoutPanel] -or $c -is [FlowLayoutPanel] -or $c -is [GroupBox]) { $c.BackColor=$panel; $c.ForeColor=$fg }
            elseif ($c -is [TextBox]) { $c.BackColor=$bg; $c.ForeColor=$fg }
            elseif ($c -is [DataGridView]) { $c.BackgroundColor=$bg; $c.ForeColor=$fg }
            elseif ($c -is [Button]) {
                if ($c.Tag -eq "active") {
                    $c.BackColor = if ($dark) { [Color]::FromArgb(80,100,150) } else { [Color]::FromArgb(205,220,255) }
                } else {
                    $c.BackColor = if ($dark) { [Color]::FromArgb(55,55,58) } else { [Color]::FromArgb(245,245,245) }
                }
                $c.ForeColor = $fg
            } else { $c.ForeColor=$fg }
        } catch {}
        foreach ($child in $c.Controls) { $queue.Enqueue($child) }
    }
}


# ----------------------------
# Ollama AI helpers (stable, GUI-safe)
# ----------------------------
function Invoke-UI {
    param(
        [Parameter(Mandatory=$true)] [System.Windows.Forms.Control] $Control,
        [Parameter(Mandatory=$true)] [ScriptBlock] $Action
    )
    if ($null -eq $Control -or $Control.IsDisposed) { return }
    if ($Control.InvokeRequired) { $null = $Control.BeginInvoke($Action) } else { & $Action }
}

function Get-OllamaModels {
    param(
        [string]$BaseUrl = "http://127.0.0.1:11434",
        [int]$TimeoutSec = 10
    )
    $uri = "$BaseUrl/api/tags"
    $handler = [System.Net.Http.HttpClientHandler]::new()
    $client  = [System.Net.Http.HttpClient]::new($handler)
    $client.Timeout = [TimeSpan]::FromSeconds($TimeoutSec)
    try {
        $json = $client.GetStringAsync($uri).GetAwaiter().GetResult()
        $obj  = $json | ConvertFrom-Json
        $names = @()
        foreach ($m in ($obj.models | ForEach-Object { $_ })) {
            if ($m.name) { $names += [string]$m.name }
        }
        return $names
    } catch {
        return @()
    } finally {
        $client.Dispose()
        $handler.Dispose()
    }
}

function Invoke-OllamaGenerate_Sync {
    param(
        [Parameter(Mandatory=$true)] [string]$BaseUrl,
        [Parameter(Mandatory=$true)] [string]$Model,
        [Parameter(Mandatory=$true)] [string]$Prompt,
        [int]$TimeoutSec = 300,
        [int]$NumPredict = 220
    )
    if ([string]::IsNullOrWhiteSpace($Model)) { throw "Model is empty. Example: llama3.1:latest" }
    if ([string]::IsNullOrWhiteSpace($Prompt)) { throw "Prompt is empty." }

    $uri = "$BaseUrl/api/generate"

    $payload = @{
        model   = $Model
        prompt  = $Prompt
        stream  = $false
        options = @{
            num_predict = $NumPredict
            temperature = 0.2
            top_p       = 0.9
        }
    } | ConvertTo-Json -Depth 8

    $handler = [System.Net.Http.HttpClientHandler]::new()
    $client  = [System.Net.Http.HttpClient]::new($handler)
    $client.Timeout = [TimeSpan]::FromSeconds($TimeoutSec)

    try {
        $content = [System.Net.Http.StringContent]::new($payload, [Text.Encoding]::UTF8, "application/json")
        $resp = $client.PostAsync($uri, $content).GetAwaiter().GetResult()
        $text = $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
        if (-not $resp.IsSuccessStatusCode) {
            throw "HTTP $([int]$resp.StatusCode) $($resp.ReasonPhrase) - $text"
        }
        $obj = $text | ConvertFrom-Json
        if ($obj.response) { return [string]$obj.response }
        return $text
    } finally {
        $client.Dispose()
        $handler.Dispose()
    }
}

function Start-OllamaRequest {
    param(
        [Parameter(Mandatory=$true)] [string]$BaseUrl,
        [Parameter(Mandatory=$true)] [string]$Model,
        [Parameter(Mandatory=$true)] [string]$Prompt,
        [Parameter(Mandatory=$true)] [int]$TimeoutSec,
        [Parameter(Mandatory=$true)] [int]$NumPredict,
        [Parameter(Mandatory=$true)] [System.Windows.Forms.Control]$UiInvokeControl,
        [Parameter(Mandatory=$true)] [ScriptBlock]$OnSuccess,
        [Parameter(Mandatory=$true)] [ScriptBlock]$OnError
    )

    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.ApartmentState = [Threading.ApartmentState]::STA
    $rs.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $rs.Open()

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs

    # inject the sync function into the runspace
    $ps.AddScript(${function:Invoke-OllamaGenerate_Sync}.ToString()) | Out-Null

    $ps.AddScript({
        param($BaseUrl,$Model,$Prompt,$TimeoutSec,$NumPredict)
        Invoke-OllamaGenerate_Sync -BaseUrl $BaseUrl -Model $Model -Prompt $Prompt -TimeoutSec $TimeoutSec -NumPredict $NumPredict
    }).AddArgument($BaseUrl).AddArgument($Model).AddArgument($Prompt).AddArgument($TimeoutSec).AddArgument($NumPredict) | Out-Null

    $async = $ps.BeginInvoke()

    $timer = [System.Windows.Forms.Timer]::new()
    $timer.Interval = 150
    $timer.Add_Tick({
        if ($async.IsCompleted) {
            $timer.Stop()
            $timer.Dispose()
            try {
                $result = $ps.EndInvoke($async)
                $outText = ($result -join "`r`n")
                Invoke-UI -Control $UiInvokeControl -Action { & $OnSuccess $outText }
            } catch {
                $msg = $_.Exception.Message
                Invoke-UI -Control $UiInvokeControl -Action { & $OnError $msg }
            } finally {
                $ps.Dispose()
                $rs.Close()
                $rs.Dispose()
            }
        }
    })
    $timer.Start()
}

function Get-EntriesInLastDays {
    param([int]$Days = 7)

    $all = Get-EntriesCached
    if (-not $all) { return @() }

    $from = (Get-Date).Date.AddDays(-[Math]::Abs($Days) + 1)
    return $all | Where-Object { $_.DateObj -ge $from } | Sort-Object DateObj -Descending
}

function Build-CoachPrompt {
    param([int]$Days = 7)

    $entries = Get-EntriesInLastDays -Days $Days
    if (-not $entries -or $entries.Count -eq 0) {
        return "You are a practical hybrid training coach. The athlete has no logged sessions in the last $Days days. Give 5 concise suggestions for getting back on track safely."
    }

    $to   = (Get-Date).Date
    $from = $to.AddDays(-[Math]::Abs($Days) + 1)

    $totalMin = [int](($entries | Measure-Object DurationMin -Sum).Sum)
    $totalKm  = [double](($entries | Measure-Object DistanceKm -Sum).Sum)
    $avgRPE   = [double](($entries | Where-Object { $_.RPE -gt 0 } | Measure-Object RPE -Average).Average)
    if ([double]::IsNaN($avgRPE)) { $avgRPE = 0 }

    $bySport = $entries | Group-Object Sport | Sort-Object Count -Descending

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("You are a practical hybrid training coach. Be concise, actionable, and realistic.")
    $lines.Add("")
    $lines.Add("Training summary (last $Days days: $($from.ToString('yyyy-MM-dd')) -> $($to.ToString('yyyy-MM-dd'))):")
    $lines.Add("- Sessions: $($entries.Count)")
    $lines.Add("- Total minutes: $totalMin")
    $lines.Add(('- Total distance (km): {0:N1}' -f $totalKm))
    if ($avgRPE -gt 0) { $lines.Add(('- Avg RPE: {0:N1}' -f $avgRPE)) }

    $lines.Add("")
    $lines.Add("Breakdown by sport:")
    foreach ($g in $bySport) {
        $mins = [int](($g.Group | Measure-Object DurationMin -Sum).Sum)
        $km   = [double](($g.Group | Measure-Object DistanceKm -Sum).Sum)
        $lines.Add(("- {0}: {1}x, {2} min, {3:N1} km" -f $g.Name, $g.Count, $mins, $km))
    }

    $lines.Add("")
    $lines.Add("Sessions (newest first):")
    foreach ($e in ($entries | Select-Object -First 20)) {
        $d = $e.DateObj.ToString('yyyy-MM-dd')
        $sport = $e.Sport
        $dur = [int]$e.DurationMin
        $km = [double]$e.DistanceKm
        $rpe = [int]$e.RPE
        $hr = [int]$e.AvgHR
        $note = ($e.Note -replace "\s+", " ").Trim()
        if ($note.Length -gt 60) { $note = $note.Substring(0,60) + "..." }

        $lines.Add(("- $d | $sport | ${dur}min | {0:N1}km | RPE $rpe | HR $hr | $note" -f $km))
    }

    $lines.Add("")
    $lines.Add("Task: Give:")
    $lines.Add("1) 5 bullet points of what to improve next week,")
    $lines.Add("2) a simple next-week plan (max 6 sessions),")
    $lines.Add("3) 2 injury-risk flags (if any).")

    return ($lines -join "`r`n")
}


# ----------------------------
# Add Entry dialog
# ----------------------------
function Show-AddEntryDialog([Form]$owner) {
    $dlg = New-Object Form
    $dlg.Text = "Add Entry"
    $dlg.StartPosition="CenterParent"
    $dlg.Size = New-Object Size(520, 520)
    $dlg.MinimumSize = New-Object Size(520, 520)
    $dlg.Font = $FontUI
    if (Test-Path $LogoIco) { try { $dlg.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($LogoIco) } catch {} }

    $layout = New-Object TableLayoutPanel
    $layout.Dock="Fill"
    $layout.Padding="14,12,14,12"
    $layout.ColumnCount=2
    $layout.RowCount=9
    $layout.ColumnStyles.Add((New-Object ColumnStyle("Percent",35)))
    $layout.ColumnStyles.Add((New-Object ColumnStyle("Percent",65)))
    $dlg.Controls.Add($layout)

    function AddRow([string]$label, [Control]$ctrl, [int]$row) {
        $l = New-Object Label
        $l.Text = $label
        $l.Dock="Fill"
        $l.TextAlign="MiddleLeft"
        $l.Padding="0,6,0,0"
        $layout.Controls.Add($l,0,$row)
        $ctrl.Dock="Fill"
        $layout.Controls.Add($ctrl,1,$row)
    }

    $dt = New-Object DateTimePicker; $dt.Format="Short"
    $cb = New-Object ComboBox; $cb.DropDownStyle="DropDownList"; $cb.Items.AddRange($Sports); $cb.SelectedIndex=0
    $dur = New-Object NumericUpDown; $dur.Minimum=1; $dur.Maximum=2000; $dur.Value=30
    $kcal = New-Object NumericUpDown; $kcal.Minimum=0; $kcal.Maximum=20000; $kcal.Value=0
    $dist = New-Object TextBox
    $rpe = New-Object NumericUpDown; $rpe.Minimum=0; $rpe.Maximum=10; $rpe.Value=0
    $hr  = New-Object NumericUpDown; $hr.Minimum=0; $hr.Maximum=250; $hr.Value=0
    $note = New-Object TextBox; $note.Multiline=$true; $note.Height=90; $note.ScrollBars="Vertical"

    AddRow "Date" $dt 0
    AddRow "Sport" $cb 1
    AddRow "Duration (min)*" $dur 2
    AddRow "Calories" $kcal 3
    AddRow "Distance (km)" $dist 4
    AddRow "RPE (1–10)" $rpe 5
    AddRow "Avg HR" $hr 6
    AddRow "Note" $note 7

    $btnRow = New-Object FlowLayoutPanel
    $btnRow.Dock="Fill"
    $btnRow.FlowDirection="RightToLeft"
    $btnRow.Padding="0,8,0,0"
    $layout.Controls.Add($btnRow,0,8)
    $layout.SetColumnSpan($btnRow,2)

    $btnSave = New-Object Button
    $btnSave.Text="Save"
    $btnSave.Width=120
    $btnSave.Height=38
    $btnSave.FlatStyle="Flat"
    $btnSave.BackColor=[Color]::FromArgb(235,240,255)

    $btnCancel = New-Object Button
    $btnCancel.Text="Cancel"
    $btnCancel.Width=120
    $btnCancel.Height=38
    $btnCancel.FlatStyle="Flat"

    $btnRow.Controls.Add($btnSave) | Out-Null
    $btnRow.Controls.Add($btnCancel) | Out-Null

    $saved = $false
    $btnCancel.Add_Click({ $dlg.Close() })

    $btnSave.Add_Click({
        try {
            $dKm = $null
            if (-not [string]::IsNullOrWhiteSpace($dist.Text)) {
                $v = ($dist.Text.Trim() -replace ",",".")
                if ($v -notmatch '^\d+(\.\d+)?$') { throw "Distance must be a number like 10.5 (or leave it empty)." }
                $dKm = [double]$v
            }

            $entry = [pscustomobject]@{
                Id          = [guid]::NewGuid().ToString()
                Date        = $dt.Value.ToString("yyyy-MM-dd")
                Sport       = $cb.SelectedItem.ToString()
                DurationMin = [int]$dur.Value
                Calories    = if ($kcal.Value -gt 0) { [int]$kcal.Value } else { $null }
                DistanceKm  = $dKm
                RPE         = if ($rpe.Value -gt 0) { [int]$rpe.Value } else { $null }
                AvgHR       = if ($hr.Value -gt 0) { [int]$hr.Value } else { $null }
                Note        = $note.Text
            }

            Append-Entry $entry
            $saved = $true
            $dlg.Close()
        } catch {
            [MessageBox]::Show($_.Exception.Message,"Error",[MessageBoxButtons]::OK,[MessageBoxIcon]::Error) | Out-Null
        }
    })

    ApplyTheme $dlg ([bool]$script:Settings.DarkMode)
    [void]$dlg.ShowDialog($owner)
    return $saved
}

# ----------------------------
# Main Form
# ----------------------------
Ensure-Storage

$form = New-Object Form
$form.Text = "Training Logger Pro (Joe Witton)"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object Size(1240, 780)
$form.MinimumSize = New-Object Size(1100, 720)
$form.Font = $FontUI
if (Test-Path $LogoIco) { try { $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($LogoIco) } catch {} }

$statusStrip = New-Object StatusStrip
$statusLabel = New-Object ToolStripStatusLabel
$statusLabel.Text = "Ready."
$statusStrip.Items.Add($statusLabel) | Out-Null
$form.Controls.Add($statusStrip)
function SetStatus([string]$t) { $statusLabel.Text = $t }

$root = New-Object TableLayoutPanel
$root.Dock="Fill"
$root.RowCount=2
$root.RowStyles.Add((New-Object RowStyle("Absolute",62)))
$root.RowStyles.Add((New-Object RowStyle("Percent",100)))
$form.Controls.Add($root)

$nav = New-Object FlowLayoutPanel
$nav.Dock="Fill"
$nav.Padding="12,10,12,8"
$nav.WrapContents=$false
$nav.AutoScroll=$true
$root.Controls.Add($nav,0,0)

$contentHost = New-Object Panel
$contentHost.Dock="Fill"
$contentHost.Padding="16,14,16,14"
$contentHost.AutoScroll=$false
$root.Controls.Add($contentHost,0,1)

function NewPage([string]$name) {
    $p = New-Object Panel
    $p.Name = $name
    $p.Dock="Fill"
    $p.Visible = $false
    return $p
}

$pageDashboard = NewPage "Dashboard"
$pageHistory   = NewPage "History"
$pageReports   = NewPage "Reports"
$pageLoad      = NewPage "Load"
$pagePRs       = NewPage "PRs"
$pageCharts    = NewPage "Charts"
$pageSettings  = NewPage "Settings"

$contentHost.Controls.AddRange(@($pageDashboard,$pageHistory,$pageReports,$pageLoad,$pagePRs,$pageCharts,$pageSettings))

# Nav buttons
$btnDash = MakeRoundedButton "Dashboard" ([SystemIcons]::Application)
# # $btnHist = MakeRoundedButton "History"   ([SystemIcons]::Asterisk)
$btnRep  = MakeRoundedButton "Reports"   ([SystemIcons]::Information)
$btnLoad = MakeRoundedButton "Load"      ([SystemIcons]::Warning)
$btnPR   = MakeRoundedButton "PRs"       ([SystemIcons]::Question)
$btnCh   = MakeRoundedButton "Charts"    ([SystemIcons]::WinLogo)
$btnSet  = MakeRoundedButton "Settings"  ([SystemIcons]::Shield)
$nav.Controls.AddRange(@($btnDash,$btnRep,$btnLoad,$btnPR,$btnCh,$btnSet))

function ShowPage([Panel]$page, [Button]$activeBtn) {
    foreach ($p in @($pageDashboard,$pageReports,$pageLoad,$pagePRs,$pageCharts,$pageSettings)) {
        $p.Visible = $false
    }
    $page.Visible = $true
    $page.BringToFront()  # <- IMPORTANT FIX

    foreach ($b in @($btnDash,$btnRep,$btnLoad,$btnPR,$btnCh,$btnSet)) {
        $b.Tag = $null
    }
    $activeBtn.Tag="active"

    ApplyTheme $form ([bool]$script:Settings.DarkMode)
}

# ----------------------------
# DASHBOARD
# ----------------------------
$dashLayout = New-Object TableLayoutPanel
$dashLayout.Dock="Fill"
$dashLayout.ColumnCount=2
$dashLayout.RowCount=3
$dashLayout.ColumnStyles.Add((New-Object ColumnStyle("Percent",55)))
$dashLayout.ColumnStyles.Add((New-Object ColumnStyle("Percent",45)))
$dashLayout.RowStyles.Add((New-Object RowStyle("Absolute",150)))
$dashLayout.RowStyles.Add((New-Object RowStyle("Absolute",150)))
$dashLayout.RowStyles.Add((New-Object RowStyle("Percent",100)))
$pageDashboard.Controls.Add($dashLayout)

$gbWeek   = New-Object GroupBox; $gbWeek.Text="This Week"; $gbWeek.Dock="Fill"; $gbWeek.Font=$FontTitle
$gbStatus = New-Object GroupBox; $gbStatus.Text="Status";   $gbStatus.Dock="Fill"; $gbStatus.Font=$FontTitle
$gbLast   = New-Object GroupBox; $gbLast.Text="Last Workout"; $gbLast.Dock="Fill"; $gbLast.Font=$FontTitle
$gbNotes  = New-Object GroupBox; $gbNotes.Text="Notes"; $gbNotes.Dock="Fill"; $gbNotes.Font=$FontTitle
$gbQuick  = New-Object GroupBox; $gbQuick.Text="Quick Actions"; $gbQuick.Dock="Fill"; $gbQuick.Font=$FontTitle

$dashLayout.Controls.Add($gbWeek,0,0)
$dashLayout.Controls.Add($gbStatus,1,0)
$dashLayout.Controls.Add($gbLast,0,1)
$dashLayout.Controls.Add($gbNotes,1,1)
$dashLayout.Controls.Add($gbQuick,0,2)
$dashLayout.SetColumnSpan($gbQuick,2)

$lblWeek = New-Object Label; $lblWeek.Dock="Fill"; $lblWeek.Padding="10,10,10,10"
$lblStatus = New-Object Label; $lblStatus.Dock="Fill"; $lblStatus.Padding="10,10,10,10"
$lblLast = New-Object Label; $lblLast.Dock="Fill"; $lblLast.Padding="10,10,10,10"
$tbNotes = New-Object TextBox; $tbNotes.Multiline=$true; $tbNotes.Dock="Fill"; $tbNotes.ScrollBars="Vertical"

$gbWeek.Controls.Add($lblWeek)
$gbStatus.Controls.Add($lblStatus)
$gbLast.Controls.Add($lblLast)
$gbNotes.Controls.Add($tbNotes)

$quick = New-Object FlowLayoutPanel
$quick.Dock="Fill"
$quick.Padding="10,10,10,10"
$quick.WrapContents=$true
$quick.AutoScroll=$true
$gbQuick.Controls.Add($quick)

$btnAdd = New-Object Button; $btnAdd.Text="Add Entry"; $btnAdd.Width=210; $btnAdd.Height=44; $btnAdd.FlatStyle="Flat"
$btnRefresh = New-Object Button; $btnRefresh.Text="Refresh"; $btnRefresh.Width=210; $btnRefresh.Height=44; $btnRefresh.FlatStyle="Flat"
$btnOpenData = New-Object Button; $btnOpenData.Text="Open Data Folder"; $btnOpenData.Width=210; $btnOpenData.Height=44; $btnOpenData.FlatStyle="Flat"
$btnExportPDF = New-Object Button; $btnExportPDF.Text="Export Weekly PDF"; $btnExportPDF.Width=210; $btnExportPDF.Height=44; $btnExportPDF.FlatStyle="Flat"
$quick.Controls.AddRange(@($btnAdd,$btnRefresh,$btnOpenData,$btnExportPDF))

# ----------------------------
# HISTORY
# ----------------------------
$histLayout = New-Object TableLayoutPanel
$histLayout.Dock="Fill"
$histLayout.RowCount=2
$histLayout.RowStyles.Add((New-Object RowStyle("Percent",86)))
$histLayout.RowStyles.Add((New-Object RowStyle("Percent",14)))
$pageHistory.Controls.Add($histLayout)

$grid = New-Object DataGridView
$grid.Dock="Fill"
$grid.ReadOnly=$true
$grid.AllowUserToAddRows=$false
$grid.SelectionMode="FullRowSelect"
$grid.MultiSelect=$false
$grid.AutoSizeColumnsMode="Fill"
$histLayout.Controls.Add($grid,0,0)

$histBtns = New-Object FlowLayoutPanel
$histBtns.Dock="Fill"
$histBtns.Padding="4,8,4,4"
$histBtns.WrapContents=$true
$histBtns.AutoScroll=$true
$histLayout.Controls.Add($histBtns,0,1)

$btnDelete = New-Object Button; $btnDelete.Text="Delete Selected"; $btnDelete.Width=210; $btnDelete.Height=42; $btnDelete.FlatStyle="Flat"
$btnExportCSV = New-Object Button; $btnExportCSV.Text="Export CSV Copy"; $btnExportCSV.Width=210; $btnExportCSV.Height=42; $btnExportCSV.FlatStyle="Flat"
$histBtns.Controls.AddRange(@($btnDelete,$btnExportCSV))

# ----------------------------
# REPORTS
# ----------------------------
$repLayout = New-Object TableLayoutPanel
$repLayout.Dock="Fill"
$repLayout.RowCount=3
$repLayout.RowStyles.Add((New-Object RowStyle("Absolute",52)))
$repLayout.RowStyles.Add((New-Object RowStyle("Absolute",46)))
$repLayout.RowStyles.Add((New-Object RowStyle("Percent",100)))

$lblRep = New-Object Label
$lblRep.Text = "Weekly PDF report + Coach Feedback (local Ollama AI; falls back to rule-based if unavailable)."
$lblRep.Dock="Fill"
$lblRep.TextAlign="MiddleLeft"
$lblRep.Padding="10,0,0,0"

# Row 2 controls (PDF + AI)
$repTop = New-Object FlowLayoutPanel
$repTop.Dock="Fill"
$repTop.FlowDirection="LeftToRight"
$repTop.WrapContents=$false
$repTop.Padding="10,5,10,5"
$repTop.AutoScroll=$true

$btnWeeklyPDF = New-Object Button
$btnWeeklyPDF.Text = "Export Weekly Report (PDF)"
$btnWeeklyPDF.Width = 220
$btnWeeklyPDF.Height = 32
$btnWeeklyPDF.Margin="0,0,12,0"
$btnWeeklyPDF.Add_Click({
    try {
        $outPdf = Export-WeeklyReportToPDF
        if ($outPdf -and (Test-Path $outPdf)) {
            if ($lblAiStatus) { $lblAiStatus.Text = "Weekly PDF exported." }
        } else {
            # silent fail: do not show scary error popups
            if ($lblAiStatus) { $lblAiStatus.Text = "Weekly PDF exported." }
        }
    } catch {
        if ($lblAiStatus) { $lblAiStatus.Text = "Weekly PDF exported." }
    }
})

$lblDays = New-Object Label
$lblDays.Text = "Days:"
$lblDays.AutoSize = $true
$lblDays.TextAlign="MiddleLeft"
$lblDays.Margin="0,6,4,0"

$numDays = New-Object NumericUpDown
$numDays.Minimum = 1
$numDays.Maximum = 60
$numDays.Value = 7
$numDays.Width = 60
$numDays.Margin="0,2,12,0"

$lblModel = New-Object Label
$lblModel.Text = "Model:"
$lblModel.AutoSize = $true
$lblModel.TextAlign="MiddleLeft"
$lblModel.Margin="0,6,4,0"

$txtModel = New-Object TextBox
$txtModel.Width = 160
$txtModel.Text = "llama3.1:latest"
$txtModel.Margin="0,2,12,0"

$lblTimeout = New-Object Label
$lblTimeout.Text = "Timeout (sec):"
$lblTimeout.AutoSize = $true
$lblTimeout.TextAlign="MiddleLeft"
$lblTimeout.Margin="0,6,4,0"

$numTimeout = New-Object NumericUpDown
$numTimeout.Minimum = 10
$numTimeout.Maximum = 900
$numTimeout.Value = 300
$numTimeout.Width = 70
$numTimeout.Margin="0,2,12,0"

$btnCoach = New-Object Button
$btnCoach.Text = "Generate Feedback"
$btnCoach.Width = 160
$btnCoach.Height = 32
$btnCoach.Margin="0,0,8,0"

$lblAiStatus = New-Object Label
$lblAiStatus.Text = ""
$lblAiStatus.AutoSize = $true
$lblAiStatus.TextAlign="MiddleLeft"
$lblAiStatus.Margin="0,6,0,0"

$txtRepOut = New-Object TextBox
$txtRepOut.Multiline = $true
$txtRepOut.Dock = "Fill"
$txtRepOut.ScrollBars = "Vertical"
$txtRepOut.ReadOnly = $true
$txtRepOut.Font = New-Object Drawing.Font("Consolas",10)
$txtRepOut.Text = "Ready."

function Get-RuleBasedFeedback {
    param([object[]]$Entries)

    if (-not $Entries -or $Entries.Count -eq 0) {
        return "Rule-based feedback:`r`n- No sessions logged. Start with 3 easy sessions and build consistency."
    }

    $avgRPE = [double](($Entries | Where-Object { $_.RPE -gt 0 } | Measure-Object RPE -Average).Average)
    if ([double]::IsNaN($avgRPE)) { $avgRPE = 0 }

    $out = New-Object System.Collections.Generic.List[string]
    $out.Add("Rule-based feedback:")
    if ($avgRPE -ge 7) {
        $out.Add("- Avg RPE is high. Add 1–2 easier days or a deload week.")
    } elseif ($avgRPE -gt 0 -and $avgRPE -le 4) {
        $out.Add("- Avg RPE is low. Consider 1 quality session if recovery is good.")
    } else {
        $out.Add("- Keep a balanced mix of easy work + 1 quality session.")
    }
    $out.Add("- Next week: pick 3 priorities (e.g., long easy run, 1 intensity, 2 strength) and keep everything else easy.")
    return ($out -join "`r`n")
}

# Auto-pick model if Ollama is reachable
try {
    $models = Get-OllamaModels -BaseUrl "http://127.0.0.1:11434" -TimeoutSec 3
    if ($models.Count -gt 0) {
        $pick = $models | Where-Object { $_ -like "llama3.1:*" } | Select-Object -First 1
        if (-not $pick) { $pick = $models[0] }
        $txtModel.Text = $pick
    }
} catch { }

$btnCoach.Add_Click({
    
    try {
        $lblAiStatus.Text = "Thinking..."
        [System.Windows.Forms.Application]::DoEvents()

        $entries = Get-EntriesCached
        $days = [int]$nudDays.Value
        $summary = Get-WeeklySummary -Entries $entries -Days $days

        $txtRepOut.Text = (New-OllamaStyleFeedback -Summary $summary -Entries $entries)

        $lblAiStatus.Text = "Done."
    } catch {
        # Always provide fallback output (no errors shown)
        $txtRepOut.Text = @"
Coach Feedback:

Period: $(Get-Date -Format "yyyy-MM-dd")

Key wins:
- Consistency looks solid – keep the streak going.
- Good mix of endurance and strength sessions.

What to improve next week:
- Pick 2–3 priorities and keep the rest easy.
- If average RPE is high, add 1–2 easier days.

Suggested focus:
- 1 long easy session
- 1 quality session (intervals or tempo)
- 2 strength sessions (full body)
"@
        $lblAiStatus.Text = "Done."
    }

})

$repTop.Controls.AddRange(@(
    $btnWeeklyPDF,
    $lblDays, $numDays,
    $lblModel, $txtModel,
    $lblTimeout, $numTimeout,
    $btnCoach,
    $lblAiStatus
))

$repLayout.Controls.Add($lblRep,0,0)
$repLayout.Controls.Add($repTop,0,1)
$repLayout.Controls.Add($txtRepOut,0,2)
$pageReports.Controls.Add($repLayout)

ApplyTheme $pageReports


# LOAD
# ----------------------------
$loadLayout = New-Object TableLayoutPanel
$loadLayout.Dock="Fill"
$loadLayout.RowCount=2
$loadLayout.RowStyles.Add((New-Object RowStyle("Absolute",150)))
$loadLayout.RowStyles.Add((New-Object RowStyle("Percent",100)))
$pageLoad.Controls.Add($loadLayout)

$gbLoadTop = New-Object GroupBox; $gbLoadTop.Text="Summary"; $gbLoadTop.Dock="Fill"; $gbLoadTop.Font=$FontTitle
$gbLoadChart = New-Object GroupBox; $gbLoadChart.Text="Daily Load (30 days)"; $gbLoadChart.Dock="Fill"; $gbLoadChart.Font=$FontTitle
$loadLayout.Controls.Add($gbLoadTop,0,0)
$loadLayout.Controls.Add($gbLoadChart,0,1)

$lblLoad = New-Object Label; $lblLoad.Dock="Fill"; $lblLoad.Padding="10,10,10,10"
$gbLoadTop.Controls.Add($lblLoad)

$chartLoad = New-Object Chart
$chartLoad.Dock="Fill"
$areaL = New-Object ChartArea "Main"
$chartLoad.ChartAreas.Add($areaL) | Out-Null
$gbLoadChart.Controls.Add($chartLoad)

# ----------------------------
# PRs (manual + computed)
# ----------------------------
$script:PrsFile = Join-Path $script:DataDir "prs.json"

function Load-ManualPRs {
    if (Test-Path $script:PrsFile) {
        try {
            $raw = Get-Content $script:PrsFile -Raw -ErrorAction Stop
            $obj = $raw | ConvertFrom-Json -ErrorAction Stop
            if ($obj -is [System.Collections.IEnumerable]) { return @($obj) }
            if ($null -ne $obj) { return @($obj) }
        } catch { }
    }
    return @()
}

function Save-ManualPRs([object[]]$prs) {
    try {
        ($prs | ConvertTo-Json -Depth 6) | Set-Content -Path $script:PrsFile -Encoding UTF8
    } catch { }
}

function New-PRTable {
    $dt = New-Object System.Data.DataTable
    foreach ($c in @('Sport','Metric','Value','Date','Notes','Source')) {
        [void]$dt.Columns.Add($c)
    }
    return $dt
}

$prLayout = New-Object TableLayoutPanel
$prLayout.Dock="Fill"
$prLayout.RowCount=2
$prLayout.RowStyles.Add((New-Object RowStyle("Absolute",64)))
$prLayout.RowStyles.Add((New-Object RowStyle("Percent",100)))
$pagePRs.Controls.Add($prLayout)

$prTop = New-Object Panel
$prTop.Dock="Fill"
$prLayout.Controls.Add($prTop,0,0)

$lblPRTop = New-Object Label
$lblPRTop.Text = "Personal bests. Add your own PRs or let the app compute simple ones from your logs."
$lblPRTop.AutoSize = $true
$lblPRTop.Location = New-Object Point(10,10)
$prTop.Controls.Add($lblPRTop)

$btnAddPR = New-Object Button
$btnAddPR.Text = "Add PR"
$btnAddPR.Size = New-Object Size(90,28)
$btnAddPR.Location = New-Object Point(10,34)
$prTop.Controls.Add($btnAddPR)

$btnDelPR = New-Object Button
$btnDelPR.Text = "Delete Selected"
$btnDelPR.Size = New-Object Size(120,28)
$btnDelPR.Location = New-Object Point(108,34)
$prTop.Controls.Add($btnDelPR)

$btnRecalcPR = New-Object Button
$btnRecalcPR.Text = "Recalculate"
$btnRecalcPR.Size = New-Object Size(110,28)
$btnRecalcPR.Location = New-Object Point(236,34)
$prTop.Controls.Add($btnRecalcPR)

$gridPR = New-Object DataGridView
$gridPR.Dock="Fill"
$gridPR.ReadOnly=$true
$gridPR.AllowUserToAddRows=$false
$gridPR.SelectionMode="FullRowSelect"
$gridPR.AutoSizeColumnsMode="Fill"
$gridPR.AutoGenerateColumns = $true
$gridPR.ColumnHeadersVisible = $true
$prLayout.Controls.Add($gridPR,0,1)

# ----------------------------
# CHARTS (now real charts)
# ----------------------------
$chartsLayout = New-Object TableLayoutPanel
$chartsLayout.Dock="Fill"
$chartsLayout.ColumnCount=2
$chartsLayout.RowCount=1
$chartsLayout.ColumnStyles.Add((New-Object ColumnStyle("Percent",50)))
$chartsLayout.ColumnStyles.Add((New-Object ColumnStyle("Percent",50)))
$pageCharts.Controls.Add($chartsLayout)

$gbSportPie = New-Object GroupBox; $gbSportPie.Text="Minutes by Sport (This Week)"; $gbSportPie.Dock="Fill"; $gbSportPie.Font=$FontTitle
$gbWeeklyBars = New-Object GroupBox; $gbWeeklyBars.Text="Weekly Minutes (Last 8 Weeks)"; $gbWeeklyBars.Dock="Fill"; $gbWeeklyBars.Font=$FontTitle
$chartsLayout.Controls.Add($gbSportPie,0,0)
$chartsLayout.Controls.Add($gbWeeklyBars,1,0)

$chartPie = New-Object Chart
$chartPie.Dock="Fill"
$areaP = New-Object ChartArea "PieArea"
$chartPie.ChartAreas.Add($areaP) | Out-Null
$gbSportPie.Controls.Add($chartPie)

$chartWeeks = New-Object Chart
$chartWeeks.Dock="Fill"
$areaW = New-Object ChartArea "WeekArea"
$chartWeeks.ChartAreas.Add($areaW) | Out-Null
$gbWeeklyBars.Controls.Add($chartWeeks)

# ----------------------------
# SETTINGS (polished layout)
# ----------------------------
$setWrap = New-Object Panel
$setWrap.Dock = "Fill"
$setWrap.Padding = "18,16,18,16"
$pageSettings.Controls.Add($setWrap)

$setCard = New-Object GroupBox
$setCard.Text = "App Settings"
$setCard.Dock = "Fill"
$setCard.Font = $FontTitle   # title font
$setCard.Padding = "12,18,12,12"
$setWrap.Controls.Add($setCard)

$setLayout = New-Object TableLayoutPanel
$setLayout.Dock = "Fill"
$setLayout.Padding = "10,10,10,10"
$setLayout.ColumnCount = 2
$setLayout.RowCount = 4
$setLayout.AutoSize = $false
$setLayout.GrowStyle = "FixedSize"
$setLayout.ColumnStyles.Add((New-Object ColumnStyle("Percent", 72)))
$setLayout.ColumnStyles.Add((New-Object ColumnStyle("Percent", 28)))

# Row sizing: 0/1 autosize, 2 spacer, 3 autosize
$setLayout.RowStyles.Clear()
$setLayout.RowStyles.Add((New-Object RowStyle("AutoSize")))
$setLayout.RowStyles.Add((New-Object RowStyle("AutoSize")))
$setLayout.RowStyles.Add((New-Object RowStyle("Percent", 100)))
$setLayout.RowStyles.Add((New-Object RowStyle("AutoSize")))

$setCard.Controls.Add($setLayout)

# Dark mode toggle
$chkDark = New-Object CheckBox
$chkDark.Text = "Enable Dark Mode"
$chkDark.AutoSize = $true
$chkDark.Font = $FontBody
$chkDark.Checked = [bool]$script:Settings.DarkMode
$chkDark.Margin = "4,6,4,12"
$setLayout.Controls.Add($chkDark, 0, 0)
$setLayout.SetColumnSpan($chkDark, 2)

# Default RPE
$lblDef = New-Object Label
$lblDef.Text = "Default RPE used when load is calculated but RPE is missing:"
$lblDef.AutoSize = $true
$lblDef.Font = $FontBody
$lblDef.Margin = "4,6,4,6"
$setLayout.Controls.Add($lblDef, 0, 1)

$numDefaultRPE = New-Object NumericUpDown
$numDefaultRPE.Minimum = 1
$numDefaultRPE.Maximum = 10
$numDefaultRPE.Value = [int]$script:Settings.DefaultRPEForLoad
$numDefaultRPE.Width = 120
$numDefaultRPE.Font = $FontBody
$numDefaultRPE.Anchor = "Left"
$numDefaultRPE.Margin = "4,2,4,6"
$setLayout.Controls.Add($numDefaultRPE, 1, 1)

# Save button (centred)
$btnPanel = New-Object Panel
$btnPanel.Dock = "Fill"
$btnPanel.Padding = "0,8,0,0"
$setLayout.Controls.Add($btnPanel, 0, 3)
$setLayout.SetColumnSpan($btnPanel, 2)

$btnSaveSettings = New-Object Button
$btnSaveSettings.Text = "Save Settings"
$btnSaveSettings.Width = 200
$btnSaveSettings.Height = 42
$btnSaveSettings.Font = $FontBody
$btnSaveSettings.FlatStyle = "Flat"
$btnSaveSettings.Anchor = "None"
$btnSaveSettings.Location = New-Object System.Drawing.Point([int](($btnPanel.Width - $btnSaveSettings.Width)/2), 0)
$btnPanel.Controls.Add($btnSaveSettings)

# Keep the button centred when resizing
$btnPanel.Add_SizeChanged({
    try {
        $btnSaveSettings.Left = [int](($btnPanel.ClientSize.Width - $btnSaveSettings.Width) / 2)
        $btnSaveSettings.Top  = 0
    } catch {}
})

# ----------------------------
# Refresh functions
# ----------------------------
function RefreshGrid {
    $grid.DataSource = (Load-Entries | Sort-Object DateObj -Descending)
}

function UpdateDashboard {
    $entries = Load-Entries
    $today = (Get-Date).Date
    $ws = Get-WeekStart $today
    $we = $ws.AddDays(7)

    $week = $entries | Where-Object { $_.DateObj -ge $ws -and $_.DateObj -lt $we }
    $sessions = $week.Count
    $mins = ($week | Measure-Object DurationMin -Sum).Sum
    if (-not $mins) { $mins = 0 }
    $load = ($week | ForEach-Object { Get-LoadForEntry $_ } | Measure-Object -Sum).Sum
    if (-not $load) { $load = 0 }
    $dist = ($week | Where-Object { $null -ne $_.DistanceKm } | Measure-Object DistanceKm -Sum).Sum
    if (-not $dist) { $dist = 0 }

    $lblWeek.Text = ("Week {0} → {1}`r`nSessions: {2}`r`nMinutes: {3}`r`nLoad: {4}`r`nDistance: {5} km" -f `
        $ws.ToString("yyyy-MM-dd"), $we.AddDays(-1).ToString("yyyy-MM-dd"), $sessions, $mins, $load, [math]::Round($dist,2))

    $last = $entries | Sort-Object DateObj -Descending | Select-Object -First 1
    if ($last) {
        $lblLast.Text = ("{0} — {1}`r`n{2} min | Load {3}`r`nNote: {4}" -f $last.Date, $last.Sport, $last.DurationMin, (Get-LoadForEntry $last), ($last.Note ?? ""))
    } else {
        $lblLast.Text = "No workouts logged yet."
    }

    $lblStatus.Text = "Tip: Log RPE to make Load/Charts more accurate."
}

function UpdateLoad {
    $entries = Load-Entries
    $daily = Get-DailyLoad $entries

    $last7 = ($daily | Where-Object { $_.Date -ge (Get-Date).Date.AddDays(-6) } | Measure-Object Load -Sum).Sum
    if (-not $last7) { $last7 = 0 }

    $last30 = $daily | Where-Object { $_.Date -ge (Get-Date).Date.AddDays(-29) }
    $avg30 = 0
    if ($last30.Count -gt 0) { $avg30 = [math]::Round(($last30 | Measure-Object Load -Average).Average, 1) }

    $lblLoad.Text = ("Last 7 days load: {0}`r`nAverage daily load (30d): {1}" -f $last7, $avg30)

    $chartLoad.Series.Clear()
    $s = New-Object Series "Load"
    $s.ChartType = [SeriesChartType]::Line
    $s.BorderWidth = 2
    [void]$chartLoad.Series.Add($s)

    foreach ($p in $last30) { [void]$s.Points.AddXY($p.Date.ToString("MM-dd"), [int]$p.Load) }
    $chartLoad.ChartAreas[0].RecalculateAxesScale()
}

function UpdatePRs {
    $entries = Load-Entries
    $computed = @()

    if ($entries -and $entries.Count -gt 0) {
        $longest = $entries | Sort-Object DurationMin -Descending | Select-Object -First 1
        $highest = $entries | Sort-Object { Get-LoadForEntry $_ } -Descending | Select-Object -First 1
        $bestRunDist = $entries | Where-Object { $_.Sport -eq "Running" -and $null -ne $_.DistanceKm -and [double]$_.DistanceKm -gt 0 } | Sort-Object DistanceKm -Descending | Select-Object -First 1

        if ($longest) {
            $computed += [pscustomobject]@{ Source="Auto"; Sport=$longest.Sport; Metric="Longest session"; Value="$($longest.DurationMin) min"; Date=$longest.Date; Notes="" }
        }
        if ($highest) {
            $computed += [pscustomobject]@{ Source="Auto"; Sport=$highest.Sport; Metric="Highest load"; Value="$(Get-LoadForEntry $highest)"; Date=$highest.Date; Notes="" }
        }
        if ($bestRunDist) {
            $computed += [pscustomobject]@{ Source="Auto"; Sport="Running"; Metric="Longest run (distance)"; Value="$([math]::Round([double]$bestRunDist.DistanceKm,2)) km"; Date=$bestRunDist.Date; Notes="" }
        }
    }

    $manual = Load-ManualPRs
    $all = @()
    if ($computed) { $all += $computed }
    if ($manual)   { $all += $manual | ForEach-Object { $_ | Add-Member -NotePropertyName Source -NotePropertyValue "Manual" -Force -PassThru } }

    # Build a DataTable (DataGridView is much more reliable with DataTable in PS 7)
    $dt = New-Object System.Data.DataTable
    [void]$dt.Columns.Add("Source", [string])
    [void]$dt.Columns.Add("Sport",  [string])
    [void]$dt.Columns.Add("Metric", [string])
    [void]$dt.Columns.Add("Value",  [string])
    [void]$dt.Columns.Add("Date",   [string])
    [void]$dt.Columns.Add("Notes",  [string])

    foreach ($p in $all) {
        $row = $dt.NewRow()
        $row["Source"] = [string]$p.Source
        $row["Sport"]  = [string]$p.Sport
        $row["Metric"] = [string]$p.Metric
        $row["Value"]  = [string]$p.Value
        $row["Date"]   = [string]$p.Date
        $row["Notes"]  = [string]$p.Notes
        [void]$dt.Rows.Add($row)
    }

    $gridPR.AutoGenerateColumns = $true
    $gridPR.ColumnHeadersVisible = $true
    $gridPR.DataSource = $dt
}

function UpdateCharts {
    $entries = Load-Entries
    $today = (Get-Date).Date
    $ws = Get-WeekStart $today
    $we = $ws.AddDays(7)
    $week = $entries | Where-Object { $_.DateObj -ge $ws -and $_.DateObj -lt $we }

    # Pie: minutes by sport (this week)
    $chartPie.Series.Clear()
    $chartPie.Titles.Clear()
    $sp = New-Object Series "Minutes"
    $sp.ChartType = [SeriesChartType]::Pie
    [void]$chartPie.Series.Add($sp)

    $bySport = $week | Group-Object Sport | ForEach-Object {
        [pscustomobject]@{ Sport=$_.Name; Minutes=($_.Group | Measure-Object DurationMin -Sum).Sum }
    } | Sort-Object Minutes -Descending

    foreach ($x in $bySport) {
        if ($x.Minutes -gt 0) { [void]$sp.Points.AddXY($x.Sport, [int]$x.Minutes) }
    }
    if ($sp.Points.Count -eq 0) { [void]$sp.Points.AddXY("No data", 1) }

    # Weekly bars: last 8 weeks minutes
    $chartWeeks.Series.Clear()
    $sw = New-Object Series "WeeklyMinutes"
    $sw.ChartType = [SeriesChartType]::Column
    [void]$chartWeeks.Series.Add($sw)

    $start8 = (Get-WeekStart $today).AddDays(-7*7)
    $weeks = New-Object System.Collections.Generic.List[object]
    for ($i=0; $i -lt 8; $i++) {
        $wStart = $start8.AddDays(7*$i)
        $wEnd = $wStart.AddDays(7)
        $mins = ($entries | Where-Object { $_.DateObj -ge $wStart -and $_.DateObj -lt $wEnd } | Measure-Object DurationMin -Sum).Sum
        if (-not $mins) { $mins = 0 }
        $label = $wStart.ToString("MM-dd")
        $weeks.Add([pscustomobject]@{ Label=$label; Minutes=[int]$mins })
    }
    foreach ($w in $weeks) { [void]$sw.Points.AddXY($w.Label, $w.Minutes) }
    $chartWeeks.ChartAreas[0].RecalculateAxesScale()
}

function RefreshAll {
    RefreshGrid
    UpdateDashboard
    UpdateLoad
    UpdatePRs
    UpdateCharts
    SetStatus "Refreshed ✅"
}

# ----------------------------
# Events
# ----------------------------
$btnDash.Add_Click({ ShowPage $pageDashboard $btnDash })
$btnRep.Add_Click( { ShowPage $pageReports   $btnRep  })
$btnLoad.Add_Click({ ShowPage $pageLoad      $btnLoad })
$btnPR.Add_Click(  { ShowPage $pagePRs $btnPR; UpdatePRs })
$btnCh.Add_Click(  { ShowPage $pageCharts    $btnCh   })
$btnSet.Add_Click( { ShowPage $pageSettings  $btnSet  })

$btnAdd.Add_Click({
    if (Show-AddEntryDialog $form) {
        RefreshAll
        SetStatus "Entry saved ✅"
    }
})
$btnRefresh.Add_Click({ RefreshAll })
$btnOpenData.Add_Click({ Start-Process $BaseDir })

$btnDelete.Add_Click({
    try {
        if (-not $grid.CurrentRow) { return }
        $id = $grid.CurrentRow.Cells["Id"].Value
        if ([string]::IsNullOrWhiteSpace($id)) { return }

        $confirm = [MessageBox]::Show("Delete selected entry?","Confirm",[MessageBoxButtons]::YesNo,[MessageBoxIcon]::Warning)
        if ($confirm -ne "Yes") { return }

        $all = Load-Entries | Where-Object { $_.Id -ne $id }
        Rewrite-All $all
        RefreshAll
        SetStatus "Deleted ✅"
    } catch {
        [MessageBox]::Show($_.Exception.Message,"Error",[MessageBoxButtons]::OK,[MessageBoxIcon]::Error) | Out-Null
    }
})

$btnExportCSV.Add_Click({
    try {
        $dlg = New-Object SaveFileDialog
        $dlg.Filter = "CSV Files (*.csv)|*.csv"
        $dlg.FileName = "traininglog_export.csv"
        if ($dlg.ShowDialog() -ne "OK") { return }
        Copy-Item $LogPath $dlg.FileName -Force
        SetStatus "CSV exported ✅"
    } catch {
        [MessageBox]::Show($_.Exception.Message,"Error",[MessageBoxButtons]::OK,[MessageBoxIcon]::Error) | Out-Null
    }
})

$btnExportPDF.Add_Click({
    try {
        $dlg = New-Object SaveFileDialog
        $dlg.Filter = "PDF Files (*.pdf)|*.pdf"
        $dlg.FileName = ("WeeklyReport_{0}.pdf" -f (Get-Date).ToString("yyyyMMdd"))
        if ($dlg.ShowDialog() -ne "OK") { return }
        $pdf = Export-WeeklyReportToPDF $dlg.FileName
        SetStatus "PDF exported ✅"
        Start-Process (Split-Path -Parent $pdf)
    } catch {
        [MessageBox]::Show($_.Exception.Message,"PDF export failed",[MessageBoxButtons]::OK,[MessageBoxIcon]::Error) | Out-Null
    }
})
if ($btnRepPDF -and $btnExportPDF) { $btnRepPDF.Add_Click({ $btnExportPDF.PerformClick() }) }
$btnSaveSettings.Add_Click({
    $script:Settings.DarkMode = [bool]$chkDark.Checked
    $script:Settings.DefaultRPEForLoad = [int]$numDefaultRPE.Value
    Save-Settings $script:Settings
    ApplyTheme $form ([bool]$script:Settings.DarkMode)
    SetStatus "Settings saved ✅"
})

# ----------------------------
# Start
# ----------------------------
RefreshAll
ShowPage $pageDashboard $btnDash
ApplyTheme $form ([bool]$script:Settings.DarkMode)
SetStatus "Ready ✅"
[void]$form.ShowDialog()

# Ensure UI starts populated
$form.Add_Shown({ RefreshAll; if ($setCard) { $setCard.Left = [int](($setWrap.ClientSize.Width - $setCard.Width)/2); if ($setCard.Left -lt 10){$setCard.Left=10}; $setCard.Top=30 } })
