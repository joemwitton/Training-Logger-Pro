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
            Date     = [datetime]$_.Name
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
# PDF Export (Edge headless)
# ----------------------------
function Get-EdgePath {
    $candidates = @(
        "$env:ProgramFiles(x86)\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe"
    )
    foreach ($p in $candidates) { if (Test-Path $p) { return $p } }
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
    if (-not $edge) { throw "Microsoft Edge not found. Install Edge (Windows 11 normally has it)." }

    $args = @("--headless","--disable-gpu","--no-first-run","--print-to-pdf=""$pdfPath""",$htmlPath)
    $p = Start-Process -FilePath $edge -ArgumentList $args -PassThru -WindowStyle Hidden
    $p.WaitForExit()

    if (-not (Test-Path $pdfPath)) { throw "PDF export failed (Edge did not create the file)." }
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
$btnHist = MakeRoundedButton "History"   ([SystemIcons]::Asterisk)
$btnRep  = MakeRoundedButton "Reports"   ([SystemIcons]::Information)
$btnLoad = MakeRoundedButton "Load"      ([SystemIcons]::Warning)
$btnPR   = MakeRoundedButton "PRs"       ([SystemIcons]::Question)
$btnCh   = MakeRoundedButton "Charts"    ([SystemIcons]::WinLogo)
$btnSet  = MakeRoundedButton "Settings"  ([SystemIcons]::Shield)
$nav.Controls.AddRange(@($btnDash,$btnHist,$btnRep,$btnLoad,$btnPR,$btnCh,$btnSet))

function ShowPage([Panel]$page, [Button]$activeBtn) {
    foreach ($p in @($pageDashboard,$pageHistory,$pageReports,$pageLoad,$pagePRs,$pageCharts,$pageSettings)) {
        $p.Visible = $false
    }
    $page.Visible = $true
    $page.BringToFront()  # <- IMPORTANT FIX

    foreach ($b in @($btnDash,$btnHist,$btnRep,$btnLoad,$btnPR,$btnCh,$btnSet)) {
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
$repLayout.RowStyles.Add((New-Object RowStyle("Absolute",46)))
$repLayout.RowStyles.Add((New-Object RowStyle("Absolute",52)))
$repLayout.RowStyles.Add((New-Object RowStyle("Percent",100)))
$pageReports.Controls.Add($repLayout)

$lblRep = New-Object Label
$lblRep.Text="Export weekly report as PDF (Mon–Sun)."
$lblRep.Dock="Fill"
$lblRep.Padding="6,10,6,0"
$repLayout.Controls.Add($lblRep,0,0)

$btnRepPDF = New-Object Button
$btnRepPDF.Text="Export Weekly Report (PDF)"
$btnRepPDF.Dock="Left"
$btnRepPDF.Width=270
$btnRepPDF.Height=42
$btnRepPDF.FlatStyle="Flat"
$repLayout.Controls.Add($btnRepPDF,0,1)

$tbRep = New-Object TextBox
$tbRep.Multiline=$true
$tbRep.ReadOnly=$true
$tbRep.ScrollBars="Vertical"
$tbRep.Dock="Fill"
$tbRep.Font = New-Object Font("Consolas", 10)
$repLayout.Controls.Add($tbRep,0,2)

# ----------------------------
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
# PRs (simple but real)
# ----------------------------
$prLayout = New-Object TableLayoutPanel
$prLayout.Dock="Fill"
$prLayout.RowCount=2
$prLayout.RowStyles.Add((New-Object RowStyle("Absolute",46)))
$prLayout.RowStyles.Add((New-Object RowStyle("Percent",100)))
$pagePRs.Controls.Add($prLayout)

$lblPRTop = New-Object Label
$lblPRTop.Text="Personal bests from your logged sessions."
$lblPRTop.Dock="Fill"
$lblPRTop.Padding="6,10,6,0"
$prLayout.Controls.Add($lblPRTop,0,0)

$gridPR = New-Object DataGridView
$gridPR.Dock="Fill"
$gridPR.ReadOnly=$true
$gridPR.AllowUserToAddRows=$false
$gridPR.SelectionMode="FullRowSelect"
$gridPR.AutoSizeColumnsMode="Fill"
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
# SETTINGS (fills nicely now)
# ----------------------------
$setWrap = New-Object Panel
$setWrap.Dock="Fill"
$pageSettings.Controls.Add($setWrap)

$setCard = New-Object GroupBox
$setCard.Text="App Settings"
$setCard.Font=$FontTitle
$setCard.Size = New-Object Size(520,220)
$setCard.Location = New-Object Point(18,18)
$setCard.Anchor = "Top,Left"
$setWrap.Controls.Add($setCard)

$setLayout = New-Object TableLayoutPanel
$setLayout.Dock="Fill"
$setLayout.Padding="12,10,12,10"
$setLayout.RowCount=4
$setLayout.ColumnCount=2
$setLayout.ColumnStyles.Add((New-Object ColumnStyle("Percent",65)))
$setLayout.ColumnStyles.Add((New-Object ColumnStyle("Percent",35)))
$setCard.Controls.Add($setLayout)

$chkDark = New-Object CheckBox
$chkDark.Text="Enable Dark Mode"
$chkDark.Checked=[bool]$script:Settings.DarkMode
$chkDark.Padding="4,6,4,6"
$setLayout.Controls.Add($chkDark,0,0)
$setLayout.SetColumnSpan($chkDark,2)

$lblDef = New-Object Label
$lblDef.Text="Default RPE used for load when RPE is missing:"
$lblDef.Padding="0,8,0,4"
$setLayout.Controls.Add($lblDef,0,1)

$numDefaultRPE = New-Object NumericUpDown
$numDefaultRPE.Minimum=1; $numDefaultRPE.Maximum=10
$numDefaultRPE.Value = [int]$script:Settings.DefaultRPEForLoad
$numDefaultRPE.Width=120
$setLayout.Controls.Add($numDefaultRPE,1,1)

$btnSaveSettings = New-Object Button
$btnSaveSettings.Text="Save Settings"
$btnSaveSettings.Width=180
$btnSaveSettings.Height=40
$btnSaveSettings.FlatStyle="Flat"
$setLayout.Controls.Add($btnSaveSettings,0,3)
$setLayout.SetColumnSpan($btnSaveSettings,2)

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
    if (-not $entries -or $entries.Count -eq 0) { $gridPR.DataSource=@(); return }

    $longest = $entries | Sort-Object DurationMin -Descending | Select-Object -First 1
    $highest = $entries | Sort-Object { Get-LoadForEntry $_ } -Descending | Select-Object -First 1
    $bestRunDist = $entries | Where-Object { $_.Sport -eq "Running" -and $null -ne $_.DistanceKm } | Sort-Object DistanceKm -Descending | Select-Object -First 1

    $prs = @()
    $prs += [pscustomobject]@{ Name="Longest session"; Value="$($longest.DurationMin) min"; When=$longest.Date; Sport=$longest.Sport }
    $prs += [pscustomobject]@{ Name="Highest load"; Value="$(Get-LoadForEntry $highest)"; When=$highest.Date; Sport=$highest.Sport }
    if ($bestRunDist) {
        $prs += [pscustomobject]@{ Name="Longest run (distance)"; Value="$([math]::Round($bestRunDist.DistanceKm,2)) km"; When=$bestRunDist.Date; Sport=$bestRunDist.Sport }
    }

    $gridPR.DataSource = $prs
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
$btnHist.Add_Click({ ShowPage $pageHistory   $btnHist })
$btnRep.Add_Click( { ShowPage $pageReports   $btnRep  })
$btnLoad.Add_Click({ ShowPage $pageLoad      $btnLoad })
$btnPR.Add_Click(  { ShowPage $pagePRs       $btnPR   })
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
$btnRepPDF.Add_Click({ $btnExportPDF.PerformClick() })

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
