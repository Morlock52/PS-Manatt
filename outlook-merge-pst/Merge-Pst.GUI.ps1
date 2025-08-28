#requires -version 5.1
<#
Simple, polished GUI wrapper for Merge-Pst.ps1

Features:
- Pick multiple source PSTs
- Choose destination: new/existing PST or Default Mailbox
- Scope: Inbox only or All folders
- Options: WhatIf, Detach source stores, Verbose logging
- Monitoring: Memory monitoring, GC cadence, report cadence
- Status bar with progress indicator
- Live output pane with Save Log

Run on Windows with Outlook installed.
#>

if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Error 'This GUI requires Windows PowerShell 5.1 or newer.'
    return
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$form              = New-Object System.Windows.Forms.Form
$form.Text         = 'Merge PSTs — Outlook Helper'
$form.StartPosition= 'CenterScreen'
$form.Size         = New-Object System.Drawing.Size(860, 640)
$form.MinimumSize  = New-Object System.Drawing.Size(860, 640)
$form.Font         = New-Object System.Drawing.Font('Segoe UI', 9)

# Menu
$menu = New-Object System.Windows.Forms.MenuStrip
$mFile = New-Object System.Windows.Forms.ToolStripMenuItem -Property @{ Text = 'File' }
$miSaveLog = New-Object System.Windows.Forms.ToolStripMenuItem -Property @{ Text = 'Save Log…' }
$miExit    = New-Object System.Windows.Forms.ToolStripMenuItem -Property @{ Text = 'Exit' }
$mFile.DropDownItems.AddRange(@($miSaveLog,$miExit))
$mHelp = New-Object System.Windows.Forms.ToolStripMenuItem -Property @{ Text = 'Help' }
$miAbout = New-Object System.Windows.Forms.ToolStripMenuItem -Property @{ Text = 'About…' }
$mHelp.DropDownItems.Add($miAbout)
$menu.Items.AddRange(@($mFile,$mHelp))
$form.MainMenuStrip = $menu
$form.Controls.Add($menu)

# Source PSTs
$lblSources = New-Object System.Windows.Forms.Label
$lblSources.Text = 'Source PST files:'
$lblSources.Location = New-Object System.Drawing.Point(12, 35)
$lblSources.AutoSize = $true

$lstSources = New-Object System.Windows.Forms.ListBox
$lstSources.Location = New-Object System.Drawing.Point(15, 58)
$lstSources.Size     = New-Object System.Drawing.Size(650, 140)
$lstSources.SelectionMode = 'MultiExtended'
$lstSources.Anchor = 'Top,Left,Right'

$btnAdd = New-Object System.Windows.Forms.Button
$btnAdd.Text = 'Add...'
$btnAdd.Location = New-Object System.Drawing.Point(680, 58)
$btnAdd.Size     = New-Object System.Drawing.Size(90, 28)
$btnAdd.Anchor   = 'Top,Right'

$btnRemove = New-Object System.Windows.Forms.Button
$btnRemove.Text = 'Remove'
$btnRemove.Location = New-Object System.Drawing.Point(680, 92)
$btnRemove.Size     = New-Object System.Drawing.Size(90, 28)
$btnRemove.Anchor   = 'Top,Right'

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = 'Clear'
$btnClear.Location = New-Object System.Drawing.Point(680, 126)
$btnClear.Size     = New-Object System.Drawing.Size(90, 28)
$btnClear.Anchor   = 'Top,Right'

# Destination
$grpDest = New-Object System.Windows.Forms.GroupBox
$grpDest.Text = 'Destination'
$grpDest.Location = New-Object System.Drawing.Point(15, 210)
$grpDest.Size = New-Object System.Drawing.Size(820, 95)
$grpDest.Anchor = 'Top,Left,Right'

$rbPst = New-Object System.Windows.Forms.RadioButton
$rbPst.Text = 'PST file:'
$rbPst.Location = New-Object System.Drawing.Point(15, 25)
$rbPst.AutoSize = $true
$rbPst.Checked = $true

$txtDest = New-Object System.Windows.Forms.TextBox
$txtDest.Location = New-Object System.Drawing.Point(90, 23)
$txtDest.Size     = New-Object System.Drawing.Size(620, 24)
$txtDest.Anchor   = 'Top,Left,Right'

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = 'Browse...'
$btnBrowse.Location = New-Object System.Drawing.Point(720, 22)
$btnBrowse.Size     = New-Object System.Drawing.Size(80, 26)
$btnBrowse.Anchor   = 'Top,Right'

$rbMailbox = New-Object System.Windows.Forms.RadioButton
$rbMailbox.Text = 'Use Default Mailbox'
$rbMailbox.Location = New-Object System.Drawing.Point(15, 55)
$rbMailbox.AutoSize = $true

$grpDest.Controls.AddRange(@($rbPst, $txtDest, $btnBrowse, $rbMailbox))

# Options
$grpOptions = New-Object System.Windows.Forms.GroupBox
$grpOptions.Text = 'Options, Monitoring & Logging'
$grpOptions.Location = New-Object System.Drawing.Point(15, 315)
$grpOptions.Size = New-Object System.Drawing.Size(820, 140)
$grpOptions.Anchor = 'Top,Left,Right'

$lblScope = New-Object System.Windows.Forms.Label
$lblScope.Text = 'Scope:'
$lblScope.Location = New-Object System.Drawing.Point(15, 30)
$lblScope.AutoSize = $true

$cmbScope = New-Object System.Windows.Forms.ComboBox
$cmbScope.DropDownStyle = 'DropDownList'
$cmbScope.Items.AddRange(@('Inbox','AllFolders'))
$cmbScope.SelectedIndex = 0
$cmbScope.Location = New-Object System.Drawing.Point(65, 27)
$cmbScope.Size = New-Object System.Drawing.Size(120, 24)

$chkWhatIf = New-Object System.Windows.Forms.CheckBox
$chkWhatIf.Text = 'WhatIf (preview only)'
$chkWhatIf.Location = New-Object System.Drawing.Point(210, 28)
$chkWhatIf.AutoSize = $true

$chkDetach = New-Object System.Windows.Forms.CheckBox
$chkDetach.Text = 'Detach source PSTs after merge'
$chkDetach.Location = New-Object System.Drawing.Point(380, 28)
$chkDetach.AutoSize = $true

$chkVerbose = New-Object System.Windows.Forms.CheckBox
$chkVerbose.Text = 'Verbose logging'
$chkVerbose.Location = New-Object System.Drawing.Point(610, 28)
$chkVerbose.AutoSize = $true

# Monitoring
$chkMon = New-Object System.Windows.Forms.CheckBox
$chkMon.Text = 'Monitor memory'
$chkMon.Location = New-Object System.Drawing.Point(15, 62)
$chkMon.AutoSize = $true

$chkDedup = New-Object System.Windows.Forms.CheckBox
$chkDedup.Text = 'Skip duplicates'
$chkDedup.Location = New-Object System.Drawing.Point(130, 62)
$chkDedup.AutoSize = $true

$lblGc = New-Object System.Windows.Forms.Label
$lblGc.Text = 'GC every:'
$lblGc.Location = New-Object System.Drawing.Point(260, 62)
$lblGc.AutoSize = $true

$numGc = New-Object System.Windows.Forms.NumericUpDown
$numGc.Minimum = 0
$numGc.Maximum = 100000
$numGc.Value   = 250
$numGc.Location= New-Object System.Drawing.Point(320, 60)
$numGc.Size    = New-Object System.Drawing.Size(70, 24)

$lblItems1 = New-Object System.Windows.Forms.Label
$lblItems1.Text = 'items'
$lblItems1.Location = New-Object System.Drawing.Point(395, 62)
$lblItems1.AutoSize = $true

$lblRep = New-Object System.Windows.Forms.Label
$lblRep.Text = 'Report every:'
$lblRep.Location = New-Object System.Drawing.Point(450, 62)
$lblRep.AutoSize = $true

$numRep = New-Object System.Windows.Forms.NumericUpDown
$numRep.Minimum = 0
$numRep.Maximum = 100000
$numRep.Value   = 200
$numRep.Location= New-Object System.Drawing.Point(535, 60)
$numRep.Size    = New-Object System.Drawing.Size(70, 24)

$lblItems2 = New-Object System.Windows.Forms.Label
$lblItems2.Text = 'items'
$lblItems2.Location = New-Object System.Drawing.Point(610, 62)
$lblItems2.AutoSize = $true

$chkSummary = New-Object System.Windows.Forms.CheckBox
$chkSummary.Text = 'Show summary dialog'
$chkSummary.Location = New-Object System.Drawing.Point(680, 62)
$chkSummary.AutoSize = $true

# Logging controls
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = 'Log file:'
$lblLog.Location = New-Object System.Drawing.Point(15, 100)
$lblLog.AutoSize = $true

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(70, 98)
$txtLog.Size     = New-Object System.Drawing.Size(620, 24)
$txtLog.Anchor   = 'Top,Left,Right'

$btnLog = New-Object System.Windows.Forms.Button
$btnLog.Text = 'Browse...'
$btnLog.Location = New-Object System.Drawing.Point(700, 97)
$btnLog.Size     = New-Object System.Drawing.Size(90, 26)
$btnLog.Anchor   = 'Top,Right'

$chkAppend = New-Object System.Windows.Forms.CheckBox
$chkAppend.Text = 'Append'
$chkAppend.Location = New-Object System.Drawing.Point(600, 124)
$chkAppend.AutoSize = $true
$chkAppend.Visible = $false # keep layout tidy; append will be used if set

$grpOptions.Controls.AddRange(@($lblScope, $cmbScope, $chkWhatIf, $chkDetach, $chkVerbose, $chkMon, $chkDedup, $lblGc, $numGc, $lblItems1, $lblRep, $numRep, $lblItems2, $chkSummary, $lblLog, $txtLog, $btnLog))

# Output
$txtOutput = New-Object System.Windows.Forms.TextBox
$txtOutput.Location = New-Object System.Drawing.Point(15, 465)
$txtOutput.Size     = New-Object System.Drawing.Size(820, 140)
$txtOutput.Multiline = $true
$txtOutput.ScrollBars = 'Both'
$txtOutput.ReadOnly = $true
$txtOutput.Font = New-Object System.Drawing.Font('Consolas', 9)
$txtOutput.Anchor   = 'Top,Left,Right,Bottom'

# Run/Close
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = 'Run Merge'
$btnRun.Location = New-Object System.Drawing.Point(15, 615)
$btnRun.Size     = New-Object System.Drawing.Size(110, 30)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(135, 615)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)

$btnSaveLog = New-Object System.Windows.Forms.Button
$btnSaveLog.Text = 'Save Log…'
$btnSaveLog.Location = New-Object System.Drawing.Point(255, 615)
$btnSaveLog.Size     = New-Object System.Drawing.Size(110, 30)

# Status bar
$status = New-Object System.Windows.Forms.StatusStrip
$stLabel = New-Object System.Windows.Forms.ToolStripStatusLabel -Property @{ Text = 'Ready' }
$stProg  = New-Object System.Windows.Forms.ToolStripProgressBar -Property @{ Style = 'Blocks'; Visible=$false; Value=0; Minimum=0; Maximum=100 }
$status.Items.AddRange(@($stLabel,$stProg))

$form.Controls.AddRange(@($lblSources, $lstSources, $btnAdd, $btnRemove, $btnClear, $grpDest, $grpOptions, $txtOutput, $btnRun, $btnClose, $btnSaveLog, $status))

# Helpers
function Add-Log([string]$line) {
    if ([string]::IsNullOrWhiteSpace($line)) { return }
    if ($txtOutput.InvokeRequired) {
        $null = $txtOutput.Invoke([Action]{ $txtOutput.AppendText($line + [Environment]::NewLine) })
    } else {
        $txtOutput.AppendText($line + [Environment]::NewLine)
    }
}

function Set-ControlsEnabled([bool]$enabled) {
    foreach ($ctl in @($btnAdd,$btnRemove,$btnClear,$rbPst,$txtDest,$btnBrowse,$rbMailbox,$cmbScope,$chkWhatIf,$chkDetach,$chkVerbose,$chkMon,$numGc,$numRep,$btnRun,$btnSaveLog,$miSaveLog)) {
        $ctl.Enabled = $enabled
    }
}

# Events
$btnAdd.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Multiselect = $true
    $dlg.Filter = 'Outlook PST (*.pst)|*.pst|All files (*.*)|*.*'
    $dlg.Title  = 'Select PST files to merge'
    if ($dlg.ShowDialog() -eq 'OK') {
        foreach ($p in $dlg.FileNames) {
            if (-not $lstSources.Items.Contains($p)) { [void]$lstSources.Items.Add($p) }
        }
    }
    $dlg.Dispose()
})

$btnRemove.Add_Click({
    $toRemove = @()
    foreach ($i in $lstSources.SelectedItems) { $toRemove += $i }
    foreach ($i in $toRemove) { $lstSources.Items.Remove($i) }
})

$btnClear.Add_Click({ $lstSources.Items.Clear() })

$rbPst.Add_CheckedChanged({
    $enable = $rbPst.Checked
    $txtDest.Enabled = $enable
    $btnBrowse.Enabled = $enable
})

$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = 'Outlook PST (*.pst)|*.pst|All files (*.*)|*.*'
    $dlg.Title  = 'Choose destination PST file'
    if (-not [string]::IsNullOrWhiteSpace($txtDest.Text)) {
        try { $dlg.FileName = [System.IO.Path]::GetFileName($txtDest.Text) } catch {}
        try { $dlg.InitialDirectory = [System.IO.Path]::GetDirectoryName($txtDest.Text) } catch {}
    }
    if ($dlg.ShowDialog() -eq 'OK') { $txtDest.Text = $dlg.FileName }
    $dlg.Dispose()
})

$btnLog.Add_Click({
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = 'Text files (*.txt)|*.txt|All files (*.*)|*.*'
    $dlg.Title  = 'Choose log file'
    if (-not [string]::IsNullOrWhiteSpace($txtLog.Text)) {
        try { $dlg.FileName = [System.IO.Path]::GetFileName($txtLog.Text) } catch {}
        try { $dlg.InitialDirectory = [System.IO.Path]::GetDirectoryName($txtLog.Text) } catch {}
    }
    if ($dlg.ShowDialog() -eq 'OK') { $txtLog.Text = $dlg.FileName }
    $dlg.Dispose()
})

$btnClose.Add_Click({ $form.Close() })

$miExit.Add_Click({ $form.Close() })

$miAbout.Add_Click({
    [void][System.Windows.Forms.MessageBox]::Show("Merge PSTs — Outlook Helper`n`nA simple Outlook automation wrapper to merge PST files and fix item types.`n`n© You")
})

$btnSaveLog.Add_Click({
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = 'Text files (*.txt)|*.txt|All files (*.*)|*.*'
    $sfd.Title  = 'Save Log'
    $sfd.FileName = 'merge-log.txt'
    if ($sfd.ShowDialog() -eq 'OK') {
        try { [IO.File]::WriteAllText($sfd.FileName, $txtOutput.Text) } catch { [void][System.Windows.Forms.MessageBox]::Show("Failed to save log: $($_.Exception.Message)") }
    }
    $sfd.Dispose()
})

function Start-Merge() {
    $sources = @()
    foreach ($i in $lstSources.Items) { $sources += $i }
    if ($sources.Count -eq 0) { [void][System.Windows.Forms.MessageBox]::Show('Add at least one source PST.'); return }

    $useMailbox = $rbMailbox.Checked
    $destPath = $txtDest.Text
    if (-not $useMailbox -and [string]::IsNullOrWhiteSpace($destPath)) {
        [void][System.Windows.Forms.MessageBox]::Show('Choose a destination PST path or select "Use Default Mailbox".')
        return
    }

    $scope = $cmbScope.SelectedItem
    $whatIf = $chkWhatIf.Checked
    $detach = $chkDetach.Checked
    $verbose= $chkVerbose.Checked
    $mon    = $chkMon.Checked
    $dedup  = $chkDedup.Checked
    $summary= $chkSummary.Checked
    $gcEvery= [int]$numGc.Value
    $repEvery=[int]$numRep.Value
    $logPath= $txtLog.Text

    $scriptPath = Join-Path -Path (Split-Path -Parent $PSCommandPath) -ChildPath 'Merge-Pst.ps1'
    if (-not (Test-Path -LiteralPath $scriptPath)) {
        [void][System.Windows.Forms.MessageBox]::Show("Cannot locate Merge-Pst.ps1 at: $scriptPath")
        return
    }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'powershell.exe'
    $psi.Arguments = (
        @(
            '-NoLogo','-NoProfile','-ExecutionPolicy','Bypass','-File', ('"{0}"' -f $scriptPath),
            '-Scope', $scope
        ) +
        $(if ($useMailbox) { @('-UseDefaultMailbox') } else { @('-DestinationPstPath', ('"{0}"' -f $destPath)) }) +
        $(if ($whatIf) { @('-WhatIf') } else { @() }) +
        $(if ($detach) { @('-DetachSourceStoresAfter') } else { @() }) +
        $(if ($verbose) { @('-VerboseLogging') } else { @() }) +
        $(if ($dedup) { @('-SkipDuplicates') } else { @() }) +
        $(if ($mon) { @('-MonitorMemory') } else { @() }) +
        $(if ($gcEvery -gt 0) { @('-GcEvery', $gcEvery) } else { @() }) +
        $(if ($repEvery -gt 0) { @('-ReportEvery', $repEvery) } else { @() }) +
        $(if (-not [string]::IsNullOrWhiteSpace($logPath)) { @('-LogPath', ('"{0}"' -f $logPath)) } else { @() }) +
        $(if ($chkAppend.Checked) { @('-AppendLog') } else { @() }) +
        $(if ($summary) { @('-ShowSummaryDialog') } else { @() }) +
        @('-SourcePstPaths') + (@($sources | ForEach-Object { '"{0}"' -f $_ }))
    ) -join ' '
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true

    Add-Log "Starting merge..."
    Add-Log $psi.Arguments
    Set-ControlsEnabled $false
    $btnRun.Text = 'Running...'
    $stLabel.Text = 'Running merge…'
    $stProg.Style = 'Marquee'
    $stProg.MarqueeAnimationSpeed = 30
    $stProg.Visible = $true

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    $null = $proc.Start()

    $proc.BeginOutputReadLine()
    $proc.BeginErrorReadLine()

    $proc.add_OutputDataReceived({ param($s,$e) if ($e.Data) { Add-Log $e.Data } })
    $proc.add_ErrorDataReceived({ param($s,$e) if ($e.Data) { Add-Log ('[ERR] ' + $e.Data) } })

    Register-ObjectEvent -InputObject $proc -EventName Exited -Action {
        $code = $proc.ExitCode
        $form.Invoke([Action]{
            Add-Log ("Process exited with code {0}" -f $code)
            if ($code -eq 0) { Add-Log 'Merge completed.' } else { Add-Log 'Merge finished with errors.' }
            Set-ControlsEnabled $true
            $btnRun.Text = 'Run Merge'
            $stLabel.Text = 'Ready'
            $stProg.Visible = $false
        }) | Out-Null
        Unregister-Event -SourceIdentifier $event.SourceIdentifier -ErrorAction SilentlyContinue
        $proc.Dispose()
    } | Out-Null

    $proc.EnableRaisingEvents = $true
}

$btnRun.Add_Click({ Start-Merge })

[void]$form.ShowDialog()
