#requires -version 5.1
<#
GUI for Export-GroupMembers.ps1
Inputs: group names (comma-separated) or CSV with GroupName column
Options: Include nested, output path, CSV fallback
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Export Group Members (DL/SG)'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(720, 420)
$form.MinimumSize = New-Object System.Drawing.Size(720, 420)
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

$lblNames = New-Object System.Windows.Forms.Label
$lblNames.Text = 'Group names (comma-separated):'
$lblNames.Location = New-Object System.Drawing.Point(12, 15)
$lblNames.AutoSize = $true

$txtNames = New-Object System.Windows.Forms.TextBox
$txtNames.Location = New-Object System.Drawing.Point(15, 38)
$txtNames.Size = New-Object System.Drawing.Size(560, 24)
$txtNames.Anchor = 'Top,Left,Right'

$lblCsv = New-Object System.Windows.Forms.Label
$lblCsv.Text = 'Or choose CSV (GroupName column):'
$lblCsv.Location = New-Object System.Drawing.Point(12, 72)
$lblCsv.AutoSize = $true

$txtCsv = New-Object System.Windows.Forms.TextBox
$txtCsv.Location = New-Object System.Drawing.Point(15, 95)
$txtCsv.Size = New-Object System.Drawing.Size(560, 24)
$txtCsv.Anchor = 'Top,Left,Right'

$btnCsv = New-Object System.Windows.Forms.Button
$btnCsv.Text = 'Browse...'
$btnCsv.Location = New-Object System.Drawing.Point(585, 94)
$btnCsv.Size = New-Object System.Drawing.Size(100, 26)
$btnCsv.Anchor = 'Top,Right'

$chkNested = New-Object System.Windows.Forms.CheckBox
$chkNested.Text = 'Include nested members'
$chkNested.Location = New-Object System.Drawing.Point(15, 135)
$chkNested.AutoSize = $true

$lblOut = New-Object System.Windows.Forms.Label
$lblOut.Text = 'Output Excel path (.xlsx):'
$lblOut.Location = New-Object System.Drawing.Point(12, 170)
$lblOut.AutoSize = $true

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Location = New-Object System.Drawing.Point(15, 193)
$txtOut.Size = New-Object System.Drawing.Size(560, 24)
$txtOut.Anchor = 'Top,Left,Right'

$btnOut = New-Object System.Windows.Forms.Button
$btnOut.Text = 'Browse...'
$btnOut.Location = New-Object System.Drawing.Point(585, 192)
$btnOut.Size = New-Object System.Drawing.Size(100, 26)
$btnOut.Anchor = 'Top,Right'

$chkCsv = New-Object System.Windows.Forms.CheckBox
$chkCsv.Text = 'Also write CSV fallback'
$chkCsv.Location = New-Object System.Drawing.Point(15, 230)
$chkCsv.AutoSize = $true

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(15, 265)
$txtLog.Size = New-Object System.Drawing.Size(670, 80)
$txtLog.Multiline = $true
$txtLog.ScrollBars = 'Both'
$txtLog.ReadOnly = $true
$txtLog.Font = New-Object System.Drawing.Font('Consolas', 9)
$txtLog.Anchor = 'Top,Left,Right,Bottom'

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = 'Run'
$btnRun.Location = New-Object System.Drawing.Point(15, 355)
$btnRun.Size = New-Object System.Drawing.Size(100, 28)
$btnRun.Anchor = 'Bottom,Left'

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(125, 355)
$btnClose.Size = New-Object System.Drawing.Size(100, 28)
$btnClose.Anchor = 'Bottom,Left'

$form.Controls.AddRange(@($lblNames,$txtNames,$lblCsv,$txtCsv,$btnCsv,$chkNested,$lblOut,$txtOut,$btnOut,$chkCsv,$txtLog,$btnRun,$btnClose))

function Add-Log([string]$line) {
  if ([string]::IsNullOrWhiteSpace($line)) { return }
  if ($txtLog.InvokeRequired) { $null = $txtLog.Invoke([Action]{ $txtLog.AppendText($line + [Environment]::NewLine) }) } else { $txtLog.AppendText($line + [Environment]::NewLine) }
}

$btnCsv.Add_Click({
  $dlg = New-Object System.Windows.Forms.OpenFileDialog
  $dlg.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
  if ($dlg.ShowDialog() -eq 'OK') { $txtCsv.Text = $dlg.FileName }
  $dlg.Dispose()
})

$btnOut.Add_Click({
  $dlg = New-Object System.Windows.Forms.SaveFileDialog
  $dlg.Filter = 'Excel Workbook (*.xlsx)|*.xlsx|All files (*.*)|*.*'
  if ($dlg.ShowDialog() -eq 'OK') { $txtOut.Text = $dlg.FileName }
  $dlg.Dispose()
})

$btnClose.Add_Click({ $form.Close() })

function Run-Export {
  $names = $txtNames.Text
  $csv   = $txtCsv.Text
  $nested = $chkNested.Checked
  $out   = $txtOut.Text
  $alsoCsv = $chkCsv.Checked

  $scriptPath = Join-Path -Path (Split-Path -Parent $PSCommandPath) -ChildPath 'Export-GroupMembers.ps1'
  if (-not (Test-Path -LiteralPath $scriptPath)) { [void][System.Windows.Forms.MessageBox]::Show("Cannot locate Export-GroupMembers.ps1 at: $scriptPath"); return }

  $argsList = @('-NoLogo','-NoProfile','-ExecutionPolicy','Bypass','-File', ('"{0}"' -f $scriptPath))
  if (-not [string]::IsNullOrWhiteSpace($names)) {
    $split = @($names.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if ($split.Count -gt 0) { $argsList += @('-GroupNames'); $argsList += (@($split | ForEach-Object { '"{0}"' -f $_ })) }
  }
  if (-not [string]::IsNullOrWhiteSpace($csv)) { $argsList += @('-InputCsvPath', ('"{0}"' -f $csv)) }
  if ($nested) { $argsList += '-IncludeNested' }
  if (-not [string]::IsNullOrWhiteSpace($out)) { $argsList += @('-OutputPath', ('"{0}"' -f $out)) }
  if ($alsoCsv) { $argsList += '-CsvFallback' }

  $psi = New-Object System.Diagnostics.ProcessStartInfo
  $psi.FileName = 'powershell.exe'
  $psi.Arguments = ($argsList -join ' ')
  $psi.RedirectStandardOutput = $true
  $psi.RedirectStandardError  = $true
  $psi.UseShellExecute = $false
  $psi.CreateNoWindow = $true

  Add-Log 'Starting export...'
  Add-Log $psi.Arguments
  $btnRun.Enabled = $false

  $proc = New-Object System.Diagnostics.Process
  $proc.StartInfo = $psi
  $null = $proc.Start()
  $proc.BeginOutputReadLine(); $proc.BeginErrorReadLine()
  $proc.add_OutputDataReceived({ param($s,$e) if ($e.Data) { Add-Log $e.Data } })
  $proc.add_ErrorDataReceived({ param($s,$e) if ($e.Data) { Add-Log ('[ERR] ' + $e.Data) } })

  Register-ObjectEvent -InputObject $proc -EventName Exited -Action {
    $code = $proc.ExitCode
    $form.Invoke([Action]{
      Add-Log ("Process exited with code {0}" -f $code)
      $btnRun.Enabled = $true
    }) | Out-Null
    Unregister-Event -SourceIdentifier $event.SourceIdentifier -ErrorAction SilentlyContinue
    $proc.Dispose()
  } | Out-Null
  $proc.EnableRaisingEvents = $true
}

$btnRun.Add_Click({ Run-Export })

[void]$form.ShowDialog()

