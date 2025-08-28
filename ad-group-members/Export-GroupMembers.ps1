#requires -version 5.1
<#
Script 2: Export AD Group Members (DL/SG) to Excel

Inputs
- -GroupNames: one or more group names (SAM, CN, or DN)
- -InputCsvPath: optional CSV with a 'GroupName' column
Options
- -IncludeNested: expand nested groups and indicate 'InheritedFrom'
- -OutputPath: target .xlsx (default: Documents\GroupMembers-<timestamp>.xlsx)
- -CsvFallback: also write a CSV alongside the Excel file

Requires
- Windows with RSAT ActiveDirectory module (preferred)
- Microsoft Excel installed for .xlsx output (COM automation). If unavailable, use -CsvFallback.
#>

[CmdletBinding()] param(
  [string[]] $GroupNames,
  [string]   $InputCsvPath,
  [switch]   $IncludeNested,
  [string]   $OutputPath,
  [switch]   $CsvFallback
)

function Write-Info { param([string]$m) Write-Host $m -ForegroundColor Green }
function Write-Warn { param([string]$m) Write-Warning $m }
function Write-Err  { param([string]$m) Write-Error $m }

function Release-Com { param($o) try { if ($o -and [Runtime.InteropServices.Marshal]::IsComObject($o)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o) } } catch { } }

# Gather groups
$allNames = @()
if ($GroupNames) { $allNames += $GroupNames | Where-Object { $_ } }
if ($InputCsvPath) {
  if (-not (Test-Path -LiteralPath $InputCsvPath)) { throw "CSV not found: $InputCsvPath" }
  try {
    $rows = Import-Csv -LiteralPath $InputCsvPath
    foreach ($r in $rows) { if ($r.GroupName) { $allNames += $r.GroupName } }
  } catch { throw "Failed to read CSV: $($_.Exception.Message)" }
}
if (-not $allNames -or $allNames.Count -eq 0) { throw 'Provide at least one group via -GroupNames or -InputCsvPath.' }
$allNames = $allNames | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique

# Try load AD module
try { Import-Module ActiveDirectory -ErrorAction Stop } catch { Write-Warn 'ActiveDirectory module not found. This script requires RSAT AD tools.'; throw }

function Resolve-GroupByName($name) {
  # Try by DN, then by SAM/CN via -LDAPFilter
  try { return Get-ADGroup -Identity $name -Properties GroupCategory,GroupScope,Mail -ErrorAction Stop } catch {}
  try { $g = Get-ADGroup -LDAPFilter "(|(samAccountName=$name)(cn=$name))" -Properties GroupCategory,GroupScope,Mail | Select-Object -First 1; if ($g) { return $g } } catch {}
  return $null
}

function Get-CategoryLabel($g) {
  try { if ($g.GroupCategory) { return [string]$g.GroupCategory } } catch {}
  return 'Unknown'
}
function Get-ScopeLabel($g) {
  try { if ($g.GroupScope) { return [string]$g.GroupScope } } catch {}
  return 'Unknown'
}

function Get-MemberInfo($dn) {
  # Identify objectClass, then fetch properties accordingly
  $obj = $null
  try { $obj = Get-ADObject -Identity $dn -Properties objectClass,displayName,mail,proxyAddresses,userPrincipalName,sAMAccountName | Select-Object -First 1 } catch {}
  if (-not $obj) { return $null }
  $cls = ($obj.ObjectClass | Select-Object -Last 1)
  $type = switch ($cls) { 'user' { 'User' } 'group' { 'Group' } 'computer' { 'Computer' } 'contact' { 'Contact' } default { $cls } }
  $enabled = $null
  if ($type -eq 'User') { try { $u = Get-ADUser -Identity $dn -Properties Enabled; $enabled = $u.Enabled } catch {} }
  elseif ($type -eq 'Computer') { try { $c = Get-ADComputer -Identity $dn -Properties Enabled; $enabled = $c.Enabled } catch {} }
  [pscustomobject]@{
    MemberType        = $type
    SamAccountName    = $obj.sAMAccountName
    DisplayName       = $obj.DisplayName
    UserPrincipalName = $obj.userPrincipalName
    Email             = ($obj.mail) ?? ''
    DistinguishedName = $dn
    Enabled           = $enabled
  }
}

function Expand-Members($group, [switch]$IncludeNested) {
  $results = @()
  $queue = New-Object System.Collections.Generic.Queue[object]
  $queue.Enqueue(@{ Group=$group; Parent=$null })
  $seen = New-Object 'System.Collections.Generic.HashSet[string]'
  while ($queue.Count -gt 0) {
    $node = $queue.Dequeue()
    $g = $node.Group
    $parent = $node.Parent
    $inheritedFrom = if ($parent) { $parent.Name } else { 'Direct' }
    $members = @()
    try { $members = Get-ADGroupMember -Identity $g.DistinguishedName -ErrorAction Stop } catch { Write-Warn "Cannot expand members for $($g.Name): $($_.Exception.Message)"; continue }
    foreach ($m in $members) {
      $info = Get-MemberInfo $m.DistinguishedName
      if ($info) {
        $row = [pscustomobject]@{
          GroupName    = $group.Name
          GroupDN      = $group.DistinguishedName
          GroupScope   = Get-ScopeLabel $group
          GroupCategory= Get-CategoryLabel $group
          InheritedFrom= $inheritedFrom
          MemberType   = $info.MemberType
          SamAccountName = $info.SamAccountName
          DisplayName    = $info.DisplayName
          UserPrincipalName = $info.UserPrincipalName
          Email            = $info.Email
          DistinguishedName = $info.DistinguishedName
          Enabled          = $info.Enabled
        }
        $results += $row
      }
      if ($IncludeNested -and $m.objectClass -eq 'group') {
        if (-not $seen.Contains($m.DistinguishedName)) {
          $seen.Add($m.DistinguishedName) | Out-Null
          try {
            $ng = Get-ADGroup -Identity $m.DistinguishedName -Properties GroupCategory,GroupScope
            $queue.Enqueue(@{ Group=$ng; Parent=$group })
          } catch {}
        }
      }
    }
  }
  return $results
}

function Get-DefaultOutputPath {
  $doc = [Environment]::GetFolderPath('MyDocuments')
  $name = 'GroupMembers-{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd-HHmmss')
  return Join-Path $doc $name
}

function Export-ToExcel([System.Collections.IEnumerable]$rows, [string]$path) {
  $excel = $workbook = $sheet = $null
  try {
    $excel = New-Object -ComObject Excel.Application
  } catch {
    Write-Warn 'Excel not available. Falling back to CSV. Install Excel for .xlsx output.'
    return $false
  }
  try {
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Name = 'Members'
    $headers = @('GroupName','GroupDN','GroupScope','GroupCategory','InheritedFrom','MemberType','SamAccountName','DisplayName','UserPrincipalName','Email','DistinguishedName','Enabled')
    for ($i=0; $i -lt $headers.Count; $i++) { $sheet.Cells.Item(1, $i+1).Value2 = $headers[$i] }
    $r = 2
    foreach ($row in $rows) {
      $sheet.Cells.Item($r,1).Value2 = $row.GroupName
      $sheet.Cells.Item($r,2).Value2 = $row.GroupDN
      $sheet.Cells.Item($r,3).Value2 = $row.GroupScope
      $sheet.Cells.Item($r,4).Value2 = $row.GroupCategory
      $sheet.Cells.Item($r,5).Value2 = $row.InheritedFrom
      $sheet.Cells.Item($r,6).Value2 = $row.MemberType
      $sheet.Cells.Item($r,7).Value2 = $row.SamAccountName
      $sheet.Cells.Item($r,8).Value2 = $row.DisplayName
      $sheet.Cells.Item($r,9).Value2 = $row.UserPrincipalName
      $sheet.Cells.Item($r,10).Value2 = $row.Email
      $sheet.Cells.Item($r,11).Value2 = $row.DistinguishedName
      $sheet.Cells.Item($r,12).Value2 = $row.Enabled
      $r++
    }
    # Formatting: header bold, autofilter, freeze top row, autofit
    $headerRange = $sheet.Range('A1', $sheet.Cells.Item(1, $headers.Count))
    $headerRange.Font.Bold = $true
    $headerRange.Interior.ColorIndex = 15 # light gray
    $used = $sheet.UsedRange
    $used.EntireColumn.AutoFit() | Out-Null
    $used.AutoFilter() | Out-Null
    $sheet.Application.ActiveWindow.SplitRow = 1
    $sheet.Application.ActiveWindow.FreezePanes = $true
    $sheet.Application.ActiveWindow.Zoom = 100

    $dir = [IO.Path]::GetDirectoryName($path)
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $workbook.SaveAs($path)
    $workbook.Close($true)
    $excel.Quit()
    return $true
  } finally {
    Release-Com $sheet; Release-Com $workbook; Release-Com $excel
  }
}

# Main
$rowsAll = @()
foreach ($name in $allNames) {
  $g = Resolve-GroupByName $name
  if (-not $g) { Write-Warn "Group not found: $name"; continue }
  Write-Info "Processing group: $($g.Name) [$([string](Get-CategoryLabel $g))/$([string](Get-ScopeLabel $g))]"
  $rows = Expand-Members -group $g -IncludeNested:$IncludeNested
  $rowsAll += $rows
}

if (-not $rowsAll -or $rowsAll.Count -eq 0) { Write-Warn 'No members found.'; return }

if (-not $OutputPath) { $OutputPath = Get-DefaultOutputPath }
Write-Info "Writing Excel: $OutputPath"
$ok = Export-ToExcel -rows $rowsAll -path $OutputPath
if (-not $ok -and $CsvFallback) {
  $csvPath = [IO.Path]::ChangeExtension($OutputPath, '.csv')
  Write-Info "Writing CSV fallback: $csvPath"
  $rowsAll | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
}

Write-Info 'Done.'

