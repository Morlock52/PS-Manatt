# Requires: Outlook for Windows (COM automation) and PowerShell 5+ or PowerShell 7 with COM
# Purpose:  Merge multiple PSTs into one destination (PST or default mailbox)
#           and sort/move items into the correct default folders (Inbox, Calendar, Contacts, etc.).
#
# Notes:
# - Run on a Windows machine with Outlook installed. Close other Outlook dialogs first.
# - It is safe to run while Outlook is open; the script will attach to the running instance.
# - Use -WhatIf first to preview moves.
# - If you want to move only items that were incorrectly filed in Inbox, use -Scope Inbox (default).

[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
param(
    [Parameter(Mandatory=$true, Position=0)]
    [string[]] $SourcePstPaths,

    [Parameter(ParameterSetName='ToPst')]
    [string] $DestinationPstPath,

    [Parameter(ParameterSetName='ToMailbox')]
    [switch] $UseDefaultMailbox,

    [ValidateSet('Inbox','AllFolders')]
    [string] $Scope = 'Inbox',

    [switch] $DetachSourceStoresAfter,

    [switch] $VerboseLogging,

    # COM/Memory safety options
    [int] $GcEvery = 250,
    [switch] $MonitorMemory,
    [int] $ReportEvery = 200,

    # De-duplication and reporting
    [switch] $SkipDuplicates,
    [switch] $ShowSummaryDialog,

    # Logging
    [string] $LogPath,
    [switch] $AppendLog,

    # Progress reporting to stdout (for GUI)
    [int] $ProgressEvery = 100
)

${script:LogFile} = $null

function Initialize-Logger {
    if ([string]::IsNullOrWhiteSpace($LogPath)) { return }
    $full = [System.IO.Path]::GetFullPath($LogPath)
    $dir = [System.IO.Path]::GetDirectoryName($full)
    if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    if (-not $AppendLog -and (Test-Path -LiteralPath $full)) {
        Remove-Item -LiteralPath $full -Force -ErrorAction SilentlyContinue
    }
    $script:LogFile = $full
    $header = "# Merge-PST log — {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff')
    Add-Content -Path $script:LogFile -Value $header
}

function Write-LogLine {
    param([string]$Level, [string]$Message)
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
    $line = "[$ts][$Level] $Message"
    if ($script:LogFile) { Add-Content -Path $script:LogFile -Value $line }
    switch ($Level) {
        'INFO'    { Write-Host $line -ForegroundColor Green }
        'WARN'    { Write-Warning $Message }
        'ERROR'   { Write-Error $Message }
        'PERF'    { Write-Host $line -ForegroundColor DarkGray }
        default   { if ($VerboseLogging) { Write-Host $line -ForegroundColor Cyan } }
    }
}

function Write-Info  { param([string]$m) Write-LogLine -Level 'INFO'  -Message $m }
function Write-Debug { param([string]$m) Write-LogLine -Level 'DEBUG' -Message $m }
function Write-Warn  { param([string]$m) Write-LogLine -Level 'WARN'  -Message $m }
function Write-ErrL  { param([string]$m) Write-LogLine -Level 'ERROR' -Message $m }

function Write-Log {
    param([string]$Message)
    if ($VerboseLogging) { Write-Debug $Message }
}

function Release-Com {
    param([object]$Object)
    try {
        if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
        }
    } catch { }
}

function Write-Perf {
    if (-not $MonitorMemory) { return }
    try {
        $ps = Get-Process -Id $PID -ErrorAction Stop
        $ol = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | Select-Object -First 1
        $psPriv = '{0:N1}' -f ($ps.PrivateMemorySize64/1MB)
        $psWS   = '{0:N1}' -f ($ps.WorkingSet64/1MB)
        $olWS   = if ($ol) { '{0:N1}' -f ($ol.WorkingSet64/1MB) } else { 'n/a' }
        Write-LogLine -Level 'PERF' -Message ("PS PrivMB={0} WS={1} | Outlook WS={2}" -f $psPriv,$psWS,$olWS)
    } catch { }
}

$script:DestFolderCache = @{}
$script:DedupeCache = @{}
$script:Summary = @{ Moved = 0; SkippedDuplicate = 0; ByType = @{ Mail=0; Appointment=0; Contact=0; Task=0; Note=0; Journal=0; Other=0 } }

function Get-Outlook {
    try {
        $app = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
    } catch { }
    if (-not $app) {
        $app = New-Object -ComObject Outlook.Application
    }
    return $app
}

function Ensure-LoggedOn($namespace) {
    # Attempt silent logon to the default profile
    try {
        $namespace.Logon($null, $null, $true, $false) | Out-Null
    } catch {
        # Fallback to UI if needed
        $namespace.Logon($null, $null, $true, $true) | Out-Null
    }
}

function Open-Store([__ComObject]$namespace, [string]$pstPath, [switch]$CreateIfMissing) {
    $full = [System.IO.Path]::GetFullPath($pstPath)
    $exists = Test-Path -LiteralPath $full
    if (-not $exists -and -not $CreateIfMissing) {
        throw "PST not found: $full"
    }
    try {
        # 3 = olStoreUnicode
        if ($exists) {
            $namespace.AddStoreEx($full, 3)
        } else {
            # Create an empty PST
            $namespace.AddStoreEx($full, 3)
        }
    } catch {
        # Older Outlook: try AddStore as a fallback
        $namespace.AddStore($full)
    }
    # Resolve to the Store object by FilePath using indexed enumeration to avoid stray RCWs
    $stores = $namespace.Stores
    try {
        $count = $stores.Count
    } catch { $count = 0 }
    $match = $null
    for ($i=1; $i -le $count; $i++) {
        $s = $null
        try { $s = $stores.Item($i) } catch { continue }
        try {
            $path = ($s.FilePath -replace '\\','\')
            if ($path.ToLowerInvariant() -eq $full.ToLowerInvariant()) { $match = $s; break }
        } catch { }
        if ($match -eq $null) { Release-Com $s }
    }
    Release-Com $stores
    if (-not $match) { throw "Unable to open/locate store for $full" }
    return $match
}

function Get-DestinationStore([__ComObject]$namespace) {
    switch ($PSCmdlet.ParameterSetName) {
        'ToPst' {
            if (-not $DestinationPstPath) { throw 'DestinationPstPath is required in ToPst mode.' }
    Write-Info "Opening destination PST: $DestinationPstPath"
            return Open-Store -namespace $namespace -pstPath $DestinationPstPath -CreateIfMissing
        }
        'ToMailbox' {
            # Default delivery store (primary mailbox)
            $defaultStore = $namespace.DefaultStore
            if (-not $defaultStore) { throw 'No default mailbox store detected.' }
            Write-Info "Using default mailbox as destination: $($defaultStore.DisplayName)"
            return $defaultStore
        }
        default {
            throw 'Specify either -DestinationPstPath or -UseDefaultMailbox.'
        }
    }
}

function Get-DefaultFolder($store, [int]$olDefaultFolderId) {
    # Outlook 2010+ supports Store.GetDefaultFolder
    try {
        return $store.GetDefaultFolder($olDefaultFolderId)
    } catch {
        # Fallback to session-wide GetDefaultFolder (works only for default store)
        return $script:Namespace.GetDefaultFolder($olDefaultFolderId)
    }
}

# Cache destination folders per store to reduce COM churn
function Get-DefaultFolderCached($store, [int]$olDefaultFolderId) {
    $key = "{0}|{1}" -f $store.StoreID, $olDefaultFolderId
    if ($script:DestFolderCache.ContainsKey($key)) { return $script:DestFolderCache[$key] }
    $f = Get-DefaultFolder -store $store -olDefaultFolderId $olDefaultFolderId
    $script:DestFolderCache[$key] = $f
    return $f
}

# Map message classes to default folders
function Get-DestinationFolderForItem($item, $destStore) {
    $mc = ''
    try { $mc = [string]$item.MessageClass } catch { $mc = '' }

    # OlDefaultFolders constants
    $olFolderInbox     = 6
    $olFolderCalendar  = 9
    $olFolderContacts  = 10
    $olFolderJournal   = 11
    $olFolderNotes     = 12
    $olFolderTasks     = 13

    switch -Regex ($mc) {
        '^IPM\.Appointment' { return Get-DefaultFolderCached -store $destStore -olDefaultFolderId $olFolderCalendar }
        '^IPM\.Contact'     { return Get-DefaultFolderCached -store $destStore -olDefaultFolderId $olFolderContacts }
        '^IPM\.Task'        { return Get-DefaultFolderCached -store $destStore -olDefaultFolderId $olFolderTasks }
        '^IPM\.Activity'    { return Get-DefaultFolderCached -store $destStore -olDefaultFolderId $olFolderJournal }
        '^IPM\.StickyNote'  { return Get-DefaultFolderCached -store $destStore -olDefaultFolderId $olFolderNotes }
        '^IPM\.DistList'    { return Get-DefaultFolderCached -store $destStore -olDefaultFolderId $olFolderContacts }
        default              { return Get-DefaultFolderCached -store $destStore -olDefaultFolderId $olFolderInbox } # Mail/Meeting/Other
    }
}

function Get-TypeKeyForMessageClass([string]$mc) {
    if ($mc -match '^IPM\.Appointment') { return 'Appointment' }
    elseif ($mc -match '^IPM\.Contact') { return 'Contact' }
    elseif ($mc -match '^IPM\.Task') { return 'Task' }
    elseif ($mc -match '^IPM\.StickyNote') { return 'Note' }
    elseif ($mc -match '^IPM\.Activity') { return 'Journal' }
    elseif ($mc -match '^IPM\.DistList') { return 'Contact' }
    else { return 'Mail' }
}

function Normalize([string]$s) { if ([string]::IsNullOrEmpty($s)) { return '' } return ($s.Trim().ToLowerInvariant()) }
function RoundToMinuteTicks([DateTime]$dt) { if (-not $dt) { return 0 } $rounded = [DateTime]::SpecifyKind([DateTime]::FromFileTimeUtc($dt.ToFileTimeUtc() - ($dt.ToFileTimeUtc() % ([TimeSpan]::FromMinutes(1).Ticks))), [DateTimeKind]::Utc); return $rounded.Ticks }

function Get-ItemFingerprint($item) {
    $mc = ''
    try { $mc = [string]$item.MessageClass } catch { $mc = '' }
    $type = Get-TypeKeyForMessageClass $mc
    switch ($type) {
        'Appointment' {
            $sub = ''; $start=0; $dur=0; $loc=''
            try { $sub = Normalize $item.Subject } catch {}
            try { $start = RoundToMinuteTicks $item.Start } catch {}
            try { $dur = [int]$item.Duration } catch {}
            try { $loc = Normalize $item.Location } catch {}
            return "apt|$sub|$start|$dur|$loc"
        }
        'Contact' {
            $name=''; $comp=''; $e1=''; $e2=''; $e3=''
            try { $name = Normalize $item.FullName } catch {}
            try { $comp = Normalize $item.CompanyName } catch {}
            try { $e1 = Normalize $item.Email1Address } catch {}
            try { $e2 = Normalize $item.Email2Address } catch {}
            try { $e3 = Normalize $item.Email3Address } catch {}
            return "ctc|$name|$comp|$e1|$e2|$e3"
        }
        'Task' {
            $sub=''; $due=0; $start=0; $pc=0
            try { $sub = Normalize $item.Subject } catch {}
            try { $due = RoundToMinuteTicks $item.DueDate } catch {}
            try { $start = RoundToMinuteTicks $item.StartDate } catch {}
            try { $pc = [int]$item.PercentComplete } catch {}
            return "tsk|$sub|$start|$due|$pc"
        }
        'Note' {
            $sub=''
            try { $sub = Normalize $item.Subject } catch {}
            if (-not $sub) { try { $sub = Normalize ($item.Body.Substring(0, [Math]::Min(50,$item.Body.Length))) } catch {} }
            return "note|$sub"
        }
        'Journal' {
            $sub=''; $start=0
            try { $sub = Normalize $item.Subject } catch {}
            try { $start = RoundToMinuteTicks $item.Start } catch {}
            return "jrn|$sub|$start"
        }
        default { # Mail and other
            $sub=''; $sent=0; $from=''; $size=0
            try { $sub = Normalize $item.Subject } catch {}
            try { $sent = RoundToMinuteTicks $item.SentOn } catch {}
            try { $from = Normalize $item.SenderEmailAddress } catch {}
            try { $size = [int]$item.Size } catch {}
            return "mail|$sub|$sent|$from|$size"
        }
    }
}

function Get-DedupeKeyForFolder($folder) {
    try { return "{0}|{1}" -f $folder.StoreID, $folder.EntryID } catch { }
    try { return $folder.FolderPath } catch { }
    return [guid]::NewGuid().ToString()
}

function Get-DedupeSetForFolder($folder) {
    $key = Get-DedupeKeyForFolder $folder
    if ($script:DedupeCache.ContainsKey($key)) { return $script:DedupeCache[$key] }
    $set = New-Object 'System.Collections.Generic.HashSet[string]'
    # Build from existing items if dedupe is enabled
    if ($SkipDuplicates) {
        $items = $folder.Items
        $cnt = 0
        try { $cnt = [int]$items.Count } catch { $cnt = 0 }
        for ($i=$cnt; $i -ge 1; $i--) {
            $it = $null
            try { $it = $items.Item($i) } catch { continue }
            try {
                $fp = Get-ItemFingerprint $it
                if ($fp) { $null = $set.Add($fp) }
            } catch {}
            Release-Com $it
            if ($i % 500 -eq 0) { [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
        }
        Release-Com $items
    }
    $script:DedupeCache[$key] = $set
    return $set
}

function Process-Folder($srcFolder, $destStore, [string]$srcStoreId) {
    $path = ''
    try { $path = $srcFolder.FolderPath } catch { }
    Write-Log "Processing folder: $path"

    # Move items by always taking the first item to avoid index invalidation
    $items = $srcFolder.Items
    $count = 0
    try { $count = [int]$items.Count } catch { $count = 0 }
    $moved = 0

    for ($i = $count; $i -ge 1; $i--) {
        $item = $null
        try { $item = $items.Item($i) } catch { continue }

        $destFolder = $null
        try { $destFolder = Get-DestinationFolderForItem -item $item -destStore $destStore } catch { $destFolder = $null }

        $desc = '(item)'
        try { if ($item.Subject) { $desc = $item.Subject } } catch { }

        $dstName = '(unknown)'
        try { if ($destFolder) { $dstName = $destFolder.FolderPath } } catch { }

        $typeKey = Get-TypeKeyForMessageClass $mc

        $isDup = $false
        if ($SkipDuplicates -and $destFolder) {
            try {
                $fp = Get-ItemFingerprint $item
                $set = Get-DedupeSetForFolder $destFolder
                if ($fp -and $set.Contains($fp)) { $isDup = $true }
            } catch { $isDup = $false }
        }

        if ($isDup) {
            $script:Summary.SkippedDuplicate++
            $script:Summary.ByType[$typeKey] = ($script:Summary.ByType[$typeKey] + 0)
        } elseif ($PSCmdlet.ShouldProcess($desc, "Move to $dstName")) {
            try {
                [void]$item.Move($destFolder)
                $moved++
                $script:__movedTotal++
                $script:Summary.Moved++
                if (-not $script:Summary.ByType.ContainsKey($typeKey)) { $script:Summary.ByType[$typeKey] = 0 }
                $script:Summary.ByType[$typeKey]++
                if ($SkipDuplicates -and $fp) {
                    try { $null = (Get-DedupeSetForFolder $destFolder).Add($fp) } catch {}
                }
                if ($ProgressEvery -gt 0 -and ($script:__movedTotal % $ProgressEvery -eq 0)) {
                    Write-ProgressTick -Moved $script:__movedTotal
                }
            } catch {
                Write-Warn "Failed to move item '$desc' to '$dstName': $($_.Exception.Message)"
            }
        }

        Release-Com $item

        if ($GcEvery -gt 0 -and ($script:__movedTotal % $GcEvery -eq 0)) {
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
            Write-Perf
        }
        if ($MonitorMemory -and $ReportEvery -gt 0 -and ($script:__movedTotal % $ReportEvery -eq 0)) {
            Write-Perf
        }
    }

    Release-Com $items

    # Recurse into subfolders using EntryIDs to avoid holding COM refs
    try {
        if (-not $srcStoreId) { try { $srcStoreId = $srcFolder.StoreID } catch { $srcStoreId = $null } }
        $subs = $srcFolder.Folders
        $subCount = 0
        try { $subCount = [int]$subs.Count } catch { $subCount = 0 }
        $childEntryIds = @()
        for ($s=1; $s -le $subCount; $s++) {
            $sf = $null
            try { $sf = $subs.Item($s) } catch { continue }
            try { $childEntryIds += $sf.EntryID } catch { }
            Release-Com $sf
        }
        Release-Com $subs
        foreach ($eid in $childEntryIds) {
            $child = $null
            try { $child = $script:Namespace.GetFolderFromID($eid, $srcStoreId) } catch { }
            if ($child) {
                Process-Folder -srcFolder $child -destStore $destStore -srcStoreId $srcStoreId
                Release-Com $child
            }
        }
    } catch { }

    Write-Log "Done folder: $path (moved $moved items)"
}

# --- Main ---

if (-not $UseDefaultMailbox -and [string]::IsNullOrWhiteSpace($DestinationPstPath)) {
    throw 'Specify either -DestinationPstPath or -UseDefaultMailbox.'
}

$SourcePstPaths = $SourcePstPaths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() } | Select-Object -Unique
if (-not $SourcePstPaths -or $SourcePstPaths.Count -eq 0) { throw 'No source PST paths provided.' }

$Outlook = Get-Outlook
$Namespace = $Outlook.GetNameSpace('MAPI')
Ensure-LoggedOn -namespace $Namespace
$script:Namespace = $Namespace
$script:__movedTotal = 0

Initialize-Logger
Write-Info "Merge started"
Write-Debug ("Params: Scope={0} UseMailbox={1} DestPST='{2}' Sources={3} SkipDup={4} Verbose={5}" -f $Scope,$UseDefaultMailbox.IsPresent,$DestinationPstPath,($SourcePstPaths -join ';'),$SkipDuplicates.IsPresent,$VerboseLogging.IsPresent)

# Resolve destination store
$destStore = Get-DestinationStore -namespace $Namespace

# Process each source PST
foreach ($pst in $SourcePstPaths) {
    Write-Info "Opening source PST: $pst"
    $srcStore = Open-Store -namespace $Namespace -pstPath $pst
    $root = $srcStore.GetRootFolder()
    $srcName = $srcStore.DisplayName
    Write-Info "Source store opened: $srcName"

    # Determine starting folders based on scope
    $startFolders = @()
    if ($Scope -eq 'Inbox') {
        try { $startFolders += $srcStore.GetDefaultFolder(6) } catch { Write-Log "Source Inbox not found; skipping." }
    } else {
        $startFolders += $root
    }

    foreach ($sf in $startFolders) { Process-Folder -srcFolder $sf -destStore $destStore -srcStoreId $srcStore.StoreID }

    # Release start folders
    foreach ($sf in $startFolders) { Release-Com $sf }

    if ($DetachSourceStoresAfter -and $PSCmdlet.ShouldProcess($srcName, 'Remove source store from profile')) {
        try { $Namespace.RemoveStore($root) } catch { Write-Warn "Unable to detach source store '$srcName': $($_.Exception.Message)" }
    }

    Release-Com $root
    Release-Com $srcStore
}

# Release cached destination folders
foreach ($kv in $script:DestFolderCache.GetEnumerator()) { Release-Com $kv.Value }

Release-Com $destStore
Release-Com $Namespace
Release-Com $Outlook

Write-Info 'Merge completed. Review your destination folders for correctness.'

# Summary output
Write-LogLine -Level 'INFO' -Message ("Moved: {0} | Skipped (dupe): {1}" -f $script:Summary.Moved, $script:Summary.SkippedDuplicate)
foreach ($k in @('Mail','Appointment','Contact','Task','Note','Journal','Other')) {
    if ($script:Summary.ByType.ContainsKey($k) -and $script:Summary.ByType[$k] -gt 0) {
        Write-LogLine -Level 'INFO' -Message ("  {0}: {1}" -f $k, $script:Summary.ByType[$k])
    }
}

if ($ShowSummaryDialog) {
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $lines = @()
        $lines += "Merge completed."
        $lines += ("Moved: {0}" -f $script:Summary.Moved)
        $lines += ("Skipped (duplicates): {0}" -f $script:Summary.SkippedDuplicate)
        foreach ($k in @('Mail','Appointment','Contact','Task','Note','Journal','Other')) {
            if ($script:Summary.ByType.ContainsKey($k) -and $script:Summary.ByType[$k] -gt 0) { $lines += ("  {0}: {1}" -f $k, $script:Summary.ByType[$k]) }
        }
function Write-ProgressTick {
    param([int]$Moved)
    # Plain line for GUI to parse; also log to file if enabled
    $line = "PROGRESS $Moved"
    Write-Host $line
    if ($script:LogFile) { Add-Content -Path $script:LogFile -Value ("[{0}][INFO] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $line) }
}

        [System.Windows.Forms.MessageBox]::Show([string]::Join([Environment]::NewLine, $lines), 'Merge Summary') | Out-Null
    } catch { }
}
