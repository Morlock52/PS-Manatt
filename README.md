# PS-manatt — PowerShell Script Collection

PS-manatt is a growing collection of practical PowerShell scripts and helpers. This repository currently includes its first utility focused on Outlook PST consolidation. More scripts will be added over time.

Scripts in this repo:
- Script 1: Merge PSTs — Outlook Helper (`outlook-merge-pst/`)
- Script 2: Export AD Group Members to Excel (`ad-group-members/`)

## Script 1: Merge PSTs — Outlook Helper

PowerShell scripts to merge multiple PST files into a single destination (PST or your default mailbox) and auto-file items into the correct folders (Inbox, Calendar, Contacts, Tasks, Notes, Journal). Includes a Windows Forms GUI and robust handling for large merges.

### Features
- Merge multiple PSTs into one destination PST or your default mailbox
- Auto-sort by item type using MessageClass → default folders
- Optional duplicate skipping per destination folder
- WhatIf dry-run mode
- Memory-safe COM usage with periodic GC and folder recursion by EntryID
- Optional memory monitoring and progress output
- Structured logging to file with timestamps and levels
- GUI with live log, status bar, and options for monitoring/dedupe/logging

### Files
- `outlook-merge-pst/Merge-Pst.ps1`: Core engine
- `outlook-merge-pst/Merge-Pst.GUI.ps1`: Windows Forms GUI wrapper

### Requirements
- Windows with Outlook for Windows installed (COM automation)
- PowerShell 5.1+ (Windows PowerShell) or PowerShell 7+ (COM works via Windows PowerShell; recommended 5.1)

### Quick Start (GUI)
1. Start Windows PowerShell.
2. Run: `./outlook-merge-pst/Merge-Pst.GUI.ps1`
3. Add source PSTs, choose destination (PST file or Default Mailbox).
4. Options:
   - Scope: `Inbox` (fix misplaced items) or `AllFolders` (full merge)
   - WhatIf (preview), Detach sources, Verbose logging
   - Monitor memory, GC/Report cadence
   - Skip duplicates, Show summary dialog
   - Log file path
5. Click Run Merge and watch the live log.

### Quick Start (CLI)
Examples:
- Dry-run to preview:
  `./outlook-merge-pst/Merge-Pst.ps1 -SourcePstPaths 'C:\PSTs\a.pst','C:\PSTs\b.pst' -DestinationPstPath 'C:\PSTs\Merged.pst' -Scope Inbox -WhatIf`
- Full merge into PST with dedupe and logging:
  `./outlook-merge-pst/Merge-Pst.ps1 -SourcePstPaths 'C:\PSTs\a.pst','C:\PSTs\b.pst' -DestinationPstPath 'C:\PSTs\Merged.pst' -Scope AllFolders -SkipDuplicates -LogPath 'C:\Logs\merge.txt'`
- Merge into default mailbox with monitoring:
  `./outlook-merge-pst/Merge-Pst.ps1 -SourcePstPaths 'C:\PSTs\old.pst' -UseDefaultMailbox -Scope AllFolders -MonitorMemory -GcEvery 250 -ReportEvery 200`

Key parameters:
- `-DestinationPstPath` or `-UseDefaultMailbox`
- `-Scope Inbox|AllFolders`
- `-SkipDuplicates`, `-ShowSummaryDialog`
- `-WhatIf`, `-DetachSourceStoresAfter`, `-VerboseLogging`
- `-LogPath <file>`, `-AppendLog`
- `-MonitorMemory`, `-GcEvery <n>`, `-ReportEvery <n>`
- `-ProgressEvery <n>`

### Notes
- Always test with `-WhatIf` first.
- Keep Outlook UI free of prompts while running.
- For extremely large PSTs, consider running in batches and enabling logging.

### Repository Layout
```
outlook-merge-pst/
  Merge-Pst.ps1       # Core merging logic
  Merge-Pst.GUI.ps1   # GUI launcher
  README.md           # Detailed breakdown + flow
ad-group-members/
  Export-GroupMembers.ps1     # Core exporter
  Export-GroupMembers.GUI.ps1 # GUI wrapper
README.md             # This file
```

### Contributing
Issues and PRs welcome. Please avoid including sample PSTs in the repo. Use logs/screenshots instead.
