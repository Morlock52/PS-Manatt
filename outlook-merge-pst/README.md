# Script 1: Merge PSTs — Outlook Helper

This script set consolidates multiple Outlook PST files into a single destination (a PST or your default mailbox) and routes items into the correct default folders based on their type (mail → Inbox, appointments → Calendar, contacts → Contacts, tasks → Tasks, notes → Notes, journal → Journal). It also includes a Windows Forms GUI.

Highlights
- Item-type routing via `MessageClass`
- Optional duplicate skipping per destination folder
- Structured logging to file with timestamps/levels
- Memory-safe COM automation with periodic GC and careful enumeration
- Progress events for GUI parsing

Files
- `Merge-Pst.ps1`: Core engine (COM automation)
- `Merge-Pst.GUI.ps1`: Windows Forms launcher with options and live logs

Usage
- GUI: `./outlook-merge-pst/Merge-Pst.GUI.ps1`
- CLI example:
  `./outlook-merge-pst/Merge-Pst.ps1 -SourcePstPaths 'C:\PSTs\a.pst','C:\PSTs\b.pst' -DestinationPstPath 'C:\PSTs\Merged.pst' -Scope AllFolders -SkipDuplicates -LogPath 'C:\Logs\merge.txt'`

Parameters (core)
- `-SourcePstPaths <string[]>`: One or more PST paths to merge
- Destination: `-DestinationPstPath <string>` or `-UseDefaultMailbox`
- `-Scope Inbox|AllFolders`: Process only Inbox or the entire PST tree
- Safety/scale: `-GcEvery <int>`, `-MonitorMemory`, `-ReportEvery <int>`
- Dedupe/report: `-SkipDuplicates`, `-ShowSummaryDialog`
- Logging: `-LogPath <file>`, `-AppendLog`
- Progress: `-ProgressEvery <int>` emits `PROGRESS N` lines to stdout

How it works (detailed)
1) Initialize and log
   - Parse parameters; initialize logger if `-LogPath` is set
   - Attach to Outlook via COM; `MAPI` namespace `Logon`

2) Destination resolution
   - If `-UseDefaultMailbox`, select the default store
   - Else open/create the destination PST store with `AddStoreEx`

3) For each source PST
   - Open the store; get root folder
   - Pick starting folder(s): Inbox (if `-Scope Inbox`) or Root (if `-Scope AllFolders`)

4) Process-Folder recursion
   - Enumerate `Items` using index (descending) to avoid collection invalidation
   - For each item:
     - Detect type by `MessageClass` and map to target default folder (cached per store)
     - If `-SkipDuplicates`, build/access a per-target-folder fingerprint `HashSet` and skip existing
     - Move the item to the destination folder
     - Periodically emit `PROGRESS N`, collect GC, and optionally log memory
   - Recurse subfolders by capturing `EntryID`s first to avoid holding live COM references

5) Optional cleanup
   - If `-DetachSourceStoresAfter`, remove the source store from the profile

6) Finalization
   - Release COM objects aggressively
   - Print/log a summary and optionally show a summary dialog

Deduplication fingerprints
- Mail: subject + sent-on (rounded to minute) + sender + size
- Appointment: subject + start + duration + location
- Contact: name + company + up to 3 emails
- Task: subject + start + due + percent complete
- Note/Journal: subject (and short body/start)

Mermaid: flow overview
```mermaid
flowchart TD
  A[Start] --> B[Parse params / Init logger]
  B --> C[Get Outlook.Application / MAPI Logon]
  C --> D{Destination}
  D -->|Default mailbox| E[Use DefaultStore]
  D -->|PST path| F[AddStoreEx/Open dest PST]
  E --> G[Loop source PSTs]
  F --> G
  G --> H[Open source store]
  H --> I{Scope}
  I -->|Inbox| J[Start at Source Inbox]
  I -->|AllFolders| K[Start at Source Root]
  J --> L[Process-Folder]
  K --> L
  subgraph Processing
    L --> M[Enumerate Items (desc index)]
    M --> N[Map MessageClass → Dest default folder]
    N --> O{SkipDuplicates?}
    O -->|Yes & dup| P[Skip item]
    O -->|No or not dup| Q[Move item]
    Q --> R[Emit PROGRESS/GC/Perf]
    P --> R
    R --> S[Recurse subfolders by EntryID]
  end
  S --> T{More sources?}
  T -->|Yes| G
  T -->|No| U[Detach sources (optional)]
  U --> V[Release COM / Summary]
  V --> W[End]
```

Notes & limits
- Requires Outlook for Windows; avoid prompts while running
- Use `-WhatIf` first to validate routing
- Progress is count-based; not a total % (GUI can estimate separately)

