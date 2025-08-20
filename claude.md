# Building Blocks Manager - Development Guide

## Project Overview

**Simple Problem**: Replace manual management of 250+ Word Building Blocks that requires technical expertise most staff don't have.

**Simple Solution**: Staff edit `AT_*.docx` files in folders. Application imports them into Word Building Blocks in template. Done.

## Current Architecture (As Built)

**Language**: C# with Windows Forms
**Dependencies**: Microsoft.Office.Interop.Word (standard Office COM)
**Target**: Windows only - this is not a cross-platform application

### Core Classes
```
MainForm            // Three-tab UI (Query, Directory, Template) + menu system
FileManager         // Scan directories, extract names/categories from file paths
WordManager         // Word COM automation - import/export Building Blocks
BuildingBlockLedger // Track import timestamps with tolerance comparison (KEY CLASS)
Settings            // Persistent settings (paths, ledger directory, logging config)
Logger              // Session-based logging to files
```

### Key Features Implemented
1. **Query Directory**: Scan files, show status with 1-minute tolerance comparison
2. **Directory Tree**: Visual tree with color-coded status (New/Modified/Up-to-date)
3. **Template Viewing**: Browse Building Blocks in current template
4. **Tolerance Logic**: Files show as "modified" only if >1 minute newer than ledger
5. **Configurable Ledger**: File → Ledger Directory to set storage location
6. **Settings Persistence**: All paths and configurations saved between sessions

### Critical Implementation Details

**BuildingBlockLedger Class** - Core timestamp tracking:
- Stores entries in readable format: `Name (45 chars) Category (30 chars) YYYY-MM-DD HH:mm`
- Uses 1-minute tolerance: `timeDifference.TotalMinutes > 1.0` for modified detection
- Configurable directory via Settings.LedgerDirectory
- Handles both default constructor (uses settings) and parameterized (for compatibility)

**File/Category Logic:**
- `Legal\Contracts\AT_Standard.docx` → Category: `InternalAutotext\Legal\Contracts`, Name: `Standard`
- Skip top-level directory in category names
- Spaces become underscores, preserve case
- Only process files starting with `AT_`

## Current UI Structure

**Main Window**: Three-tab interface with menu system
- **Query Tab**: Template/Directory paths, Query buttons, Results text area
- **Directory Tab**: Tree view with color-coded file status
- **Template Tab**: List of Building Blocks in current template

**File Menu**: 
- View Log File, Logging Configuration, Ledger Directory, Rollback

**Key UI Behaviors:**
- Auto-switches to Directory tab after Query Directory
- Auto-queries when switching to empty Directory/Template tabs
- Color coding: Green=New, Blue=Modified, Black=Up-to-date, Red=Invalid

## Development Critical Notes

**Query Directory Process** (MainForm.cs:756):
```csharp
var analysis = ledger.AnalyzeChanges(files);  // Uses tolerance comparison
PopulateDirectoryTree(files, analysis);      // Passes analysis results to tree
```

**Tolerance Comparison** (BuildingBlockLedger.cs:259):
```csharp
var timeDifference = file.LastModified - ledgerEntry.LastModified;
if (timeDifference.TotalMinutes > 1.0) // Only >1 minute = modified
```

**Settings System** (Settings.cs):
- LedgerDirectory: Configurable path for ledger file location
- All settings auto-saved, loaded on startup
- Settings file: `%LOCALAPPDATA%\BuildingBlocksManager\settings.txt`

## Development Workflow

**Working Directory**: `/Users/davidparry/claude_code/Autotext Import_Export Tool`
**Repository**: https://github.com/olPoot-EN/building-blocks-manager

**Development**: Mac M4 Mini with Claude Code (writes C# code only)
**Testing**: Windows laptop (tests actual Word automation) - testing occurs separately on Windows machine

**Git**: 
- **ALWAYS commit and push changes after completing development tasks**
- Use descriptive commit messages explaining what was implemented
- Git commands must be run from within the working directory: `cd "/Users/davidparry/claude_code/Autotext Import_Export Tool"`

**Version Tracking**: 
- **Help → About menu displays version number** for version verification during testing
- **Increment version number after each commit** in GetGitCommitId() method in MainForm.cs
- When testing on Windows, check Help → About to confirm you're running the latest version
- Version number correlates to git commit ID (documented in code comments)
- This helps identify when compiled version is outdated vs. source code changes

**Visual Studio Git Issues**: VS sometimes doesn't sync properly with remote changes
- Use command line git if VS Git is stuck
- Delete and re-clone project folder if VS won't pull latest changes
- **After pulling latest changes, always rebuild before testing** to ensure UI/logic changes take effect

## Common Development Pitfalls

**Ledger System**:
- Query Directory uses `AnalyzeChanges()` with tolerance - don't bypass this
- Directory tree must receive analysis results to show correct status
- Both constructors (default + parameterized) must support configurable ledger directory

**UI State Management**:
- Query Directory auto-switches to Directory tab and passes analysis results
- Don't rely on individual file.IsNew/IsModified - use analysis results instead

**Scoping Issues**:
- Analysis results must be passed from Query → PopulateDirectoryTree → CreateDirectoryNode
- Check all recursive method calls receive needed parameters

**Settings**:
- All configurable paths should use Settings system
- Settings auto-save when changed, auto-load on startup