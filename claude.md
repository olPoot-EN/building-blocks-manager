# Building Blocks Manager - Development Guide

## Project Overview

**Simple Problem**: Replace manual management of 250+ Word Building Blocks that requires technical expertise most staff don't have.

**Simple Solution**: Staff edit `AT_*.docx` files in folders. Application imports them into Word Building Blocks in template. Done.

## Core Requirements

**What it does:**
1. **Import**: Scan folder for `AT_*.docx` files, create Building Blocks in Word template
2. **Export**: Take Building Blocks from template, save as `AT_*.docx` files in folders
3. **Query**: Show which files changed without importing
4. **Selective**: Import single file or export selected Building Blocks

**File/Category Logic:**
- `Legal\Contracts\AT_Standard.docx` → Category: `InternalAutotext\Legal\Contracts`, Name: `Standard`
- Skip top-level directory in category names
- Spaces become underscores, preserve case
- Only process files starting with `AT_`

## Technical Implementation

**Language**: C# with Windows Forms
**Dependencies**: Microsoft.Office.Interop.Word (standard Office COM)
**Target**: Windows only - this is not a cross-platform application

### Core Classes Needed
```
FileManager     // Scan directories, track timestamps
WordManager     // Open Word, manipulate Building Blocks, save template  
ImportTracker   // Simple text file with filepath|timestamp pairs
Logger          // Write to log file
MainForm        // Basic Windows Form UI
```

### Word Automation - Keep It Simple
```csharp
using Word = Microsoft.Office.Interop.Word;

// Open template, add Building Block, save, close
// Use using statements for COM object disposal
// Create backup before operations
// Handle common errors (file locked, Word not responding)
```

**Key Operations:**
- `BuildingBlocks.Add()` - create new Building Block
- `BuildingBlocks[name].Delete()` - remove existing
- Copy formatted content between documents
- Save template

## UI Requirements - Barebones Only

**Main Window Controls:**
- Template file path + Browse button
- Directory path + Browse button  
- Checkboxes: "Ignore folder/category structure for: ☐ Import ☐ Export"
- Buttons: Query Directory, Import All, Import Selected File, Export All, Export Selected, Rollback
- Results text area
- Progress bar

**Additional Dialogs:**
- File browser for single file import
- Multi-select list for selective export (with Select All/None)
- Input prompts for flat import category name and flat export folder

**No fancy UI** - basic Windows Forms controls, gray background, standard buttons. This is a utility tool.

## Development Workflow

**Development**: Mac M4 Mini with Claude Code (writes C# code only)
**Testing**: Windows laptop (tests actual Word automation)
**Git**: 
- Repository: https://github.com/olPoot-EN/building-blocks-manager
- **ALWAYS commit and push changes after completing development tasks**
- Commit after each Claude Code session, pull and test on Windows
- Use descriptive commit messages explaining what was implemented

**Important**: Code must work on Windows with Word - don't worry about Mac compatibility. The Mac just writes the C# code.

## Features Summary

### Flat Structure Options
- **Flat Import**: All files → single category (user specifies name)
- **Flat Export**: All Building Blocks → single folder (user specifies path)

### Selective Operations  
- **Import Selected**: User picks one `AT_*.docx` file, import just that
- **Export Selected**: User picks Building Blocks from list, export just those

### Standard Operations
- **Import All**: Process all changed files in directory
- **Export All**: Export all Building Blocks to directory structure
- **Query**: Show file status without importing
- **Rollback**: Restore template from backup

## Error Handling
- Create template backup before operations
- Skip locked files, continue processing others
- Show clear error messages for invalid filenames
- Log everything to text file
- Rollback capability if operations fail

## Implementation Priority
1. **Core Word automation** - get Building Block creation/export working
2. **File operations** - directory scanning, timestamp tracking
3. **Basic UI** - Windows Form with essential buttons
4. **Error handling** - backups, logging, rollback
5. **Selective features** - single file operations

**Keep it simple.** This is straightforward Windows automation - scan files, manipulate Word Building Blocks, done. No exotic libraries or complex architectures needed.