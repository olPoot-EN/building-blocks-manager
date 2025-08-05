# Building Blocks Manager - Development Guide

## Project Overview

This project replaces a cumbersome Word auto text/Building Blocks management system where 250+ Building Blocks require manual maintenance. The current system forces staff to learn Word's complex auto text injection process, which most refuse to do.

**Core Problem Solved:** Enable non-technical staff to update Building Block content by editing Word documents in a shared directory structure, while maintaining the performance benefits of pre-loaded Building Blocks in templates.

## Solution Architecture

**Two-phase approach:**
1. **Content Creation**: Staff edit `AT_[name].docx` files in organized directory structure
2. **Build Process**: External application imports these files into Word Building Blocks in master template

**Key Benefits:**
- Staff work with familiar Word documents
- Content updates via OneDrive sync
- Runtime performance from pre-loaded Building Blocks
- Eliminates manual Building Block management

## Development Workflow

### Cross-Platform Development Setup
This project uses a cross-platform development approach due to tooling constraints:

- **Development Machine**: Mac M4 Mini running Claude Code
- **Testing Machine**: Windows laptop with Visual Studio/VSCode (no Claude Code permissions)
- **Target Platform**: Windows (Word automation via COM Interop)

### Git Workflow Requirements
- **Automatic Push**: After each Claude Code session, automatically commit and push changes to GitHub
- **Commit Strategy**: Use timestamped commits or feature branches as appropriate for iterative development
- **Pull Frequency**: Windows machine will frequently pull latest changes for testing and validation
- **Testing Priority**: All Word automation features must be tested on actual Windows hardware (not VM) due to COM behavior differences

### Cross-Platform Development Considerations
- **Windows Compatibility**: Code must be fully compatible with Windows Word COM automation
- **Path Handling**: Avoid Mac-specific paths, use `Path.Combine()` and Windows-compatible file operations
- **System Integration**: No Mac-specific system calls or file handling patterns
- **Automated Deployment**: Consider GitHub Actions for automated build/packaging to streamline Windows deployment
- **Commit History**: Maintain clean, descriptive commits for easy tracking during cross-platform testing cycles

### Testing Strategy
1. **Development Testing**: Basic logic and file operations tested via Claude Code on Mac
2. **Integration Testing**: Word COM automation must be validated on Windows testing machine
3. **Validation Cycle**: Frequent commits from Mac → pull and test on Windows → feedback loop

## Technical Specifications

### Platform & Language
- **Language**: C# 
- **Framework**: .NET Framework 4.8 or .NET 6+
- **UI Framework**: Windows Forms (barebones, minimal UX focus)
- **Dependencies**: Microsoft Word (COM Interop)
- **Deployment**: Standalone executable

### Core Functionality

#### 1. Import System
- **Batch Import**: Scan directory recursively for `AT_*.docx` files (max 5 levels deep)
- **Selective Import**: User selects single file, check timestamp, confirm if unchanged
- **Hierarchical Import**: Convert folder structure to Building Block categories: `folder1\folder2` → `InternalAutotext\folder1\folder2`
- **Flat Import**: Import all files into single user-specified category (prompt user for category name)
- Strip `AT_` prefix from filename for Building Block name
- Handle spaces as underscores in names and categories
- Track import timestamps to identify changed files only

#### 2. Export System (Reverse Process)
- **Batch Export**: Extract all Building Blocks from template with `InternalAutotext*` categories
- **Selective Export**: Load Building Blocks into multi-select dialog, export chosen items only
- **Hierarchical Export**: Recreate directory structure from category paths
- **Flat Export**: Save all Building Blocks to single user-specified folder
- Save as `AT_[name].docx` files in appropriate location(s)

#### 3. Query System
- Report file modification dates vs. last import dates
- Identify new/modified/up-to-date files without performing import

## Implementation Priorities

### Phase 1: Core Import Engine
1. **File System Operations**
   - Directory scanning with depth limit
   - File timestamp tracking and comparison
   - Robust file access error handling

2. **Word Automation Foundation**
   - COM Interop setup with proper disposal patterns
   - Building Block creation/update methods
   - Template backup and rollback system

3. **Error Recovery**
   - Automatic template backup before operations
   - Rollback capability on failures
   - Comprehensive logging system

### Phase 2: User Interface (Minimal)
- **Simple Windows Form** with basic controls:
  - Directory/template selection (browse buttons)
  - Structure options: checkboxes for flat import/export
  - Action buttons: Query, Import All, Import Selected File, Export All, Export Selected, Rollback
  - Results text area
  - Progress indicator
- **Dialog Windows**:
  - File browser for selective import (filtered to AT_*.docx)
  - Building Block selection dialog for selective export (multi-select with Select All/None)
- **User Prompts**: When flat options selected, prompt for category name (import) or folder path (export)
- **Confirmation Dialogs**: For selective import when file unchanged since last import
- **No fancy styling** - focus on functionality over aesthetics
- **Barebones UX** - power user tool, not consumer application

### Phase 3: Polish
- Installer creation
- Enhanced logging and error reporting
- Performance optimization

## Critical Implementation Details

### Word COM Interop Best Practices
```csharp
// Always use using statements or explicit disposal
using (var wordApp = new Word.Application())
{
    // Operations here
} // Automatically releases COM objects

// Handle Word automation failures gracefully
try 
{
    // Word operations
}
catch (COMException ex)
{
    // Log and potentially retry or rollback
}
finally
{
    // Ensure Word objects are released
}
```

### File Processing Logic
- **File Filter**: Only process `.docx` files with `AT_` prefix
- **Name Conversion**: `AT_ContractLanguage.docx` → Building Block name: `ContractLanguage`
- **Category Mapping**: `Legal\Contracts\AT_Standard.docx` → Category: `InternalAutotext\Legal\Contracts`
- **Top-Level Directory**: Exclude top-level directory from category structure
- **Invalid Characters**: Skip files with special characters (`/ \ : * ? " < > |`), provide specific error messages

### Error Scenarios to Handle
1. **File Access**: Locked files, permission issues, network connectivity
2. **Word Automation**: Application not responding, template corruption, Building Block creation failures
3. **Resource Management**: Memory leaks from COM objects, zombie Word processes
4. **Data Integrity**: Template backup failures, partial import states

### Logging Requirements
- **Location**: `%USERPROFILE%\AppData\Local\BuildingBlocksManager\Logs\`
- **Format**: Timestamped entries with operation type and details
- **Content**: Success/failure for each file, error details, operation summaries
- **Retention**: 30 days

### Data Persistence
- **Import Tracking**: Simple text file with `filepath|timestamp` pairs
- **Settings**: Last used directories and template paths
- **Location**: `%USERPROFILE%\AppData\Local\BuildingBlocksManager\`

## Development Notes

### Testing Strategy
1. **Create test directory structure** with sample `AT_*.docx` files
2. **Use throwaway templates** - never test against production templates
3. **Test error scenarios**: locked files, invalid characters, missing directories
4. **Verify round-trip consistency**: Import then export should recreate identical structure

### Code Organization
```
BuildingBlocksManager/
├── Core/
│   ├── FileManager.cs        // Directory scanning, file operations
│   ├── WordManager.cs        // COM Interop, Building Block operations
│   ├── ImportTracker.cs      // Timestamp tracking, change detection
│   └── Logger.cs             // Logging functionality
├── UI/
│   └── MainForm.cs           // Minimal Windows Form
└── Program.cs                // Entry point, configuration
```

### Performance Considerations
- **Batch Processing**: Handle large numbers of files efficiently
- **Memory Management**: Dispose Word objects immediately after use
- **Progress Reporting**: Show progress for long operations
- **Target Performance**: Handle 500 documents in under 5 minutes

## Architecture Decisions Made

### Why External Application vs. VBA
- **Reliability**: Better error handling and recovery than VBA
- **Maintenance**: Easier to update logic without touching user templates
- **Independence**: Can run updates without opening Word documents
- **Deployment**: Single executable easier to distribute than VBA solutions

### Why Directory Structure vs. Database
- **User Familiarity**: Staff already understand file/folder organization
- **Ease of Editing**: Direct Word document editing vs. database forms
- **OneDrive Sync**: Automatic distribution of content updates
- **Formatting Preservation**: Full Word formatting maintained in source documents

### Why Building Blocks vs. Runtime Document Merging
- **Performance**: Pre-loaded Building Blocks much faster than opening 250+ documents at runtime
- **User Experience**: Building Blocks available for manual report composition
- **Reliability**: Eliminates network/file access issues during report generation

## Future Enhancement Opportunities
- Scheduled automatic imports
- Multiple template support
- Change history tracking
- SharePoint integration
- Email notifications for import completion

## Development Philosophy
This is a **productivity tool for power users**, not a consumer application. Prioritize:
1. **Reliability** over aesthetics
2. **Functionality** over UI polish
3. **Error recovery** over edge case prevention
4. **Clear logging** over silent operation

The goal is to eliminate a painful manual process, not to win UI design awards.