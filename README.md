# Building Blocks Manager

A Windows application for importing and exporting Microsoft Word Building Blocks from directory-based AutoText files.

## Overview

This tool solves the problem of managing 250+ Word Building Blocks that previously required manual maintenance. It enables non-technical staff to update Building Block content by editing Word documents in a shared directory structure, while maintaining the performance benefits of pre-loaded Building Blocks in templates.

## Key Features

### Core Functionality
- **Batch Import**: Scan directories recursively for `AT_*.docx` files and import as Building Blocks
- **Selective Import**: Import individual files with confirmation dialogs
- **Batch Export**: Extract all Building Blocks to recreate directory structure
- **Selective Export**: Choose specific Building Blocks to export via multi-select dialog
- **Query Mode**: Analyze directories without performing import operations

### Enhanced Tracking & Logging
- **File Exchange Logging**: Track all imports/exports with detailed information
- **New File Detection**: Automatic warnings when new files require Automator script updates
- **Missing File Alerts**: Detect files that disappear between scans
- **File Manifest System**: Complete inventory tracking across application runs
- **Comprehensive Logging**: All operations logged to `%USERPROFILE%\AppData\Local\BuildingBlocksManager\Logs\`

### Safety Features
- **Automatic Backups**: Template backed up before each import operation
- **Rollback Capability**: Restore from timestamped backups if operations fail
- **Error Recovery**: Comprehensive error handling with automatic rollbacks
- **COM Cleanup**: Proper disposal of Word objects to prevent memory leaks

### User Experience
- **Settings Persistence**: Remember directories, templates, and user preferences
- **Progress Reporting**: Real-time status updates and progress bars
- **Windows Forms UI**: Clean, functional interface focused on productivity
- **File Validation**: Skip files with invalid characters, provide specific error messages

## Technical Specifications

- **Platform**: Windows 10/11
- **Framework**: .NET Framework 4.8
- **Dependencies**: Microsoft Word (COM Interop)
- **UI**: Windows Forms
- **Deployment**: Standalone executable

## File Processing Rules

### Import Rules
- **File Pattern**: Only processes `.docx` files with `AT_` prefix
- **Naming Convention**: `AT_ContractLanguage.docx` → Building Block: `ContractLanguage`
- **Category Mapping**: Folder structure → Building Block categories
  - Example: `Legal\Contracts\AT_Standard.docx` → Category: `InternalAutotext\Legal\Contracts`
- **Character Handling**: Spaces converted to underscores, invalid characters rejected
- **Directory Depth**: Supports up to 5 levels of nested directories

### Export Rules
- **Hierarchical Export**: Recreates original directory structure from categories
- **Flat Export**: Saves all files to single user-specified folder
- **File Naming**: Building Blocks exported as `AT_[name].docx`
- **Conflict Resolution**: Automatic numbering for duplicate filenames

## Architecture

### Core Classes
- **FileManager**: Directory scanning, file validation, category mapping
- **WordManager**: COM Interop, Building Block CRUD operations
- **ImportTracker**: Timestamp tracking, change detection, file manifests
- **Logger**: Enhanced logging with file presence tracking
- **ImportExportManager**: High-level import/export operations
- **SettingsManager**: User preferences and state persistence

### Data Storage
- **Import Tracking**: `%USERPROFILE%\AppData\Local\BuildingBlocksManager\import_tracking.txt`
- **File Manifest**: `%USERPROFILE%\AppData\Local\BuildingBlocksManager\file_manifest.txt`
- **Settings**: `%USERPROFILE%\AppData\Local\BuildingBlocksManager\settings.json`
- **Logs**: `%USERPROFILE%\AppData\Local\BuildingBlocksManager\Logs\BBM_YYYYMMDD_HHMMSS.log`

## Automator Integration Warning System

When new `AT_*.docx` files are detected, the application displays prominent warnings:

```
⚠️ WARNING: New AutoText Entries Detected

The following new files will create new Building Blocks:
• AT_NewContract.docx → Building Block: "NewContract"
• AT_PolicyUpdate.docx → Building Block: "PolicyUpdate"

IMPORTANT: These new autotext entries will require updating
the Automator configuration to target them programmatically.

Please update your Automator scripts to reference:
• "NewContract"
• "PolicyUpdate"
```

## Development Workflow

This project uses a cross-platform development approach:

- **Development**: Mac M4 Mini with Claude Code
- **Testing**: Windows laptop with Visual Studio/VSCode
- **Deployment**: Windows target environment

### Git Workflow
- Frequent commits from Mac development environment
- Windows machine pulls latest changes for testing
- All Word COM automation tested on actual Windows hardware

## Installation

1. Ensure Microsoft Word is installed on the target machine
2. Download the latest release executable
3. Run the application (no installation required)

## Usage

1. **Select Source Directory**: Browse to folder containing `AT_*.docx` files
2. **Select Template File**: Choose the Word template (.dotm) to update
3. **Choose Options**: Enable flat import/export if desired
4. **Query Directory**: Analyze files before importing (recommended)
5. **Import**: Use batch or selective import as needed
6. **Export**: Extract Building Blocks back to directory structure when needed

## Logging and Monitoring

All operations are comprehensively logged:

- File exchanges (imports/exports)
- New file detections with Automator warnings
- Missing file alerts
- Error conditions and recovery actions
- Performance metrics and processing times

Logs are retained for 30 days and automatically cleaned up.

## Performance

- **Target Capacity**: Handle 500+ documents efficiently
- **Processing Speed**: Typical loads complete within 5 minutes
- **Memory Usage**: Maximum 500MB RAM during operations
- **Directory Query**: Complete within 30 seconds

## Error Handling

- **File Access Errors**: Skip locked files, continue processing
- **Word Automation Failures**: Retry operations, rollback on critical failures
- **Template Corruption**: Automatic restore from backup
- **Network Issues**: Graceful handling of connectivity problems

## Development Notes

This is a **productivity tool for power users**, prioritizing:
1. **Reliability** over aesthetics
2. **Functionality** over UI polish
3. **Error recovery** over edge case prevention
4. **Clear logging** over silent operation

## Contributing

This project was developed using Claude Code for cross-platform development efficiency. When contributing:

- Test all Word automation features on Windows hardware
- Use proper COM object disposal patterns
- Follow the existing error handling patterns
- Update documentation for any new features

## License

Internal tool - see organization policies for usage rights.

---

🤖 Generated with [Claude Code](https://claude.ai/code)