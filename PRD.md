# Building Blocks Manager - Product Requirements Document

## 1. Product Overview

### 1.1 Purpose
Create a standalone Windows application to manage Word Building Blocks which provides import/export of formatted content from directory-based Word documents, replacing the current cumbersome auto text management system.

### 1.2 Target Users
- Primary: IT administrator managing report templates
- Secondary: Staff who need updated Building Blocks in their templates

### 1.3 Success Criteria
- Reduce Building Block maintenance time from hours to minutes
- Enable non-technical staff to update content by editing Word documents
- Eliminate manual Building Block creation process
- Maintain formatting fidelity during import process

## 2. Functional Requirements

### 2.1 Core Import Functionality

#### 2.1.1 File Processing Rules
- **Target Files**: Only process .docx files with "AT_" prefix
- **Naming Convention**: 
  - Source file: `AT_ContractLanguage.docx`
  - Building Block name: `ContractLanguage`
- **Directory Depth**: Support up to 5 levels of nested directories
- **Category Mapping**: Use folder path as Building Block category
  - Example: `Legal\Contracts\AT_Standard.docx` → Category: "Legal\Contracts", Name: "Standard"

#### 2.1.2 Directory Structure to Category Mapping
- **Category Creation**: Convert folder path to Building Block category using backslashes
- **Root Category**: All Building Blocks use "InternalAutotext" as root category
- **Example Mapping**:
  - `Legal\Contracts\AT_Standard.docx` → Category: `InternalAutotext\Legal\Contracts`, Name: `Standard`
  - `HR\AT_Policy.docx` → Category: `InternalAutotext\HR`, Name: `Policy`
  - `AT_General.docx` (root level) → Category: `InternalAutotext`, Name: `General`

#### 2.1.3 Character Handling
- **Spaces in Folders**: Convert to underscores in categories (`Legal Documents` → `Legal_Documents`)
- **Spaces in Filenames**: Convert to underscores in Building Block names
- **Special Characters**: Display notification with specific invalid characters found, skip problematic files
- **Case Sensitivity**: Preserve original case from filenames and folders

#### 2.1.8 Import Process
1. Create timestamped backup of target template
2. Scan directory recursively for AT_ prefixed .docx files (up to 5 levels deep)
3. Compare file modification dates against last import tracking
4. For each changed file:
   - Extract folder path and convert to category structure
   - Remove AT_ prefix from filename for Building Block name
   - Open source document and extract formatted content
   - Create/update Building Block with category `InternalAutotext\[folder_path]`
5. Save template with new Building Blocks
6. Log all operations

### 2.2 Template Management

#### 2.2.1 Target Template
- Single master template file (.dotm format)
- User selects template file via browse dialog
- Application remembers last used template path

#### 2.2.2 Backup System
- Create backup before each import operation
- Backup naming: `[TemplateName]_Backup_YYYYMMDD_HHMMSS.dotm`
- Store backups in same directory as original template

#### 2.2.3 Rollback Capability
- If import fails, automatically restore from backup
- Manual rollback option in case of post-import issues
- Retain last 5 backups, delete older ones

### 2.3 Query/Status Functionality

#### 2.3.1 Directory Analysis
- Scan directory without performing import
- Report file modification dates vs. last import dates
- Identify new files (never imported)
- Identify modified files (changed since last import)
- Count ignored files (non-AT_ prefixed)

#### 2.3.2 Status Report Format
```
Files Ready for Import: 12
- New Files: 8
- Modified Files: 4
- Up-to-date Files: 15
- Ignored Files: 23

Detailed Listing:
Legal\Contracts\AT_Standard.docx - Modified (Last: 2025-01-15, Imported: 2025-01-10)
Legal\Forms\AT_Disclaimer.docx - New (Never imported)
...
```

### 2.4 Export Functionality

#### 2.4.1 Export Operation Types
- **Batch Export**: Export all Building Blocks from template
- **Selective Export**: Export user-selected Building Blocks only

#### 2.4.2 Batch Export Process
1. Open target template file
2. Enumerate all Building Blocks in "InternalAutotext" category and subcategories
3. Determine export structure based on user settings (hierarchical or flat)
4. For each Building Block:
   - **Hierarchical Export**: Parse category structure to determine folder path, create directory structure
   - **Flat Export**: Save all files to single user-specified folder
   - Create new Word document with Building Block content
   - Save as `AT_[BuildingBlockName].docx`
5. Generate export summary

#### 2.4.3 Selective Export Process
1. Open target template file
2. Load all Building Blocks from "InternalAutotext" category into selectable list
3. Display Building Block selection dialog with:
   - Multi-select list showing category and name for each Building Block
   - "Select All" and "Select None" buttons
4. User selects specific Building Blocks to export
5. Export selected Building Blocks according to structure settings
6. Generate export summary

#### 2.4.4 Category to Directory Structure Conversion (Hierarchical Export)
- **Category Parsing**: Convert Building Block categories back to folder paths
- **Top-Level Directory**: Create within user-selected export directory
- **Example Conversion**:
  - Category: `InternalAutotext\Legal\Contracts`, Name: `Standard` → `Legal\Contracts\AT_Standard.docx`
  - Category: `InternalAutotext\HR`, Name: `Policy` → `HR\AT_Policy.docx`
  - Category: `InternalAutotext`, Name: `General` → `AT_General.docx` (root level)
- **Folder Names**: Convert underscores back to spaces in folder names (`Legal_Documents` → `Legal Documents`)

#### 2.4.5 Flat Export Processing
- **User Prompt**: Request target folder for all exported files
- **File Placement**: All Building Blocks exported to single specified folder as `AT_[name].docx`
- **Category Ignored**: Original Building Block categories not reflected in folder structure

#### 2.4.6 File Naming for Export
- Building Block name "Standard_Contract" → `AT_Standard_Contract.docx`
- Preserve underscores in exported filenames (these represent spaces in original names)
- Handle name conflicts by appending numbers: `AT_Standard_Contract_2.docx`

#### 2.4.7 Export Summary
```
Export Completed Successfully

Building Blocks Exported: 45
- Created Directories: 8
- Files Created: 45
- Export Location: C:\Reports\AutoText\
Processing Time: 1.8 seconds

Exported Files:
Legal\Contracts\AT_Standard.docx
Legal\Contracts\AT_Premium.docx
Legal\Forms\AT_Disclaimer.docx
HR\Policies\AT_Vacation.docx
Finance\Reports\AT_Summary.docx
...
```

### 2.5 Import Summary

#### 2.5.1 Success Summary
- Count of Building Blocks imported
- List of imported Building Block names with categories
- Count of files ignored
- Processing time

#### 2.5.2 Example Output
```
Import Completed Successfully

Building Blocks Imported: 12
- InternalAutotext\Legal\Contracts: Standard, Premium, Basic
- InternalAutotext\Legal\Forms: Disclaimer, Terms
- InternalAutotext\HR\Policies: Vacation, Sick_Leave
- InternalAutotext\Finance\Reports: Summary, Detailed

Files Ignored: 23 (notes, documentation, non-AT_ files)
Processing Time: 2.3 seconds
```

#### 2.5.3 Special Character Handling
When files contain invalid characters, display specific guidance:
```
Files Skipped Due to Special Characters: 3

Invalid characters found: / \ : * ? " < > |
Please rename the following files and remove these characters:
- AT_Contract/Template.docx (contains: /)
- AT_Policy*.docx (contains: *)
- AT_Form?.docx (contains: ?)

Valid characters: Letters, numbers, spaces, hyphens, underscores
```

## 3. User Interface Requirements

### 3.1 Main Window
- **Window Title**: "Building Blocks Manager"
- **Size**: 800x600 pixels, resizable
- **Layout**: Vertical arrangement of controls

### 3.2 File Selection Section
- **Source Directory**: 
  - Label: "Source Directory"
  - Text box showing selected path
  - "Browse" button to select folder
- **Target Template**:
  - Label: "Template File"
  - Text box showing selected template path
  - "Browse" button to select .dotm file

### 3.3 Structure Options Section
- **Label**: "Ignore folder/category structure for:"
- **Import Checkbox**: "Import" - when checked, prompts for single category name
- **Export Checkbox**: "Export" - when checked, prompts for single output folder

### 3.4 Action Buttons
- **"Query Directory"**: Run status check without importing
- **"Import All"**: Import all changed files using batch process
- **"Import Selected File"**: Browse and import single file
- **"Export All"**: Export all Building Blocks to directory structure
- **"Export Selected"**: Show Building Block selection dialog, export chosen items
- **"Rollback"**: Restore from most recent backup
- **"Exit"**: Close application

### 3.5 Results Display
- **Multi-line text box**: Display operation results and status
- **Progress bar**: Show progress during operations
- **Status label**: Current operation status

### 3.5 Menu Bar
- **File Menu**:
  - Settings
  - View Log File
  - Exit
- **Help Menu**:
  - About

## 4. Technical Requirements

### 4.1 Platform
- **Target OS**: Windows 10/11
- **Framework**: .NET Framework 4.8 or .NET 6+
- **Dependencies**: Microsoft Word (for Building Block manipulation)

### 4.2 File System Operations
- **Directory Scanning**: Recursive traversal up to 5 levels deep
- **File Monitoring**: Track modification timestamps
- **Backup Management**: Automated backup creation and cleanup

### 4.3 Word Integration
- **COM Interop**: Use Word Object Model for Building Block manipulation
- **Error Handling**: Graceful handling of Word automation failures
- **Resource Management**: Proper disposal of Word objects

### 4.4 Data Persistence
- **Import Tracking**: Store last import timestamps in text file: `%USERPROFILE%\AppData\Local\BuildingBlocksManager\import_tracking.txt`
- **Tracking Format**: `filepath|timestamp` pairs, one per line
- **Settings**: Remember last used directories and template paths in `settings.txt`
- **Example tracking file**:
```
C:\AutoText\Legal\AT_Standard.docx|2025-08-03 14:30:17
C:\AutoText\HR\AT_Policy.docx|2025-08-02 09:15:22
```

## 5. Logging and Error Handling

### 5.1 Log File
- **Location**: `%USERPROFILE%\AppData\Local\BuildingBlocksManager\Logs\`
- **Naming**: `BBM_YYYYMMDD_HHMMSS.log`
- **Retention**: Keep logs for 30 days

### 5.2 Log Content
```
[2025-08-03 14:30:15] INFO: Starting import process
[2025-08-03 14:30:16] INFO: Backup created: Template_Backup_20250803_143016.dotm
[2025-08-03 14:30:17] SUCCESS: Imported Legal\Contracts\Standard from AT_Standard.docx
[2025-08-03 14:30:18] WARNING: Skipped invalid_file.docx (special characters in name)
[2025-08-03 14:30:19] ERROR: Failed to process AT_Problem.docx - File locked
[2025-08-03 14:30:20] INFO: Import completed - 11 successful, 1 failed
```

### 5.3 Error Recovery
- **File Lock Errors**: Skip locked files, continue with others
- **Word Automation Failures**: Attempt retry, then rollback if critical failure
- **Template Corruption**: Automatic restore from backup
- **User Notification**: Clear error messages with suggested actions

## 6. Performance Requirements

### 6.1 Processing Speed
- Handle up to 500 source documents
- Import process should complete within 5 minutes for typical loads
- Directory query should complete within 30 seconds

### 6.2 Memory Usage
- Maximum 500MB RAM usage during operations
- Proper cleanup of Word COM objects to prevent memory leaks

## 7. Security and Data Integrity

### 7.1 File Validation
- Verify .docx file integrity before processing
- Validate template file format before modification
- Check write permissions before starting operations

### 7.2 Backup Protection
- Prevent overwriting of backup files
- Verify backup creation success before proceeding with import
- Maintain backup file integrity

## 8. Installation and Deployment

### 8.1 Installer Requirements
- Single MSI installer package
- No registry modifications required
- Install to Program Files directory
- Create desktop shortcut option

### 8.2 Dependencies
- .NET runtime (included in installer if needed)
- Microsoft Word must be installed on target machine
- No additional third-party dependencies

## 9. Future Considerations

### 9.1 Potential Enhancements
- Scheduled automatic imports
- Network template support
- Multiple template management
- Content preview before import
- Change history tracking

### 9.2 Integration Possibilities
- SharePoint integration for source documents
- Email notifications for import completion
- API for integration with other systems