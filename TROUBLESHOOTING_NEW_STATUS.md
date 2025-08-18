# Troubleshooting: Building Blocks Show as "(New)" After Import

## Problem
Freshly imported building blocks continue to show as "(New)" in the directory tree even though they have been successfully imported and the ledger contains correct timestamps.

## Root Cause Analysis

### Investigation Steps Completed
1. **Added debug logging** to trace name/category extraction during directory scans
2. **Added debug logging** to trace ledger lookup process
3. **Identified ledger format/parser mismatch** as the core issue

### Key Findings

#### Debug Output Analysis
```
[LEDGER] Lookup Key: 'ExeSum_Generic_PRC024|ExecSum'
[LEDGER] Found in ledger: False
[LEDGER] Available keys in ledger: [empty]
```

**Problem**: Ledger lookup returns `False` even though entries exist in the ledger file.

#### Ledger File Format vs Parser Mismatch
**Ledger file contains** (space-aligned format):
```
ExeSum_Generic_PRC024                       ExecSum                           2025-08-18 14:26
```

**Parser was expecting** tab-delimited or pipe-delimited format, not space-aligned.

### Solution Implemented
1. **Fixed ledger save format** to use readable space-aligned columns
2. **Updated parser** to handle space-aligned format using regex:
   - Extracts date from end of line: `(\d{4}-\d{2}-\d{2} \d{2}:\d{2})$`
   - Splits name/category by finding 2+ consecutive spaces: `^(\S+(?:\s+\S+)*?)\s{2,}(.+)$`
3. **Removed unnecessary parser fallbacks** for tab/pipe formats

## Current Status
- **Code changes committed** to repository (commits: 28f4c88, 90cd020, 28d5ebf, 4ae0787, 0d2cc83, a56c2ef, 738cb1c, 51fa07f)
- **Parser logic fixed** for space-aligned ledger format
- **Debug output still available** via `System.Diagnostics.Debug.WriteLine` 

## Next Steps Required

### Immediate Testing
1. **Delete existing ledger file** to force regeneration with new format
2. **Import any building block** to regenerate ledger with fixed parser
3. **Run directory query** to verify debug output shows:
   ```
   [LEDGER] Found in ledger: True
   ```

### If Issue Persists
1. **Check actual ledger file content** - verify it matches expected space-aligned format
2. **Verify regex parsing** - ensure name/category extraction works for your specific data
3. **Check for edge cases**:
   - Building block names with multiple words/underscores
   - Category names with special characters
   - File paths with unusual characters

## Technical Details

### File Locations
- **Ledger file**: `C:\Kestrel\Templates\Office\Report automator testing\BBM_Logs\building_blocks_ledger.txt`
- **Debug output**: Available in Visual Studio Debug Output window

### Key Code Files Modified
- `BuildingBlockLedger.cs` - Ledger save/parse logic
- `FileManager.cs` - Directory scanning and name extraction
- `MainForm.cs` - Import process logging

### Expected Debug Flow
1. **[SCAN]** - Shows extracted names/categories from file scanning
2. **[LEDGER]** - Shows lookup keys and whether found in ledger
3. **[IMPORT]** - Shows names/categories stored during import

### IsNew Logic
```csharp
public bool IsNew => LastImported == DateTime.MinValue;
```
If `GetLastImportTime()` returns `DateTime.MinValue`, file shows as "(New)".

## Commit History
- `28f4c88` - Add debug logging to troubleshoot "new" status issue
- `90cd020` - Fix debug output for Windows Forms app  
- `28d5ebf` - Remove excessive ledger save success messages
- `4ae0787` - Add category extraction debug output
- `0d2cc83` - Fix ledger save format to use proper tab delimiters
- `a56c2ef` - Remove annoying debug message boxes
- `738cb1c` - Fix ledger format and parser for readable space-aligned format
- `51fa07f` - Remove unnecessary parser fallbacks