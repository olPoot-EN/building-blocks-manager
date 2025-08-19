# Building Blocks Manager - Troubleshooting Guide

## Common Issues & Solutions

### Files Show as "Modified" When They Shouldn't

**Symptoms**: Fresh exports or unchanged files show as "Modified" in directory tree

**Root Cause**: Tolerance comparison not being used, or analysis results not passed to UI

**Solution Checklist**:
1. Verify `Query Directory` calls `AnalyzeChanges()`: Check MainForm.cs:756
2. Verify directory tree uses analysis results: Check `PopulateDirectoryTree(files, analysis)`
3. Verify tolerance logic: `timeDifference.TotalMinutes > 1.0` in BuildingBlockLedger.cs:261
4. Check all method parameters include analysis: `CreateDirectoryNode(..., analysis)`

### Ledger Shows Empty / All Files Show as "New"

**Symptoms**: `GetLastImportTime()` always returns `DateTime.MinValue`

**Root Causes**:
- Ledger file doesn't exist where expected
- Ledger parsing fails due to format mismatch
- Different ledger directories used for save vs load

**Diagnostic Steps**:
1. Check ledger file exists: Use `BuildingBlockLedger.LedgerFileExists()`
2. Verify ledger directory: File → Ledger Directory shows correct path
3. Check file format: Ledger should have `Name (45 chars) Category (30 chars) YYYY-MM-DD HH:mm`
4. Verify parsing: Load() method uses regex `@"\s+(\d{4}-\d{2}-\d{2} \d{2}:\d{2})$"`

### Visual Studio Git Won't Pull Latest Changes

**Symptoms**: Code changes appear in GitHub but not in Visual Studio

**Corporate Environment**: Command line git may not be available

**Solutions (try in order)**:
1. Git Changes window → Refresh icon → Pull
2. View → Team Explorer → Sync → Fetch All → Pull
3. Close VS → Reopen solution → Try pull again
4. File → Close Solution → Open from Source Control → Re-enter GitHub URL
5. **Nuclear option**: Delete project folder → Clone fresh

### Directory Tree Doesn't Update After Query

**Symptoms**: Tree shows old data or remains empty after Query Directory

**Root Causes**:
- Query Directory doesn't call `PopulateDirectoryTree()`
- Analysis results not passed to tree population
- Tree view not switching to Directory tab

**Solution**: Verify MainForm.cs Query Directory process:
```csharp
var analysis = ledger.AnalyzeChanges(files);
PopulateDirectoryTree(files, analysis);
tabControl.SelectedTab = tabDirectory;
```

### Settings Not Persisting

**Symptoms**: Paths or configuration reset on restart

**Root Cause**: Settings not being saved or loaded properly

**Check**:
- Settings.Save() called when values change
- Settings.Load() called during startup
- Settings file location: `%LOCALAPPDATA%\BuildingBlocksManager\settings.txt`

### "The name 'analysis' does not exist" Build Error

**Symptoms**: Build fails with variable scope error

**Root Cause**: Analysis variable not in scope for method call

**Solution**: Check parameter passing chain:
1. `BtnQueryDirectory_Click` creates analysis
2. `PopulateDirectoryTree(files, analysis)` receives analysis
3. `CreateDirectoryNode(dir, files, analysis)` passes analysis down
4. All recursive calls include analysis parameter

### Timestamp Comparison Issues

**Symptoms**: Files with tiny time differences show as modified

**Root Cause**: Not using tolerance comparison

**Expected Behavior**: Only files modified >1 minute after ledger timestamp show as "Modified"

**Verify**: BuildingBlockLedger.cs AnalyzeChanges uses:
```csharp
if (timeDifference.TotalMinutes > 1.0)
    // Modified
else
    // Unchanged
```

## Debug Output Removal

All `System.Diagnostics.Debug.WriteLine()` troubleshooting output has been removed for production. 

If debugging needed:
- Add temporary debug output
- Remove before committing
- Don't leave debug spam in production code