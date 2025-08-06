# Building Blocks Manager - Todo List

## Completed ✅
1. Create vertical button groups with labels: Import (Folder -> Template) and Export (Template -> Folder)
2. Arrange 'All' and 'Selected' buttons vertically in Import group
3. Arrange 'All' and 'Selected' buttons vertically in Export group
4. Add Query group with 'Query Source Directory' button
5. Remove Rollback button from main interface
6. Add Rollback option to File menu
7. Create TabControl with three tabs: Results, Directory, Template
8. Implement Results tab - maintain current program status display functionality
9. Implement Directory tab - show tree view of crawled directory structure
10. Populate Directory tab when Query Source Directory button is pressed
11. Add Template button to populate Template tab
12. Implement Template tab - show building blocks with category filtering/grouping
13. Add template existence check before populating Template tab
14. Implement category filtering and grouping in Template tab

## Pending ⏳
15. Add Filter button to Template tab that opens modal dialog
16. Create filter dialog with CheckedListBox for multi-category selection
17. Include 'System/Hex Entries' as filterable category in dialog
18. Add Select All/Select None buttons to filter dialog
19. Update Filter button text to show active filter count (e.g. 'Filter: 3 categories')
20. Apply selected filters to ListView and refresh display
21. Remove COM fallback code from WordManager.GetBuildingBlocks()

## Current Status Summary

**Major Success**: Fast XML approach working! Template tab loads quickly and shows all building blocks

**Performance**: Switched from slow COM (20-30s) to fast XML parsing (VB-style approach)

**Next Phase - Template Filtering**:
- **Problem**: Template shows system/hex entries alongside real building blocks
- **Solution**: Modal filter dialog (like existing Export Selected dialog)
- **Features**: Multi-category selection, hide system entries, Select All/None buttons
- **UX**: Filter button shows active filter count, clean Template tab layout

**Technical Notes**:
- Using .NET Framework 4.8 for Office interop compatibility
- XML approach: reads `word/glossary/document.xml` from Word document ZIP
- COM still needed for actual import/export operations (not just reading)
- Tab control sizing fixed, no spillover issues