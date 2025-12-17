using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Word = Microsoft.Office.Interop.Word;

namespace BuildingBlocksManager
{
    public class WordManager : IDisposable
    {
        private string templatePath;
        private Word.Application wordApp;
        private Word.Document templateDoc;
        private bool disposed = false;

        // Dialog dismisser thread
        private Thread dialogDismisserThread;
        private volatile bool dismissDialogs = false;

        // Windows API for finding and clicking dialog buttons
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        private const uint WM_CLOSE = 0x0010;
        private const uint BM_CLICK = 0x00F5;

        public WordManager(string templatePath)
        {
            this.templatePath = templatePath;
        }

        /// <summary>
        /// Start background thread to auto-dismiss Word security dialogs
        /// </summary>
        private void StartDialogDismisser()
        {
            dismissDialogs = true;
            dialogDismisserThread = new Thread(DialogDismisserLoop)
            {
                IsBackground = true,
                Name = "WordDialogDismisser"
            };
            dialogDismisserThread.Start();
        }

        /// <summary>
        /// Stop the dialog dismisser thread
        /// </summary>
        private void StopDialogDismisser()
        {
            dismissDialogs = false;
            if (dialogDismisserThread != null && dialogDismisserThread.IsAlive)
            {
                dialogDismisserThread.Join(1000); // Wait up to 1 second
            }
        }

        /// <summary>
        /// Background loop that finds and clicks "Disable Macros" on Word security dialogs
        /// </summary>
        private void DialogDismisserLoop()
        {
            while (dismissDialogs)
            {
                try
                {
                    EnumWindows((hWnd, lParam) =>
                    {
                        if (!IsWindowVisible(hWnd))
                            return true; // Continue enumeration

                        var sb = new System.Text.StringBuilder(256);
                        GetWindowText(hWnd, sb, 256);
                        string title = sb.ToString();

                        // Look for Word security dialog
                        if (title.Contains("Microsoft Word Security Notice"))
                        {
                            System.Diagnostics.Debug.WriteLine($"[DIAG] Found security dialog, looking for button...");

                            // Find and click the "Disable Macros" button
                            EnumChildWindows(hWnd, (childHwnd, childLParam) =>
                            {
                                var childText = new System.Text.StringBuilder(256);
                                GetWindowText(childHwnd, childText, 256);
                                var childClass = new System.Text.StringBuilder(256);
                                GetClassName(childHwnd, childClass, 256);

                                string buttonText = childText.ToString();
                                string className = childClass.ToString();

                                // Look for button with "Disable Macros" text
                                if (className.Contains("Button") && buttonText.Contains("Disable Macros"))
                                {
                                    System.Diagnostics.Debug.WriteLine($"[DIAG] Clicking 'Disable Macros' button...");
                                    SendMessage(childHwnd, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
                                    return false; // Stop enumeration
                                }

                                return true; // Continue enumeration
                            }, IntPtr.Zero);
                        }

                        return true; // Continue enumeration
                    }, IntPtr.Zero);
                }
                catch
                {
                    // Ignore errors in dialog detection
                }

                Thread.Sleep(100); // Check every 100ms
            }
        }

        private void InitializeWord()
        {
            if (wordApp == null)
            {
                // Start dialog dismisser BEFORE creating Word instance
                StartDialogDismisser();

                System.Diagnostics.Debug.WriteLine("[DIAG] Creating Word.Application...");
                wordApp = new Word.Application();
                System.Diagnostics.Debug.WriteLine("[DIAG] Word.Application created");

                // Set security/alert settings BEFORE doing anything else
                // Value 1 = msoAutomationSecurityLow (run all macros without prompts)
                System.Diagnostics.Debug.WriteLine("[DIAG] Setting AutomationSecurity...");
                ((dynamic)wordApp).AutomationSecurity = 1;
                System.Diagnostics.Debug.WriteLine("[DIAG] Setting DisplayAlerts...");
                wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;

                // Additional settings to suppress security dialogs
                System.Diagnostics.Debug.WriteLine("[DIAG] Setting Options...");
                wordApp.Options.ConfirmConversions = false;
                wordApp.Options.DoNotPromptForConvert = true;

                System.Diagnostics.Debug.WriteLine("[DIAG] Setting Visible/ScreenUpdating...");
                wordApp.Visible = false;
                wordApp.ScreenUpdating = false;
                System.Diagnostics.Debug.WriteLine("[DIAG] InitializeWord complete");
            }
        }

        private void OpenTemplate()
        {
            if (templateDoc == null)
            {
                InitializeWord();
                System.Diagnostics.Debug.WriteLine($"[DIAG] Opening template: {templatePath}");
                templateDoc = wordApp.Documents.Open(templatePath);
                System.Diagnostics.Debug.WriteLine("[DIAG] Template opened successfully");
            }
        }

        public void CreateBackup()
        {
            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Template file not found: {templatePath}");

            // Check if template file is accessible before proceeding
            if (!IsFileAccessible(templatePath))
                throw new IOException($"Template file is locked or in use by another process: {templatePath}");

            var directory = Path.GetDirectoryName(templatePath);
            var fileName = Path.GetFileNameWithoutExtension(templatePath);
            var extension = Path.GetExtension(templatePath);
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmm");
            
            // Create backups directory
            var backupDirectory = Path.Combine(directory, "BBM_Backups");
            Directory.CreateDirectory(backupDirectory);
            
            var backupPath = Path.Combine(backupDirectory, $"{fileName}_Backup_{timestamp}{extension}");
            File.Copy(templatePath, backupPath, true);

            // Clean up old backups (keep only last 5)
            CleanupOldBackups(backupDirectory, fileName, extension);
        }
        
        private bool IsFileAccessible(string filePath)
        {
            try
            {
                using (var fileStream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    return true;
                }
            }
            catch (IOException)
            {
                return false;
            }
            catch (UnauthorizedAccessException)
            {
                return false;
            }
        }

        private void CleanupOldBackups(string directory, string fileName, string extension)
        {
            try
            {
                var backupFiles = Directory.GetFiles(directory, $"{fileName}_Backup_*{extension}")
                    .Select(f => new FileInfo(f))
                    .OrderByDescending(f => f.CreationTime)
                    .Skip(5);

                foreach (var file in backupFiles)
                {
                    File.Delete(file.FullName);
                }
            }
            catch
            {
                // Silently handle backup cleanup errors
            }
        }

        public List<BuildingBlockInfo> GetBuildingBlocks()
        {
            var allBuildingBlocks = new List<BuildingBlockInfo>();
            
            // Scan the selected template
            allBuildingBlocks.AddRange(GetBuildingBlocksFromTemplate(templatePath));
            
            
            
            return allBuildingBlocks;
        }

        
        
        private List<BuildingBlockInfo> GetBuildingBlocksFromTemplate(string templateFilePath)
        {
            var buildingBlocks = new List<BuildingBlockInfo>();

            try
            {
                using (var doc = WordprocessingDocument.Open(templateFilePath, false))
                {
                    var glossaryPart = doc.MainDocumentPart?.GlossaryDocumentPart;
                    if (glossaryPart?.GlossaryDocument?.DocParts != null)
                    {
                        foreach (var docPart in glossaryPart.GlossaryDocument.DocParts.Elements<DocPart>())
                        {
                            var properties = docPart.DocPartProperties;
                            if (properties != null)
                            {
                                var name = properties.DocPartName?.Val?.Value ?? "";
                                
                                // Extract category and gallery using the discovered structure
                                var category = "";
                                var gallery = "";
                                var types = "";
                                
                                var categoryElement = properties.GetFirstChild<Category>();
                                if (categoryElement != null)
                                {
                                    var categoryName = categoryElement.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Name>();
                                    if (categoryName?.Val != null)
                                        category = categoryName.Val.Value;
                                    
                                    var galleryElement = categoryElement.GetFirstChild<Gallery>();
                                    if (galleryElement?.Val != null)
                                        gallery = galleryElement.Val.Value.ToString();
                                }
                                
                                // Get DocPartTypes as fallback
                                var docPartTypes = properties.GetFirstChild<DocPartTypes>();
                                if (docPartTypes != null)
                                {
                                    var docPartType = docPartTypes.GetFirstChild<DocPartType>();
                                    if (docPartType?.Val != null)
                                        types = docPartType.Val.Value.ToString();
                                }

                                // Use gallery or fall back to types, then map to display names
                                string galleryValue = !string.IsNullOrEmpty(gallery) ? gallery : types;
                                string galleryDisplay = MapGalleryValue(galleryValue);
                                
                                // Ensure we always have a gallery - default to AutoText if empty
                                if (string.IsNullOrEmpty(galleryDisplay))
                                {
                                    galleryDisplay = "AutoText";
                                }
                                
                                buildingBlocks.Add(new BuildingBlockInfo
                                {
                                    Name = name,
                                    Category = category,
                                    Gallery = galleryDisplay,
                                    Template = Path.GetFileNameWithoutExtension(templateFilePath)
                                });
                                
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"OpenXML SDK error for {templateFilePath}: {ex.Message}");
            }

            return buildingBlocks;
        }

        private string MapGalleryValue(string gallery)
        {
            // Map OpenXML gallery enum values to display names
            if (gallery == "autoTxt") return "AutoText";
            if (gallery == "bbPlcHdr") return "Built-In";
            if (gallery == "docPartObj") return "Quick Parts";
            if (gallery == "placeholder") return "Placeholder";
            return gallery;
        }
        
        private Word.WdBuildingBlockTypes GetBuildingBlockType(string galleryType)
        {
            // Convert gallery type string to Word building block type enum
            switch (galleryType)
            {
                case "AutoText":
                    return Word.WdBuildingBlockTypes.wdTypeAutoText;
                case "Quick Parts":
                    return Word.WdBuildingBlockTypes.wdTypeQuickParts;
                case "Custom Gallery 1":
                    return Word.WdBuildingBlockTypes.wdTypeCustom1;
                case "Custom Gallery 2":
                    return Word.WdBuildingBlockTypes.wdTypeCustom2;
                case "Custom Gallery 3":
                    return Word.WdBuildingBlockTypes.wdTypeCustom3;
                case "Custom Gallery 4":
                    return Word.WdBuildingBlockTypes.wdTypeCustom4;
                case "Custom Gallery 5":
                    return Word.WdBuildingBlockTypes.wdTypeCustom5;
                default:
                    return Word.WdBuildingBlockTypes.wdTypeAutoText; // Default to AutoText
            }
        }

        private List<BuildingBlockInfo> GetBuildingBlocksFromManualXml()
        {
            // Fallback to original manual XML parsing
            var buildingBlocks = new List<BuildingBlockInfo>();
            // ... existing manual XML code would go here if needed
            return buildingBlocks;
        }


        public void ImportBuildingBlock(string sourceFile, string category, string name)
        {
            ImportBuildingBlock(sourceFile, category, name, "AutoText"); // Default to AutoText
        }
        
        public ImportResult ImportBuildingBlock(string sourceFile, string category, string name, string galleryType)
        {
            Word.Document sourceDoc = null;
            var result = new ImportResult { Success = false };

            try
            {
                System.Diagnostics.Debug.WriteLine($"[DIAG] ImportBuildingBlock started: {name}");

                // Validate parameters before starting Word operations
                ValidateBuildingBlockParameters(name, category);
                System.Diagnostics.Debug.WriteLine("[DIAG] Parameters validated");

                OpenTemplate();
                System.Diagnostics.Debug.WriteLine("[DIAG] Template open, now opening source file...");
                System.Diagnostics.Debug.WriteLine($"[DIAG] Source file: {sourceFile}");
                sourceDoc = wordApp.Documents.Open(
                    sourceFile,
                    ReadOnly: true,
                    AddToRecentFiles: false);
                System.Diagnostics.Debug.WriteLine("[DIAG] Source file opened");

                // Remove existing Building Block with same name if it exists
                System.Diagnostics.Debug.WriteLine("[DIAG] Removing existing building block if exists...");
                RemoveBuildingBlock(name, category);
                System.Diagnostics.Debug.WriteLine("[DIAG] Remove complete");

                // Use source document content directly
                System.Diagnostics.Debug.WriteLine("[DIAG] Getting source content...");
                var sourceRange = sourceDoc.Content;
                System.Diagnostics.Debug.WriteLine("[DIAG] Got source content");

                // Convert gallery type to Word enum
                var buildingBlockType = GetBuildingBlockType(galleryType);

                // Access template's BuildingBlockEntries
                System.Diagnostics.Debug.WriteLine("[DIAG] Getting template attachment...");
                Word.Template template = (Word.Template)templateDoc.get_AttachedTemplate();
                System.Diagnostics.Debug.WriteLine("[DIAG] Got template attachment");
                
                // Sanitize parameters before Word API call
                string sanitizedName = SanitizeBuildingBlockName(name);
                string sanitizedCategory = SanitizeBuildingBlockCategory(category);
                
                // Check for empty category - Word doesn't like empty strings for BuildingBlockEntries
                bool categoryWasEmpty = string.IsNullOrWhiteSpace(sanitizedCategory);
                if (categoryWasEmpty)
                {
                    sanitizedCategory = "General"; // Use default category instead of empty/null
                    result.CategoryChanged = true;
                    result.OriginalCategory = category;
                    result.AssignedCategory = sanitizedCategory;
                }
                
                // Add Building Block using modern BuildingBlockEntries API
                System.Diagnostics.Debug.WriteLine($"[DIAG] Adding building block: {sanitizedName} to category: {sanitizedCategory}");
                try
                {
                    template.BuildingBlockEntries.Add(
                        sanitizedName,
                        buildingBlockType,
                        sanitizedCategory,
                        sourceRange);
                    System.Diagnostics.Debug.WriteLine("[DIAG] Building block added successfully");
                }
                catch (COMException comEx) when ((uint)comEx.ErrorCode == 0x800A16DD && galleryType == "AutoText")
                {
                    System.Diagnostics.Debug.WriteLine("[DIAG] AutoText failed, trying Quick Parts...");
                    // Try Quick Parts as fallback for AutoText compatibility issues
                    var quickPartsType = GetBuildingBlockType("Quick Parts");
                    template.BuildingBlockEntries.Add(
                        sanitizedName,
                        quickPartsType,
                        sanitizedCategory,
                        sourceRange);
                    System.Diagnostics.Debug.WriteLine("[DIAG] Quick Parts fallback succeeded");
                }

                // No temporary document to close

                // Save template
                System.Diagnostics.Debug.WriteLine("[DIAG] Saving template...");
                templateDoc.Save();
                System.Diagnostics.Debug.WriteLine("[DIAG] Template saved");
                
                result.Success = true;
                result.ImportedName = sanitizedName;
                result.FinalCategory = sanitizedCategory;
                
                return result;
            }
            catch (COMException comEx)
            {
                string detailedMessage = GetDetailedComError(comEx, name, category);
                throw new InvalidOperationException($"Failed to import Building Block '{name}': {detailedMessage}", comEx);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to import Building Block '{name}': {ex.Message}", ex);
            }
            finally
            {
                if (sourceDoc != null)
                {
                    sourceDoc.Close(false);
                    Marshal.ReleaseComObject(sourceDoc);
                }
            }
        }

        private void ValidateBuildingBlockParameters(string name, string category)
        {
            // Building Block name validation
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("Building Block name cannot be empty");
            
            if (name.Length > 32) // Word AutoText name limit is 32 characters
                throw new ArgumentException($"Building Block name too long (max 32 chars): '{name}' ({name.Length} chars)");
            
            // Check for invalid characters in name
            char[] invalidNameChars = { '/', '\\', ':', '*', '?', '"', '<', '>', '|', '\t', '\n', '\r' };
            if (name.IndexOfAny(invalidNameChars) >= 0)
                throw new ArgumentException($"Building Block name contains invalid characters: '{name}'");
            
            // Category validation (can be empty)
            if (!string.IsNullOrEmpty(category))
            {
                if (category.Length > 128) // Conservative limit
                    throw new ArgumentException($"Category too long (max 128 chars): '{category}' ({category.Length} chars)");
                
                // Check for invalid characters in category
                char[] invalidCategoryChars = { '*', '?', '"', '<', '>', '|', '\t', '\n', '\r' };
                if (category.IndexOfAny(invalidCategoryChars) >= 0)
                    throw new ArgumentException($"Category contains invalid characters: '{category}'");
            }
        }
        
        private string SanitizeBuildingBlockName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Untitled";
            
            // Replace problematic characters with underscores
            string sanitized = name;
            char[] problematicChars = { '/', '\\', ':', '*', '?', '"', '<', '>', '|', '\t', '\n', '\r' };
            
            foreach (char c in problematicChars)
            {
                sanitized = sanitized.Replace(c, '_');
            }
            
            // Trim and limit length
            sanitized = sanitized.Trim();
            if (sanitized.Length > 32)
                sanitized = sanitized.Substring(0, 32).Trim();

            return sanitized;
        }

        private string SanitizeBuildingBlockCategory(string category)
        {
            if (string.IsNullOrWhiteSpace(category))
                return "";
            
            // Replace problematic characters with underscores
            string sanitized = category;
            char[] problematicChars = { '*', '?', '"', '<', '>', '|', '\t', '\n', '\r' };
            
            foreach (char c in problematicChars)
            {
                sanitized = sanitized.Replace(c, '_');
            }
            
            // Trim and limit length
            sanitized = sanitized.Trim();
            if (sanitized.Length > 128)
                sanitized = sanitized.Substring(0, 128).Trim();
            
            return sanitized;
        }
        
        private string GetDetailedComError(COMException comEx, string name, string category)
        {
            string baseMessage = comEx.Message;
            
            // Common COM error codes for Building Blocks
            switch ((uint)comEx.ErrorCode)
            {
                case 0x800A13E9: // Invalid parameter (generic)
                    return $"Invalid parameter. Name: '{name}' (length: {name?.Length ?? 0}), Category: '{category}' (length: {category?.Length ?? 0}). {baseMessage}";
                    
                case 0x800A16DD: // Invalid parameter (specific to Building Block operations)
                    return $"Invalid Building Block parameter. This usually means:\n" +
                           $"• Name '{name}' contains invalid characters or is too long\n" +
                           $"• Category '{category}' is malformed\n" +
                           $"• Document content is corrupted or inaccessible\n" +
                           $"• Template file has issues\n" +
                           $"Name length: {name?.Length ?? 0}, Category length: {category?.Length ?? 0}. {baseMessage}";
                    
                case 0x800A03EC: // Name already exists
                    return $"Building Block with name '{name}' already exists. {baseMessage}";
                    
                case 0x800A175D: // Document is read-only
                    return $"Template document is read-only. {baseMessage}";
                    
                case 0x80080005: // Server execution failed
                    return $"Word automation failed. Try closing Word and running again. {baseMessage}";
                    
                default:
                    return $"COM Error 0x{comEx.ErrorCode:X8}: {baseMessage}";
            }
        }

        private void RemoveBuildingBlock(string name, string category)
        {
            try
            {
                Word.BuildingBlock targetBB = FindBuildingBlockByName(name);
                if (targetBB != null)
                {
                    targetBB.Delete();
                }
            }
            catch
            {
                // Silently handle if Building Block doesn't exist
            }
        }

        public void ExportBuildingBlock(string buildingBlockName, string category, string outputPath)
        {
            Word.Document newDoc = null;
            
            try
            {
                OpenTemplate();
                
                // Find the Building Block by name only
                Word.BuildingBlock targetBB = FindBuildingBlockByName(buildingBlockName);

                if (targetBB == null)
                {
                    throw new InvalidOperationException($"Building Block '{buildingBlockName}' not found");
                }

                // Create new document and insert Building Block content
                newDoc = wordApp.Documents.Add();
                var range = newDoc.Range();
                targetBB.Insert(range);

                // Remove the trailing paragraph mark that Word adds automatically
                // Check if the last paragraph is empty and remove it
                if (newDoc.Paragraphs.Count > 1)
                {
                    Word.Paragraph lastPara = newDoc.Paragraphs[newDoc.Paragraphs.Count];
                    // Check if the last paragraph is just a single paragraph mark
                    if (lastPara.Range.Text == "\r" || string.IsNullOrWhiteSpace(lastPara.Range.Text))
                    {
                        lastPara.Range.Delete();
                    }
                }
                else if (newDoc.Paragraphs.Count == 1)
                {
                    // For single paragraph content, remove trailing paragraph mark if present
                    Word.Paragraph para = newDoc.Paragraphs[1];
                    string text = para.Range.Text;
                    if (text.EndsWith("\r"))
                    {
                        // Select the final paragraph mark and delete it
                        Word.Range endRange = newDoc.Range();
                        endRange.SetRange(newDoc.Content.End - 1, newDoc.Content.End);
                        if (endRange.Text == "\r")
                        {
                            endRange.Delete();
                        }
                    }
                }

                // Save the new document
                newDoc.SaveAs2(outputPath);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to export Building Block '{buildingBlockName}': {ex.Message}", ex);
            }
            finally
            {
                if (newDoc != null)
                {
                    newDoc.Close(false);
                    Marshal.ReleaseComObject(newDoc);
                }
            }
        }

        private Word.BuildingBlock FindBuildingBlockByName(string buildingBlockName)
        {
            Word.Template template = (Word.Template)templateDoc.get_AttachedTemplate();
            
            // Simple name-only search - assumes Building Block names are unique
            for (int i = 1; i <= template.BuildingBlockEntries.Count; i++)
            {
                Word.BuildingBlock bb = template.BuildingBlockEntries.Item(i);
                if (bb.Name == buildingBlockName)
                {
                    return bb;
                }
            }

            return null;
        }

        private void ExportBuildingBlockOpenXML(string buildingBlockName, string category, string outputPath)
        {
            using (var templateDoc = WordprocessingDocument.Open(templatePath, false))
            {
                var glossaryPart = templateDoc.MainDocumentPart?.GlossaryDocumentPart;
                if (glossaryPart?.GlossaryDocument?.DocParts == null)
                    throw new InvalidOperationException("No Building Blocks found in template");

                DocPart targetDocPart = null;
                
                // Find the target Building Block
                foreach (var docPart in glossaryPart.GlossaryDocument.DocParts.Elements<DocPart>())
                {
                    var properties = docPart.DocPartProperties;
                    if (properties != null)
                    {
                        var name = properties.DocPartName?.Val?.Value ?? "";
                        
                        // Extract category
                        var actualCategory = "";
                        var categoryElement = properties.GetFirstChild<Category>();
                        if (categoryElement != null)
                        {
                            var categoryName = categoryElement.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Name>();
                            if (categoryName?.Val != null)
                                actualCategory = categoryName.Val.Value;
                        }
                        
                        // Match logic: if search category is empty, match empty categories; otherwise exact match
                        bool categoryMatch = string.IsNullOrEmpty(category) ? 
                            string.IsNullOrEmpty(actualCategory) : 
                            actualCategory == category;
                        
                        if (name == buildingBlockName && categoryMatch)
                        {
                            targetDocPart = docPart;
                            break;
                        }
                    }
                }

                if (targetDocPart == null)
                {
                    // Create detailed error message
                    var availableNames = new List<string>();
                    foreach (var docPart in glossaryPart.GlossaryDocument.DocParts.Elements<DocPart>())
                    {
                        var properties = docPart.DocPartProperties;
                        if (properties != null)
                        {
                            var name = properties.DocPartName?.Val?.Value ?? "";
                            var actualCategory = "";
                            var categoryElement = properties.GetFirstChild<Category>();
                            if (categoryElement != null)
                            {
                                var categoryName = categoryElement.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Name>();
                                if (categoryName?.Val != null)
                                    actualCategory = categoryName.Val.Value;
                            }
                            availableNames.Add($"'{name}' (Category: '{actualCategory}')");
                        }
                    }
                    
                    var errorMsg = $"Building Block '{buildingBlockName}' not found in category '{category}'. Available Building Blocks: {string.Join(", ", availableNames.Take(10))}";
                    if (availableNames.Count > 10)
                        errorMsg += $"... and {availableNames.Count - 10} more";
                    
                    throw new InvalidOperationException(errorMsg);
                }

                // Create new document and copy the Building Block content
                using (var newDoc = WordprocessingDocument.Create(outputPath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    // Add main document part
                    var mainPart = newDoc.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body());

                    // Copy the Building Block content to the new document
                    var targetBody = targetDocPart.DocPartBody;
                    if (targetBody != null)
                    {
                        // Copy children from DocPartBody to the new document's Body
                        // DocPartBody and Body are different types but have similar content structure
                        foreach (var element in targetBody.ChildElements)
                        {
                            var clonedElement = element.CloneNode(true);
                            mainPart.Document.Body.Append(clonedElement);
                        }
                    }

                    mainPart.Document.Save();
                }
            }
        }

        public void RollbackFromBackup()
        {
            try
            {
                var directory = Path.GetDirectoryName(templatePath);
                var fileName = Path.GetFileNameWithoutExtension(templatePath);
                var extension = Path.GetExtension(templatePath);

                // Look in backups directory
                var backupDirectory = Path.Combine(directory, "BBM_Backups");
                if (!Directory.Exists(backupDirectory))
                    throw new DirectoryNotFoundException("No backup directory found");

                var backupFiles = Directory.GetFiles(backupDirectory, $"{fileName}_Backup_*{extension}")
                    .Select(f => new FileInfo(f))
                    .OrderByDescending(f => f.CreationTime)
                    .FirstOrDefault();

                if (backupFiles == null)
                    throw new FileNotFoundException("No backup files found");

                // Close template if open
                CloseTemplate();

                // Replace current template with backup
                File.Copy(backupFiles.FullName, templatePath, true);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to rollback from backup: {ex.Message}", ex);
            }
        }


        public void DeleteBuildingBlock(string buildingBlockName, string category)
        {
            try
            {
                // Ensure any COM connections are closed to avoid file access conflicts
                CloseTemplate();
                
                // Check file accessibility before proceeding
                if (!IsFileAccessible(templatePath))
                    throw new IOException($"Template file is locked or in use: {templatePath}");

                // Pure XML deletion approach
                DeleteBuildingBlockXML(buildingBlockName, category);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to delete Building Block '{buildingBlockName}': {ex.Message}", ex);
            }
        }

        private void DeleteBuildingBlockXML(string buildingBlockName, string category)
        {
            using (var doc = WordprocessingDocument.Open(templatePath, true)) // Writable
            {
                var glossaryPart = doc.MainDocumentPart?.GlossaryDocumentPart;
                if (glossaryPart?.GlossaryDocument?.DocParts == null)
                    throw new InvalidOperationException("No Building Blocks found in template");

                DocPart targetDocPart = null;
                
                // Find the target Building Block
                foreach (var docPart in glossaryPart.GlossaryDocument.DocParts.Elements<DocPart>())
                {
                    var properties = docPart.DocPartProperties;
                    if (properties != null)
                    {
                        var name = properties.DocPartName?.Val?.Value ?? "";
                        
                        // Extract category
                        var actualCategory = "";
                        var categoryElement = properties.GetFirstChild<Category>();
                        if (categoryElement != null)
                        {
                            var categoryName = categoryElement.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Name>();
                            if (categoryName?.Val != null)
                                actualCategory = categoryName.Val.Value;
                        }
                        
                        // Match building block name and category
                        if (name == buildingBlockName && 
                            (string.IsNullOrEmpty(category) || actualCategory == category))
                        {
                            targetDocPart = docPart;
                            break;
                        }
                    }
                }

                if (targetDocPart == null)
                {
                    throw new InvalidOperationException($"Building Block '{buildingBlockName}' not found");
                }

                // Remove the DocPart from the XML
                targetDocPart.Remove();
                
                // Save changes (automatically handled by WordprocessingDocument disposal)
            }
        }

        private class BuildingBlockLocation
        {
            public int Index { get; set; }
            public string Name { get; set; }
            public string Category { get; set; }
            public bool Found { get; set; }
        }

        private BuildingBlockLocation FindBuildingBlockLocationXML(string buildingBlockName, string category)
        {
            try
            {
                using (var doc = WordprocessingDocument.Open(templatePath, false))
                {
                    var glossaryPart = doc.MainDocumentPart?.GlossaryDocumentPart;
                    if (glossaryPart?.GlossaryDocument?.DocParts != null)
                    {
                        int index = 0;
                        foreach (var docPart in glossaryPart.GlossaryDocument.DocParts.Elements<DocPart>())
                        {
                            index++; // COM collections are 1-based
                            var properties = docPart.DocPartProperties;
                            if (properties != null)
                            {
                                var name = properties.DocPartName?.Val?.Value ?? "";
                                
                                // Extract category
                                var actualCategory = "";
                                var categoryElement = properties.GetFirstChild<Category>();
                                if (categoryElement != null)
                                {
                                    var categoryName = categoryElement.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Name>();
                                    if (categoryName?.Val != null)
                                        actualCategory = categoryName.Val.Value;
                                }
                                
                                // Match building block name and category
                                if (name == buildingBlockName && 
                                    (string.IsNullOrEmpty(category) || actualCategory == category))
                                {
                                    return new BuildingBlockLocation
                                    {
                                        Index = index,
                                        Name = name,
                                        Category = actualCategory,
                                        Found = true
                                    };
                                }
                            }
                        }
                    }
                }
                
                return new BuildingBlockLocation { Found = false };
            }
            catch
            {
                return new BuildingBlockLocation { Found = false };
            }
        }

        private Word.BuildingBlock GetBuildingBlockByIndex(int index)
        {
            Word.Template template = (Word.Template)templateDoc.get_AttachedTemplate();
            if (index >= 1 && index <= template.BuildingBlockEntries.Count)
            {
                return template.BuildingBlockEntries.Item(index);
            }
            return null;
        }

        private void CloseTemplate()
        {
            if (templateDoc != null)
            {
                templateDoc.Close(false);
                Marshal.ReleaseComObject(templateDoc);
                templateDoc = null;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    // Stop dialog dismisser thread
                    StopDialogDismisser();
                }

                // Dispose COM objects aggressively
                CloseTemplate();
                
                if (wordApp != null)
                {
                    try
                    {
                        // Force close any remaining documents in our Word instance
                        while (wordApp.Documents.Count > 0)
                        {
                            wordApp.Documents[1].Close(false);
                        }
                        
                        // Quit Word gracefully first
                        wordApp.Quit(false);
                        
                        // Wait a moment for Word to close gracefully
                        System.Threading.Thread.Sleep(1000);
                    }
                    catch (COMException)
                    {
                        // Word might have already been closed or become unresponsive
                        // Just ignore - the finally block will handle cleanup
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(wordApp);
                        wordApp = null;
                        
                        // Force garbage collection to clean up COM references
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
                }

                disposed = true;
            }
        }

        ~WordManager()
        {
            Dispose(false);
        }
        
        
        /// <summary>
        /// Check if the template file might be locked by checking for Word processes with the file open
        /// </summary>
        public static bool IsTemplateFileLocked(string templatePath)
        {
            // First try to open the file directly
            try
            {
                using (var fileStream = File.Open(templatePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    return false; // File is accessible
                }
            }
            catch (IOException)
            {
                return true; // File is locked
            }
            catch (UnauthorizedAccessException)
            {
                return true; // File is locked
            }
        }
        
        /// <summary>
        /// Get Word processes that might be using our template (for informational purposes)
        /// </summary>
        public static List<Process> GetWordProcessesUsingFile(string templatePath)
        {
            var suspiciousProcesses = new List<Process>();
            
            try
            {
                var wordProcesses = Process.GetProcessesByName("WINWORD");
                
                // We can't easily determine which specific Word process has a file open
                // but we can return all Word processes for user information
                foreach (var process in wordProcesses)
                {
                    try
                    {
                        // Only add processes that are still running
                        if (!process.HasExited)
                        {
                            suspiciousProcesses.Add(process);
                        }
                    }
                    catch
                    {
                        // Process may have exited while we were checking
                    }
                }
            }
            catch
            {
                // Ignore errors in process enumeration
            }
            
            return suspiciousProcesses;
        }
        
        /// <summary>
        /// Force kill specific Word processes
        /// </summary>
        public static bool KillWordProcesses(List<Process> processesToKill)
        {
            bool allKilled = true;
            
            foreach (var process in processesToKill)
            {
                try
                {
                    if (!process.HasExited)
                    {
                        process.Kill();
                        process.WaitForExit(3000); // Wait up to 3 seconds
                    }
                }
                catch
                {
                    allKilled = false;
                }
            }
            
            return allKilled;
        }
    }

    public class BuildingBlockInfo
    {
        public string Name { get; set; }
        public string Category { get; set; }
        public string Gallery { get; set; }
        public string Template { get; set; }
        
        public override string ToString()
        {
            return Name;
        }
    }

    public class ImportResult
    {
        public bool Success { get; set; }
        public string ImportedName { get; set; }
        public string FinalCategory { get; set; }
        public bool CategoryChanged { get; set; }
        public string OriginalCategory { get; set; }
        public string AssignedCategory { get; set; }
    }
}