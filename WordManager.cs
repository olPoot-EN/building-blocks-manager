using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
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

        public WordManager(string templatePath)
        {
            this.templatePath = templatePath;
        }

        private void InitializeWord()
        {
            if (wordApp == null)
            {
                wordApp = new Word.Application();
                wordApp.Visible = false;
                wordApp.ScreenUpdating = false;
                wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            }
        }

        private void OpenTemplate()
        {
            if (templateDoc == null)
            {
                InitializeWord();
                templateDoc = wordApp.Documents.Open(templatePath);
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
        
        public void ImportBuildingBlock(string sourceFile, string category, string name, string galleryType)
        {
            Word.Document sourceDoc = null;
            
            try
            {
                // Validate parameters before starting Word operations
                ValidateBuildingBlockParameters(name, category);
                
                System.Diagnostics.Debug.WriteLine($"[WordManager] Starting import of '{name}' from '{sourceFile}'");
                
                OpenTemplate();
                sourceDoc = wordApp.Documents.Open(sourceFile);

                // Verify source document has content
                if (sourceDoc.Content.Text.Trim().Length <= 1) // Word docs always have at least 1 char (paragraph mark)
                {
                    throw new InvalidOperationException("Source document appears to be empty");
                }

                // Remove existing Building Block with same name if it exists
                RemoveBuildingBlock(name, category);

                // Copy content from source document
                sourceDoc.Content.Copy();

                // Create new Building Block in template
                var range = templateDoc.Range();
                range.Paste();

                // Verify pasted content
                if (range.Text.Trim().Length == 0)
                {
                    throw new InvalidOperationException("Failed to paste content from source document");
                }

                // Convert gallery type to Word enum
                var buildingBlockType = GetBuildingBlockType(galleryType);

                // Access template's BuildingBlockEntries
                Word.Template template = (Word.Template)templateDoc.get_AttachedTemplate();
                
                // Sanitize parameters one more time before Word API call
                string sanitizedName = SanitizeBuildingBlockName(name);
                string sanitizedCategory = SanitizeBuildingBlockCategory(category);
                
                template.BuildingBlockEntries.Add(
                    sanitizedName,
                    buildingBlockType,
                    sanitizedCategory,
                    range);

                // Clear the pasted content from template
                range.Delete();
                
                // Save template
                templateDoc.Save();
                
                System.Diagnostics.Debug.WriteLine($"[WordManager] Successfully imported '{sanitizedName}' in category '{sanitizedCategory}'");
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
            
            if (name.Length > 64) // Word has a limit around 64 characters
                throw new ArgumentException($"Building Block name too long (max 64 chars): '{name}' ({name.Length} chars)");
            
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
            if (sanitized.Length > 64)
                sanitized = sanitized.Substring(0, 64).Trim();
            
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
                case 0x800A13E9: // Invalid parameter
                    return $"Invalid parameter. Name: '{name}' (length: {name?.Length ?? 0}), Category: '{category}' (length: {category?.Length ?? 0}). {baseMessage}";
                    
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
                    // Dispose managed resources
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
}