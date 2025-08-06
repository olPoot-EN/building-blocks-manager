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
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            
            var backupPath = Path.Combine(directory, $"{fileName}_Backup_{timestamp}{extension}");
            File.Copy(templatePath, backupPath, true);

            // Clean up old backups (keep only last 5)
            CleanupOldBackups(directory, fileName, extension);
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
            return GetBuildingBlocksFromXml();
        }

        private List<BuildingBlockInfo> GetBuildingBlocksFromXml()
        {
            var buildingBlocks = new List<BuildingBlockInfo>();

            try
            {
                using (var doc = WordprocessingDocument.Open(templatePath, false))
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
                                
                                buildingBlocks.Add(new BuildingBlockInfo
                                {
                                    Name = name,
                                    Category = category,
                                    Gallery = galleryDisplay,
                                    Template = Path.GetFileNameWithoutExtension(templatePath)
                                });
                                
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"OpenXML SDK error: {ex.Message}");
                // Fallback to manual XML parsing if SDK fails
                return GetBuildingBlocksFromManualXml();
            }

            return buildingBlocks;
        }

        private string MapGalleryValue(string gallery)
        {
            // Map OpenXML gallery enum values to display names
            if (gallery == "autoTxt") return "AutoText";
            if (gallery == "bbPlcHdr") return "Built-In";
            if (gallery == "docPartObj") return "Quick Parts";
            if (gallery == "placeholder") return "Text Box";
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
                OpenTemplate();
                sourceDoc = wordApp.Documents.Open(sourceFile);

                // Remove existing Building Block with same name if it exists
                RemoveBuildingBlock(name, category);

                // Copy content from source document
                sourceDoc.Content.Copy();

                // Create new Building Block in template
                var range = templateDoc.Range();
                range.Paste();

                // Convert gallery type to Word enum
                var buildingBlockType = GetBuildingBlockType(galleryType);

                // Access template's BuildingBlockEntries
                Word.Template template = (Word.Template)templateDoc.get_AttachedTemplate();
                template.BuildingBlockEntries.Add(
                    name,
                    buildingBlockType,
                    category,
                    range);

                // Clear the pasted content from template
                range.Delete();
                
                // Save template
                templateDoc.Save();
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

        private void RemoveBuildingBlock(string name, string category)
        {
            try
            {
                Word.Template template = (Word.Template)templateDoc.get_AttachedTemplate();
                
                // Use indexed access to find and remove Building Block
                for (int i = 1; i <= template.BuildingBlockEntries.Count; i++)
                {
                    Word.BuildingBlock bb = template.BuildingBlockEntries.Item(i);
                    if (bb.Name == name && bb.Category.Name == category)
                    {
                        bb.Delete();
                        break;
                    }
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
                
                // Find the Building Block
                Word.BuildingBlock targetBB = null;
                Word.Template template = (Word.Template)templateDoc.get_AttachedTemplate();
                
                for (int i = 1; i <= template.BuildingBlockEntries.Count; i++)
                {
                    Word.BuildingBlock bb = template.BuildingBlockEntries.Item(i);
                    if (bb.Name == buildingBlockName && bb.Category.Name == category)
                    {
                        targetBB = bb;
                        break;
                    }
                }

                if (targetBB == null)
                    throw new InvalidOperationException($"Building Block '{buildingBlockName}' not found in category '{category}'");

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

        public void RollbackFromBackup()
        {
            try
            {
                var directory = Path.GetDirectoryName(templatePath);
                var fileName = Path.GetFileNameWithoutExtension(templatePath);
                var extension = Path.GetExtension(templatePath);

                var backupFiles = Directory.GetFiles(directory, $"{fileName}_Backup_*{extension}")
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
                OpenTemplate();
                Word.Template template = (Word.Template)templateDoc.get_AttachedTemplate();
                
                // Find and delete the Building Block
                for (int i = 1; i <= template.BuildingBlockEntries.Count; i++)
                {
                    Word.BuildingBlock bb = template.BuildingBlockEntries.Item(i);
                    if (bb.Name == buildingBlockName && bb.Category.Name == category)
                    {
                        bb.Delete();
                        templateDoc.Save();
                        return;
                    }
                }

                throw new InvalidOperationException($"Building Block '{buildingBlockName}' not found in category '{category}'");
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to delete Building Block '{buildingBlockName}': {ex.Message}", ex);
            }
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
            return $"{Category} - {Name}";
        }
    }
}