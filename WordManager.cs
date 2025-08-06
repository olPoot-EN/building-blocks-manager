using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml;
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
                // EnableEvents doesn't exist in Word COM - remove it
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

            var directory = Path.GetDirectoryName(templatePath);
            var fileName = Path.GetFileNameWithoutExtension(templatePath);
            var extension = Path.GetExtension(templatePath);
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            
            var backupPath = Path.Combine(directory, $"{fileName}_Backup_{timestamp}{extension}");
            File.Copy(templatePath, backupPath, true);

            // Clean up old backups (keep only last 5)
            CleanupOldBackups(directory, fileName, extension);
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

            using (var archive = ZipFile.OpenRead(templatePath))
            {
                var glossaryEntry = archive.GetEntry("word/glossary/document.xml");
                
                if (glossaryEntry != null)
                {
                    using (var stream = glossaryEntry.Open())
                    {
                        var doc = new XmlDocument();
                        doc.Load(stream);

                        var namespaceManager = new XmlNamespaceManager(doc.NameTable);
                        namespaceManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                        var docParts = doc.SelectNodes("//w:docPart", namespaceManager);

                        if (docParts != null)
                        {
                            int debugCount = 0;
                            foreach (XmlNode docPart in docParts)
                            {
                                // Debug: dump first few entries to see XML structure
                                if (debugCount < 3)
                                {
                                    System.Diagnostics.Debug.WriteLine($"=== DocPart {debugCount} XML ===");
                                    System.Diagnostics.Debug.WriteLine(docPart.OuterXml);
                                    debugCount++;
                                }

                                // Try multiple possible XPath patterns for category
                                var nameNode = docPart.SelectSingleNode(".//w:docPartPr/w:name/@w:val", namespaceManager) ??
                                              docPart.SelectSingleNode(".//w:name/@w:val", namespaceManager);
                                              
                                var categoryNode = docPart.SelectSingleNode(".//w:docPartPr/w:category/@w:val", namespaceManager) ??
                                                  docPart.SelectSingleNode(".//w:category/@w:val", namespaceManager) ??
                                                  docPart.SelectSingleNode(".//w:docPartPr/w:category/w:name/@w:val", namespaceManager);
                                                  
                                var galleryNode = docPart.SelectSingleNode(".//w:docPartPr/w:gallery/@w:val", namespaceManager) ??
                                                 docPart.SelectSingleNode(".//w:gallery/@w:val", namespaceManager) ??
                                                 docPart.SelectSingleNode(".//w:docPartPr/w:types/w:type/@w:val", namespaceManager) ??
                                                 docPart.SelectSingleNode(".//w:types/w:type/@w:val", namespaceManager);

                                // Debug: show what we found
                                if (debugCount <= 3)
                                {
                                    System.Diagnostics.Debug.WriteLine($"Name: {nameNode?.Value ?? "NULL"}");
                                    System.Diagnostics.Debug.WriteLine($"Category: {categoryNode?.Value ?? "NULL"}");
                                    System.Diagnostics.Debug.WriteLine($"Gallery: {galleryNode?.Value ?? "NULL"}");
                                    System.Diagnostics.Debug.WriteLine("---");
                                }

                                if (nameNode != null)
                                {
                                    // Debug: show all possible gallery-related nodes
                                    if (debugCount <= 3)
                                    {
                                        var allGalleryNodes = docPart.SelectNodes(".//w:*[contains(local-name(), 'gallery') or contains(local-name(), 'type')]", namespaceManager);
                                        if (allGalleryNodes != null)
                                        {
                                            foreach (XmlNode node in allGalleryNodes)
                                            {
                                                System.Diagnostics.Debug.WriteLine($"Gallery-related node: {node.Name} = {node.InnerText} (attrs: {node.Attributes?.Count})");
                                                if (node.Attributes != null)
                                                {
                                                    foreach (XmlAttribute attr in node.Attributes)
                                                    {
                                                        System.Diagnostics.Debug.WriteLine($"  @{attr.Name} = {attr.Value}");
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    
                                    buildingBlocks.Add(new BuildingBlockInfo
                                    {
                                        Name = nameNode.Value,
                                        Category = categoryNode?.Value ?? "",
                                        Gallery = galleryNode?.Value ?? ""
                                    });
                                }
                            }
                        }
                    }
                }
            }

            return buildingBlocks;
        }


        public void ImportBuildingBlock(string sourceFile, string category, string name)
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

                // Access template's BuildingBlockEntries
                Word.Template template = (Word.Template)templateDoc.get_AttachedTemplate();
                template.BuildingBlockEntries.Add(
                    name,
                    Word.WdBuildingBlockTypes.wdTypeCustom1,
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

                // Dispose COM objects
                CloseTemplate();
                
                if (wordApp != null)
                {
                    wordApp.Quit();
                    Marshal.ReleaseComObject(wordApp);
                    wordApp = null;
                }

                disposed = true;
            }
        }

        ~WordManager()
        {
            Dispose(false);
        }
    }

    public class BuildingBlockInfo
    {
        public string Name { get; set; }
        public string Category { get; set; }
        public string Gallery { get; set; }
        
        public override string ToString()
        {
            return $"{Category} - {Name}";
        }
    }
}