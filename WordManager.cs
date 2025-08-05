using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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
            var buildingBlocks = new List<BuildingBlockInfo>();
            
            try
            {
                OpenTemplate();
                
                foreach (Word.BuildingBlock bb in templateDoc.BuildingBlockEntries)
                {
                    if (bb.Category.StartsWith("InternalAutotext"))
                    {
                        buildingBlocks.Add(new BuildingBlockInfo
                        {
                            Name = bb.Name,
                            Category = bb.Category,
                            Gallery = bb.Gallery.ToString()
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to retrieve Building Blocks: {ex.Message}", ex);
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

                templateDoc.BuildingBlockEntries.Add(
                    name,
                    Word.WdBuildingBlockTypes.wdTypeCustom1,
                    category,
                    range,
                    insertOptions: Word.WdDocPartInsertOptions.wdInsertContent);

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
                foreach (Word.BuildingBlock bb in templateDoc.BuildingBlockEntries)
                {
                    if (bb.Name == name && bb.Category == category)
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
                foreach (Word.BuildingBlock bb in templateDoc.BuildingBlockEntries)
                {
                    if (bb.Name == buildingBlockName && bb.Category == category)
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