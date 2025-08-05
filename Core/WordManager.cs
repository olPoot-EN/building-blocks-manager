using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;

namespace BuildingBlocksManager.Core
{
    public class WordManager : IDisposable
    {
        private Microsoft.Office.Interop.Word.Application _wordApp;
        private Document _templateDoc;
        private bool _disposed = false;

        public class BuildingBlockInfo
        {
            public string Name { get; set; }
            public string Category { get; set; }
            public string Gallery { get; set; }
            public string Description { get; set; }
            public DateTime CreatedDate { get; set; }
            public DateTime ModifiedDate { get; set; }
        }

        public class ImportResult
        {
            public bool Success { get; set; }
            public string ErrorMessage { get; set; }
            public string BuildingBlockName { get; set; }
            public string Category { get; set; }
        }

        public class ExportResult
        {
            public bool Success { get; set; }
            public string ErrorMessage { get; set; }
            public string FilePath { get; set; }
            public string BuildingBlockName { get; set; }
        }

        public WordManager()
        {
            InitializeWordApplication();
        }

        private void InitializeWordApplication()
        {
            try
            {
                _wordApp = new Microsoft.Office.Interop.Word.Application();
                _wordApp.Visible = false;
                _wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                _wordApp.ScreenUpdating = false;
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException($"Failed to initialize Word application: {ex.Message}", ex);
            }
        }

        public void OpenTemplate(string templatePath)
        {
            if (!File.Exists(templatePath))
            {
                throw new FileNotFoundException($"Template file not found: {templatePath}");
            }

            try
            {
                CloseTemplate();
                _templateDoc = _wordApp.Documents.Open(templatePath);
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException($"Failed to open template: {ex.Message}", ex);
            }
        }

        public void CloseTemplate()
        {
            if (_templateDoc != null)
            {
                try
                {
                    _templateDoc.Close(SaveChanges: false);
                    Marshal.ReleaseComObject(_templateDoc);
                }
                catch (COMException)
                {
                    // Ignore errors when closing
                }
                finally
                {
                    _templateDoc = null;
                }
            }
        }

        public string CreateBackup(string templatePath)
        {
            if (!File.Exists(templatePath))
            {
                throw new FileNotFoundException($"Template file not found: {templatePath}");
            }

            var directory = Path.GetDirectoryName(templatePath);
            var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(templatePath);
            var extension = Path.GetExtension(templatePath);
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            
            var backupFileName = $"{fileNameWithoutExtension}_Backup_{timestamp}{extension}";
            var backupPath = Path.Combine(directory, backupFileName);

            try
            {
                File.Copy(templatePath, backupPath);
                return backupPath;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to create backup: {ex.Message}", ex);
            }
        }

        public void RestoreFromBackup(string backupPath, string templatePath)
        {
            if (!File.Exists(backupPath))
            {
                throw new FileNotFoundException($"Backup file not found: {backupPath}");
            }

            try
            {
                CloseTemplate();
                
                if (File.Exists(templatePath))
                {
                    File.Delete(templatePath);
                }
                
                File.Copy(backupPath, templatePath);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to restore from backup: {ex.Message}", ex);
            }
        }

        public List<string> GetBackupFiles(string templatePath)
        {
            var directory = Path.GetDirectoryName(templatePath);
            var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(templatePath);
            var backups = new List<string>();

            try
            {
                var pattern = $"{fileNameWithoutExtension}_Backup_*.dotm";
                var files = Directory.GetFiles(directory, pattern);
                backups.AddRange(files);
                
                // Sort by creation time, newest first
                backups.Sort((x, y) => File.GetCreationTime(y).CompareTo(File.GetCreationTime(x)));
            }
            catch (Exception)
            {
                // Return empty list if directory scan fails
            }

            return backups;
        }

        public void CleanupOldBackups(string templatePath, int keepCount = 5)
        {
            var backups = GetBackupFiles(templatePath);
            
            if (backups.Count <= keepCount)
                return;

            // Delete backups beyond the keep count
            for (int i = keepCount; i < backups.Count; i++)
            {
                try
                {
                    File.Delete(backups[i]);
                }
                catch (Exception)
                {
                    // Ignore errors when deleting old backups
                }
            }
        }

        public ImportResult ImportBuildingBlock(string sourceFilePath, string buildingBlockName, string category)
        {
            var result = new ImportResult
            {
                BuildingBlockName = buildingBlockName,
                Category = category
            };

            if (_templateDoc == null)
            {
                result.ErrorMessage = "No template document is open";
                return result;
            }

            Document sourceDoc = null;

            try
            {
                // Open source document
                sourceDoc = _wordApp.Documents.Open(sourceFilePath, ReadOnly: true);
                
                // Get the content from source document
                Range sourceRange = sourceDoc.Content;
                sourceRange.Copy();

                // Create new document to hold the content temporarily
                Document tempDoc = _wordApp.Documents.Add();
                tempDoc.Content.Paste();

                // Create Building Block
                var attachedTemplate = (Template)_templateDoc.get_AttachedTemplate();
                BuildingBlock bb = attachedTemplate.BuildingBlockEntries.Add(
                    Name: buildingBlockName,
                    Type: WdBuildingBlockTypes.wdTypeCustom1,
                    Category: category,
                    Range: tempDoc.Content,
                    Description: $"Imported from {Path.GetFileName(sourceFilePath)}",
                    InsertOptions: WdDocPartInsertOptions.wdInsertContent
                );

                tempDoc.Close(SaveChanges: false);
                Marshal.ReleaseComObject(tempDoc);

                result.Success = true;
            }
            catch (COMException ex)
            {
                result.ErrorMessage = $"Word automation error: {ex.Message}";
            }
            catch (Exception ex)
            {
                result.ErrorMessage = $"Import error: {ex.Message}";
            }
            finally
            {
                if (sourceDoc != null)
                {
                    try
                    {
                        sourceDoc.Close(SaveChanges: false);
                        Marshal.ReleaseComObject(sourceDoc);
                    }
                    catch (COMException)
                    {
                        // Ignore cleanup errors
                    }
                }
            }

            return result;
        }

        public ExportResult ExportBuildingBlock(string buildingBlockName, string category, string outputFilePath)
        {
            var result = new ExportResult
            {
                BuildingBlockName = buildingBlockName,
                FilePath = outputFilePath
            };

            if (_templateDoc == null)
            {
                result.ErrorMessage = "No template document is open";
                return result;
            }

            Document newDoc = null;

            try
            {
                // Find the Building Block
                BuildingBlock bb = null;
                var attachedTemplate = (Template)_templateDoc.get_AttachedTemplate();
                foreach (BuildingBlock block in attachedTemplate.BuildingBlockEntries)
                {
                    if (block.Name == buildingBlockName && block.Category.ToString() == category)
                    {
                        bb = block;
                        break;
                    }
                }

                if (bb == null)
                {
                    result.ErrorMessage = $"Building Block '{buildingBlockName}' not found in category '{category}'";
                    return result;
                }

                // Create new document and insert Building Block content
                newDoc = _wordApp.Documents.Add();
                bb.Insert(newDoc.Content, true);

                // Ensure output directory exists
                var outputDir = Path.GetDirectoryName(outputFilePath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the document
                newDoc.SaveAs2(outputFilePath, FileFormat: WdSaveFormat.wdFormatXMLDocument);
                
                result.Success = true;
            }
            catch (COMException ex)
            {
                result.ErrorMessage = $"Word automation error: {ex.Message}";
            }
            catch (Exception ex)
            {
                result.ErrorMessage = $"Export error: {ex.Message}";
            }
            finally
            {
                if (newDoc != null)
                {
                    try
                    {
                        newDoc.Close(SaveChanges: false);
                        Marshal.ReleaseComObject(newDoc);
                    }
                    catch (COMException)
                    {
                        // Ignore cleanup errors
                    }
                }
            }

            return result;
        }

        public List<BuildingBlockInfo> GetBuildingBlocks(string categoryPrefix = "InternalAutotext")
        {
            var buildingBlocks = new List<BuildingBlockInfo>();

            if (_templateDoc == null)
                return buildingBlocks;

            try
            {
                var attachedTemplate = (Template)_templateDoc.get_AttachedTemplate();
                foreach (BuildingBlock bb in attachedTemplate.BuildingBlockEntries)
                {
                    var categoryString = bb.Category.ToString();
                    if (string.IsNullOrEmpty(categoryPrefix) || categoryString.StartsWith(categoryPrefix))
                    {
                        buildingBlocks.Add(new BuildingBlockInfo
                        {
                            Name = bb.Name,
                            Category = categoryString,
                            Gallery = bb.Type.ToString(),
                            Description = bb.Description,
                            CreatedDate = DateTime.Now, // COM doesn't expose creation date
                            ModifiedDate = DateTime.Now  // COM doesn't expose modification date
                        });
                    }
                }
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException($"Failed to retrieve Building Blocks: {ex.Message}", ex);
            }

            return buildingBlocks;
        }

        public bool BuildingBlockExists(string name, string category)
        {
            if (_templateDoc == null)
                return false;

            try
            {
                var attachedTemplate = (Template)_templateDoc.get_AttachedTemplate();
                foreach (BuildingBlock bb in attachedTemplate.BuildingBlockEntries)
                {
                    if (bb.Name == name && bb.Category.ToString() == category)
                    {
                        return true;
                    }
                }
            }
            catch (COMException)
            {
                return false;
            }

            return false;
        }

        public bool DeleteBuildingBlock(string name, string category)
        {
            if (_templateDoc == null)
                return false;

            try
            {
                var attachedTemplate = (Template)_templateDoc.get_AttachedTemplate();
                foreach (BuildingBlock bb in attachedTemplate.BuildingBlockEntries)
                {
                    if (bb.Name == name && bb.Category.ToString() == category)
                    {
                        bb.Delete();
                        return true;
                    }
                }
            }
            catch (COMException)
            {
                return false;
            }

            return false;
        }

        public void SaveTemplate()
        {
            if (_templateDoc == null)
                return;

            try
            {
                _templateDoc.Save();
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException($"Failed to save template: {ex.Message}", ex);
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // Dispose managed resources
                }

                // Dispose unmanaged resources
                CloseTemplate();

                if (_wordApp != null)
                {
                    try
                    {
                        _wordApp.Quit(SaveChanges: false);
                        Marshal.ReleaseComObject(_wordApp);
                    }
                    catch (COMException)
                    {
                        // Ignore errors when closing Word
                    }
                    finally
                    {
                        _wordApp = null;
                    }
                }

                _disposed = true;
            }
        }

        ~WordManager()
        {
            Dispose(false);
        }
    }
}