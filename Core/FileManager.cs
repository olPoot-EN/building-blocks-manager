using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace BuildingBlocksManager.Core
{
    public class FileManager
    {
        private static readonly char[] InvalidFileNameChars = { '/', '\\', ':', '*', '?', '"', '<', '>', '|' };
        private const int MaxDirectoryDepth = 5;
        private const string FilePrefix = "AT_";
        private const string FileExtension = ".docx";

        public class FileInfo
        {
            public string FullPath { get; set; }
            public string RelativePath { get; set; }
            public string FileName { get; set; }
            public string BuildingBlockName { get; set; }
            public string Category { get; set; }
            public DateTime LastModified { get; set; }
            public bool IsValid { get; set; }
            public string ValidationError { get; set; }
        }

        public class ScanResult
        {
            public List<FileInfo> ValidFiles { get; set; } = new List<FileInfo>();
            public List<FileInfo> InvalidFiles { get; set; } = new List<FileInfo>();
            public List<string> IgnoredFiles { get; set; } = new List<string>();
            public int TotalFilesScanned { get; set; }
            public TimeSpan ScanDuration { get; set; }
        }

        public ScanResult ScanDirectory(string rootPath)
        {
            var startTime = DateTime.Now;
            var result = new ScanResult();

            if (!Directory.Exists(rootPath))
            {
                throw new DirectoryNotFoundException($"Directory not found: {rootPath}");
            }

            try
            {
                var allFiles = GetAllDocxFiles(rootPath);
                result.TotalFilesScanned = allFiles.Count;

                foreach (var filePath in allFiles)
                {
                    var fileName = Path.GetFileName(filePath);
                    
                    if (!fileName.StartsWith(FilePrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        result.IgnoredFiles.Add(filePath);
                        continue;
                    }

                    var fileInfo = CreateFileInfo(filePath, rootPath);
                    
                    if (fileInfo.IsValid)
                    {
                        result.ValidFiles.Add(fileInfo);
                    }
                    else
                    {
                        result.InvalidFiles.Add(fileInfo);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error scanning directory: {ex.Message}", ex);
            }

            result.ScanDuration = DateTime.Now - startTime;
            return result;
        }

        private List<string> GetAllDocxFiles(string rootPath)
        {
            var files = new List<string>();
            
            try
            {
                GetDocxFilesRecursive(rootPath, rootPath, files, 0);
            }
            catch (UnauthorizedAccessException)
            {
                // Log and continue - some directories may be inaccessible
            }

            return files;
        }

        private void GetDocxFilesRecursive(string currentPath, string rootPath, List<string> files, int currentDepth)
        {
            if (currentDepth >= MaxDirectoryDepth)
                return;

            try
            {
                // Get .docx files in current directory
                var docxFiles = Directory.GetFiles(currentPath, "*" + FileExtension, SearchOption.TopDirectoryOnly);
                files.AddRange(docxFiles);

                // Recursively scan subdirectories
                var subdirectories = Directory.GetDirectories(currentPath);
                foreach (var subdirectory in subdirectories)
                {
                    GetDocxFilesRecursive(subdirectory, rootPath, files, currentDepth + 1);
                }
            }
            catch (UnauthorizedAccessException)
            {
                // Skip directories we can't access
            }
            catch (DirectoryNotFoundException)
            {
                // Skip directories that no longer exist
            }
        }

        private FileInfo CreateFileInfo(string fullPath, string rootPath)
        {
            var fileName = Path.GetFileName(fullPath);
            var fileInfo = new FileInfo
            {
                FullPath = fullPath,
                RelativePath = Path.GetRelativePath(rootPath, fullPath),
                FileName = fileName,
                LastModified = File.GetLastWriteTime(fullPath)
            };

            // Validate filename
            var validation = ValidateFileName(fileName);
            fileInfo.IsValid = validation.IsValid;
            fileInfo.ValidationError = validation.Error;

            if (fileInfo.IsValid)
            {
                // Extract Building Block name by removing AT_ prefix and .docx extension
                var nameWithoutPrefix = fileName.Substring(FilePrefix.Length);
                fileInfo.BuildingBlockName = Path.GetFileNameWithoutExtension(nameWithoutPrefix);
                
                // Convert spaces to underscores in Building Block name
                fileInfo.BuildingBlockName = fileInfo.BuildingBlockName.Replace(' ', '_');

                // Generate category from directory structure
                fileInfo.Category = GenerateCategory(fullPath, rootPath);
            }

            return fileInfo;
        }

        private (bool IsValid, string Error) ValidateFileName(string fileName)
        {
            // Check for invalid characters
            var invalidChars = fileName.Where(c => InvalidFileNameChars.Contains(c)).ToList();
            if (invalidChars.Any())
            {
                var invalidCharsString = string.Join(", ", invalidChars.Distinct());
                return (false, $"Contains invalid characters: {invalidCharsString}");
            }

            // Check if it starts with AT_ and ends with .docx
            if (!fileName.StartsWith(FilePrefix, StringComparison.OrdinalIgnoreCase))
            {
                return (false, "Filename must start with 'AT_'");
            }

            if (!fileName.EndsWith(FileExtension, StringComparison.OrdinalIgnoreCase))
            {
                return (false, "File must have .docx extension");
            }

            // Check if there's a name between prefix and extension
            var nameWithoutPrefix = fileName.Substring(FilePrefix.Length);
            var nameWithoutExtension = Path.GetFileNameWithoutExtension(nameWithoutPrefix);
            
            if (string.IsNullOrWhiteSpace(nameWithoutExtension))
            {
                return (false, "Filename must contain a name after 'AT_' prefix");
            }

            return (true, null);
        }

        private string GenerateCategory(string fullPath, string rootPath)
        {
            var relativePath = Path.GetRelativePath(rootPath, fullPath);
            var directoryPath = Path.GetDirectoryName(relativePath);

            if (string.IsNullOrEmpty(directoryPath) || directoryPath == ".")
            {
                // File is in root directory
                return "InternalAutotext";
            }

            // Convert directory path to category format
            // Replace path separators with backslashes and spaces with underscores
            var category = directoryPath.Replace('/', '\\');
            var categoryParts = category.Split('\\');
            
            // Convert spaces to underscores in each part
            for (int i = 0; i < categoryParts.Length; i++)
            {
                categoryParts[i] = categoryParts[i].Replace(' ', '_');
            }

            return "InternalAutotext\\" + string.Join("\\", categoryParts);
        }

        public bool IsValidDocxFile(string filePath)
        {
            if (!File.Exists(filePath))
                return false;

            try
            {
                // Basic check - try to access file properties
                var fileInfo = new System.IO.FileInfo(filePath);
                return fileInfo.Length > 0 && fileInfo.Extension.Equals(".docx", StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        public string SanitizePathForWindows(string path)
        {
            if (string.IsNullOrEmpty(path))
                return path;

            // Replace forward slashes with backslashes for Windows compatibility
            return path.Replace('/', '\\');
        }

        public List<string> GetInvalidCharacterFiles(IEnumerable<FileInfo> files)
        {
            return files.Where(f => !f.IsValid && f.ValidationError?.Contains("invalid characters") == true)
                       .Select(f => f.FileName)
                       .ToList();
        }
    }
}