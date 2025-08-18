using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace BuildingBlocksManager
{
    public class FileManager
    {
        public class FileInfo
        {
            public string FilePath { get; set; }
            public DateTime LastModified { get; set; }
            public DateTime LastImported { get; set; }
            public bool IsNew => LastImported == DateTime.MinValue;
            public bool IsModified => LastModified > LastImported;
            public string Category { get; set; }
            public string Name { get; set; }
            public List<string> InvalidCharacters { get; set; } = new List<string>();
            public bool IsValid => InvalidCharacters.Count == 0;
        }

        private string sourceDirectory;
        private BuildingBlockLedger ledger;

        // Characters that are invalid in Word Building Block names
        private static readonly char[] InvalidChars = { '/', '\\', ':', '*', '?', '"', '<', '>', '|' };

        public FileManager(string sourceDirectory, string logDirectory = null)
        {
            this.sourceDirectory = sourceDirectory;
            this.ledger = string.IsNullOrEmpty(logDirectory) 
                ? new BuildingBlockLedger() 
                : new BuildingBlockLedger(logDirectory);
        }

        public List<FileInfo> ScanDirectory()
        {
            var files = new List<FileInfo>();

            if (!Directory.Exists(sourceDirectory))
                throw new DirectoryNotFoundException($"Source directory not found: {sourceDirectory}");

            try
            {
                // Recursively scan for AT_*.docx files (up to 5 levels deep)
                var docxFiles = GetDocxFiles(sourceDirectory, 0, 5);

                foreach (var filePath in docxFiles)
                {
                    var fileName = Path.GetFileName(filePath);
                    
                    // Only process files that start with AT_
                    if (!fileName.StartsWith("AT_", StringComparison.OrdinalIgnoreCase))
                        continue;

                    var fileInfo = new System.IO.FileInfo(filePath);
                    var name = ExtractName(fileName);
                    var category = ExtractCategory(filePath);
                    var lastImported = ledger.GetLastImportTime(name, category);
                    
                    // DEBUG: Show debug info for first few files
                    if (files.Count < 3)
                    {
                        var debugInfo = $"[SCAN] File: {fileName}\n" +
                                      $"Extracted Name: '{name}'\n" +
                                      $"Extracted Category: '{category}'\n" +
                                      $"Ledger Key: '{name}|{category ?? ""}'\n" +
                                      $"LastImported: {(lastImported == DateTime.MinValue ? "DateTime.MinValue (NEW)" : lastImported.ToString())}\n" +
                                      $"File LastModified: {fileInfo.LastWriteTime}\n" +
                                      $"IsNew: {lastImported == DateTime.MinValue}";
                        
                        System.Windows.Forms.MessageBox.Show(debugInfo, $"Debug Info - File {files.Count + 1}");
                    }
                    
                    var fileData = new FileInfo
                    {
                        FilePath = filePath,
                        LastModified = fileInfo.LastWriteTime,
                        LastImported = lastImported,
                        Category = category,
                        Name = name,
                        InvalidCharacters = GetInvalidCharacters(fileName)
                    };

                    files.Add(fileData);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to scan directory: {ex.Message}", ex);
            }

            return files.OrderBy(f => f.FilePath).ToList();
        }

        private List<string> GetDocxFiles(string directory, int currentDepth, int maxDepth)
        {
            var files = new List<string>();

            if (currentDepth > maxDepth)
                return files;

            try
            {
                // Add .docx files from current directory
                files.AddRange(Directory.GetFiles(directory, "*.docx"));

                // Recursively scan subdirectories
                if (currentDepth < maxDepth)
                {
                    foreach (var subDir in Directory.GetDirectories(directory))
                    {
                        files.AddRange(GetDocxFiles(subDir, currentDepth + 1, maxDepth));
                    }
                }
            }
            catch (UnauthorizedAccessException)
            {
                // Skip directories we don't have access to
            }

            return files;
        }

        public string ExtractCategory(string filePath)
        {
            try
            {
                var relativePath = GetRelativePath(sourceDirectory, filePath);
                var directory = Path.GetDirectoryName(relativePath);

                if (string.IsNullOrEmpty(directory) || directory == ".")
                {
                    // File is in root directory - no category
                    return "";
                }

                // Convert folder path to category
                // Replace path separators with backslashes and spaces with underscores
                var category = directory.Replace(Path.DirectorySeparatorChar, '\\')
                                      .Replace(Path.AltDirectorySeparatorChar, '\\');
                
                // Convert spaces to underscores in folder names
                var parts = category.Split('\\');
                for (int i = 0; i < parts.Length; i++)
                {
                    parts[i] = parts[i].Replace(' ', '_');
                }
                category = string.Join("\\", parts);

                return category;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to extract category from path '{filePath}': {ex.Message}", ex);
            }
        }

        public string ExtractName(string fileName)
        {
            try
            {
                if (!fileName.StartsWith("AT_", StringComparison.OrdinalIgnoreCase))
                    throw new ArgumentException($"File name must start with 'AT_': {fileName}");

                // Remove AT_ prefix and .docx extension
                var name = fileName.Substring(3); // Remove "AT_"
                
                if (name.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                {
                    name = name.Substring(0, name.Length - 5); // Remove ".docx"
                }

                // Convert spaces to underscores
                name = name.Replace(' ', '_');

                return name;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to extract name from file '{fileName}': {ex.Message}", ex);
            }
        }

        public bool IsValidFileName(string fileName)
        {
            return GetInvalidCharacters(fileName).Count == 0;
        }

        public List<string> GetInvalidCharacters(string fileName)
        {
            var invalid = new List<string>();

            foreach (char c in InvalidChars)
            {
                if (fileName.Contains(c))
                {
                    invalid.Add(c.ToString());
                }
            }

            return invalid.Distinct().ToList();
        }

        public List<FileInfo> GetNewFiles()
        {
            return ScanDirectory().Where(f => f.IsNew).ToList();
        }

        public List<FileInfo> GetModifiedFiles()
        {
            return ScanDirectory().Where(f => f.IsModified && !f.IsNew).ToList();
        }

        public List<FileInfo> GetUpToDateFiles()
        {
            return ScanDirectory().Where(f => !f.IsNew && !f.IsModified).ToList();
        }

        public List<FileInfo> GetInvalidFiles()
        {
            return ScanDirectory().Where(f => !f.IsValid).ToList();
        }

        public int GetIgnoredFileCount()
        {
            try
            {
                var allDocxFiles = GetDocxFiles(sourceDirectory, 0, 5);
                var atFiles = allDocxFiles.Where(f => Path.GetFileName(f).StartsWith("AT_", StringComparison.OrdinalIgnoreCase)).Count();
                return allDocxFiles.Count - atFiles;
            }
            catch
            {
                return 0;
            }
        }

        public string GetSummary()
        {
            try
            {
                var allFiles = ScanDirectory();
                var newFiles = allFiles.Where(f => f.IsNew).Count();
                var modifiedFiles = allFiles.Where(f => f.IsModified && !f.IsNew).Count();
                var upToDateFiles = allFiles.Where(f => !f.IsNew && !f.IsModified).Count();
                var invalidFiles = allFiles.Where(f => !f.IsValid).Count();
                var ignoredFiles = GetIgnoredFileCount();

                return $"Files Ready for Import: {newFiles + modifiedFiles}\n" +
                       $"- New Files: {newFiles}\n" +
                       $"- Modified Files: {modifiedFiles}\n" +
                       $"- Up-to-date Files: {upToDateFiles}\n" +
                       $"- Invalid Files: {invalidFiles}\n" +
                       $"- Ignored Files: {ignoredFiles}";
            }
            catch (Exception ex)
            {
                return $"Error scanning directory: {ex.Message}";
            }
        }

        private string GetRelativePath(string fromPath, string toPath)
        {
            if (string.IsNullOrEmpty(fromPath)) return toPath;
            if (string.IsNullOrEmpty(toPath)) return string.Empty;

            Uri fromUri = new Uri(AppendDirectorySeparatorChar(fromPath));
            Uri toUri = new Uri(toPath);

            if (fromUri.Scheme != toUri.Scheme) return toPath;

            Uri relativeUri = fromUri.MakeRelativeUri(toUri);
            string relativePath = Uri.UnescapeDataString(relativeUri.ToString());

            if (string.Equals(toUri.Scheme, Uri.UriSchemeFile, StringComparison.OrdinalIgnoreCase))
            {
                relativePath = relativePath.Replace('/', Path.DirectorySeparatorChar);
            }

            return relativePath;
        }

        private string AppendDirectorySeparatorChar(string path)
        {
            if (!Path.HasExtension(path) && !path.EndsWith(Path.DirectorySeparatorChar.ToString()))
            {
                return path + Path.DirectorySeparatorChar;
            }
            return path;
        }
    }
}