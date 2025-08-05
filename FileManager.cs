using System;
using System.Collections.Generic;
using System.IO;

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
        }

        private string sourceDirectory;

        public FileManager(string sourceDirectory)
        {
            this.sourceDirectory = sourceDirectory;
        }

        public List<FileInfo> ScanDirectory()
        {
            // TODO: Implement directory scanning
            // 1. Recursively scan for AT_*.docx files (up to 5 levels)
            // 2. Extract category from folder path
            // 3. Extract name from filename
            // 4. Compare with import tracking data
            return new List<FileInfo>();
        }

        public string ExtractCategory(string filePath)
        {
            // TODO: Implement category extraction
            // Convert folder path to Building Block category
            // Example: Legal\Contracts\AT_Standard.docx -> InternalAutotext\Legal\Contracts
            return "";
        }

        public string ExtractName(string fileName)
        {
            // TODO: Implement name extraction
            // Remove AT_ prefix and .docx extension
            // Example: AT_Standard.docx -> Standard
            return "";
        }

        public bool IsValidFileName(string fileName)
        {
            // TODO: Implement filename validation
            // Check for special characters that Word doesn't support
            return true;
        }

        public List<string> GetInvalidCharacters(string fileName)
        {
            // TODO: Return list of invalid characters found in filename
            return new List<string>();
        }
    }
}