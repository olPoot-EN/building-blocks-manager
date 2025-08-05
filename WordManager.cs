using System;
using System.Collections.Generic;
using System.IO;

namespace BuildingBlocksManager
{
    public class WordManager
    {
        private string templatePath;

        public WordManager(string templatePath)
        {
            this.templatePath = templatePath;
        }

        public void CreateBackup()
        {
            // TODO: Implement backup creation
            // Format: TemplateName_Backup_YYYYMMDD_HHMMSS.dotm
        }

        public List<string> GetBuildingBlocks()
        {
            // TODO: Implement Building Block enumeration
            // Return list of Building Blocks from InternalAutotext category
            return new List<string>();
        }

        public void ImportBuildingBlock(string sourceFile, string category, string name)
        {
            // TODO: Implement Building Block import
            // 1. Open source document
            // 2. Extract formatted content
            // 3. Create/update Building Block in template
            // 4. Save template
        }

        public void ExportBuildingBlock(string buildingBlockName, string outputPath)
        {
            // TODO: Implement Building Block export
            // 1. Find Building Block in template
            // 2. Create new document with Building Block content
            // 3. Save as AT_[name].docx
        }

        public void RollbackFromBackup()
        {
            // TODO: Implement rollback functionality
            // 1. Find most recent backup
            // 2. Replace current template with backup
        }

        public void Dispose()
        {
            // TODO: Implement proper COM object disposal
        }
    }
}