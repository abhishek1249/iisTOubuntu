using System;
using System.Collections.Generic;

namespace ExFormOfficeAddInEntities
{
    public class TemplateInfo
    {
        public int TemplateId { get; set; }
        public int? CompanyId { get; set; }
        public string TemplateName { get; set; }
        public string Description { get; set; }        
        public bool IsActive { get; set; }
        public DateTime? CreatedOn { get; set; }
        public int? CreatedBy { get; set; }
        public DateTime? UpdatedOn { get; set; }
        public int? UpdatedBy { get; set; }
        public string SubFolderName { get; set; }
        public string FileNamePart { get; set; }
        public List<TemplateFile> Files { get; set; }
    }
    public class TemplateFile
    {
        public int TemplateFileId { get; set; }
        public int DynamicFieldsCount { get; set; }
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public bool IsXFA { get; set; }
        public string MappedPercentage { get; set; }
        public byte[] PdfBytes { get; set; }
    }
    public class ExcelVersionResponse
    {
        public string ExcelVersion { get; set; }
        public string UpdatedBy { get; set; }
        public string UpdatedOn { get; set; }
        public string Error { get; set; }
    }
    public class ExcelFileInfo
    {
        public byte[] ExcelBytes { get; set; }
        public string TemplateName { get; set; }
        public string FileExtension { get; set; }
    }
}
