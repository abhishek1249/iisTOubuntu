using System;
using System.Collections.Generic;

namespace ExFormOfficeAddInEntities
{
    public class Template
    {
        public int TemplateId { get; set; }
        public int? CompanyId { get; set; }
        public int? TeamId { get; set; }
        public string TemplateName { get; set; }
        public string Description { get; set; }
        public byte[] TemplateFileZip { get; set; }
        public bool IsActive { get; set; }
        public DateTime? CreatedOn { get; set; }
        public int? CreatedBy { get; set; }
        public DateTime? UpdatedOn { get; set; }
        public int? UpdatedBy { get; set; }
        public int PdfCount { get; set; }
        public bool IsDemoTemplate { get; set; }
    }
    public class DuplicateTemplateParam
    {
        public string Id { get; set; }
        public string CompanyId { get; set; }
        public string UserId { get; set; }
        public string TeamId { get; set; }
    }
    public class ParentChildTable
    {
        public string ParentTabel { get; set; }
        public string ChildTabel { get; set; }
        public bool IsMapped { get; set; }
    }
    public class Pdf
    {
        public Pdf()
        { }
        public Pdf(string fileName, string folder)
        {
            FileName = fileName; Folder = folder;
        }
        public string FileName { get; set; }
        public string Folder { get; set; }
        public bool IsXFA { get; set; }
    }
    public class PdfFile
    {
        public int? TemplateFileId { get; set; }
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public bool IsXFA { get; set; }
    }
    public class ParentChildFields
    {       
        public List<string> ChildFields { get; set; }
        public List<string> ParentFields { get; set; }
        public string ParentField { get; set; }
        public string ChildField { get; set; }
    }
}
