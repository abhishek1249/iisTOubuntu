using System;
using System.Collections.Generic;
using System.Text;

namespace ExFormOfficeAddInEntities
{
    public class TemplateFileFieldsMapping
    {
        public int TemplateFileMappingId { get; set; }
        public int TemplateId { get; set; }
        public int TemplateFileId { get; set; }
        public string PDFFieldName { get; set; }        
        public string ExcelFieldName { get; set; }
        public bool IsMapped { get; set; }
        public string FilePath { get; set; }
        public string SheetName { get; set; }
        public string ExcelTableName { get; set; }
        public string FieldId { get; set; }
        public string ParentFieldId { get; set; }
        public string XPath { get; set; }
        public bool IsDynamic { get; set; }
        public bool HasChildFields { get; set; }
    }
}
