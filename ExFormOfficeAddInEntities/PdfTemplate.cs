using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExFormOfficeAddInEntities
{
    public class PdfTemplate
    {
        public int? CompanyId { get; set; }
        public string TemplateName { get; set; }
        public string Description { get; set; }
        public byte[] TemplateFileZip { get; set; }
        public bool IsActive { get; set; }
        public DateTime? CreatedOn { get; set; }
        public int? CreatedBy { get; set; }
        public DateTime? UpdatedOn { get; set; }
        public int? UpdatedBy { get; set; }
        public byte[] ExcelZip { get; set; }
        public byte[] ExcelBytes { get; set; }
        public DataTable TemplateFile { get; set; }
        public DataTable TemplateFileFieldMapping { get; set; }
        public int TemplateFolderId { get; set; }
        public string FolderName { get; set; }
        public string SubFolderName { get; set; }
        public string FileNamePart { get; set; }
        public string ExcelVersion { get; set; }
        public string FileExtension { get; set; }
        public bool IsDemoTemplate { get; set; }
    }

    public class JTemplate
    {
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
        public byte[] ExcelZip { get; set; }
        public string[] TemplateFile { get; set; }
        public DataTable TemplateFileFieldMapping { get; set; }
        public int TemplateFolderId { get; set; }
        public string FolderName { get; set; }
        public string SubFolderName { get; set; }
        public string FileNamePart { get; set; }
        public string ExcelVersion { get; set; }
        public bool IsDemoSet { get; set; }
    }
    public class EditPdfTemplate
    {
        public string DeletedTemplateFileIds { get; set; }
        public int? UpdatedBy { get; set; }
        public DataTable Template { get; set; }
        public DataTable TemplateFile { get; set; }
        public DataTable TemplateFileFieldMapping { get; set; }
        public byte[] ExcelZip { get; set; }
        public byte[] TemplateFileZip { get; set; }
    }
    public class JEditPdfTemplate
    {
        public int TemplateId { get; set; }
        public string DeletedTemplateFileIds { get; set; }
        public string TemplateName { get; set; }
        public string Description { get; set; }
        public string SubFolderName { get; set; }
        public string FileNamePart { get; set; }
        public bool IsFileDelete { get; set; }
        public int? UpdatedBy { get; set; }        
        public byte[] ExcelZip { get; set; }
        public byte[] TemplateFileZip { get; set; }
    }
    public class MappingFields
    {
        public List<MappingField> ParentFields { get; set; }
        public List<MappingField> ChildFields { get; set; }
        public List<string> DynamicFieldIds { get; set; }
        public int TemplateId { get; set; }
    }
    public class MapFieldParam
    {
        public int TemplateFileMappingId { get; set; }
        public string SheetName { get; set; }
        public string TableName { get; set; }
        public string ColumnName { get; set; }
        public bool IsDynamicElement { get; set; }
    }
    public class EditMapFieldBackResponse
    {
        public bool IsAnyFiedMapped { get; set; }
        public int DynamicFieldsCount { get; set; }
        public string ParentTable { get; set; }
        public int ChildTablesCount { get; set; }
        public int ParentChildRelationshipCount { get; set; }
    }
    public class RangeParamResponse
    {        
        public string SheetName { get; set; }
        public string TableName { get; set; }
        public string ColumnName { get; set; }
    }
    public class MappingField
    {
        public int TemplateFileMappingId { get; set; }
        public string ExcelFieldName { get; set; }
        public string SheetName { get; set; }
        public string ExcelTableName { get; set; }
        public bool IsMapped { get; set; }
        public string ParentFieldId { get; set; }
    }
    public class ParentChildMapping
    {
        public int TemplateFileId { get; set; }
        public string ParentTable { get; set; }
        public string ParentTableField { get; set; }
        public string ChildTableField { get; set; }
        public string ChildTable { get; set; }
    }
    public class ColumnValue
    {
        public int TemplateId { get; set; }
        public string OldValue { get; set; }
        public string NewValue { get; set; }
    }
    public class NamingOption
    {
        public int TemplateId { get; set; }
        public string Fields { get; set; }
        public int UpdatedBy { get; set; }
    }
    public class ExcelFile
    {
        public int TemplateId { get; set; }
        public byte[] fileBytes { get; set; }
    }
    public class DynamicParam
    {
        public int TemplateFileMappingId { get; set; }
        public bool IsDynamicField { get; set; }
    }
    public class ExcelVersion
    {
        public int TemplateId { get; set; }
        public int UserId { get; set; }
        public string ExcelVersionId { get; set; }
        public byte[] fileBytes { get; set; }
        public string FileExtension { get; set; }
    }
    public class CreateTemplateResponse
    {
        public int TemplateId { get; set; }
        public string Error { get; set; }
    }
    public class SendDataParamResponse
    {
        public DataTable Params { get; set; }
        public string Error { get; set; }
    }

    public class IsExcelVersionExistResponse
    {
        public bool IsExcelVersionExist { get; set; }
        public string Error { get; set; }
    }
    public class NamingOptions
    {
        public DataTable Names { get; set; }
        public string Error { get; set; }
    }
    public class SendToTemplateSetData
    {
        public List<string> ParentTableData { get; set; }
        public List<List<string>> ChildTableData { get; set; }
        public int TemplateId { get; set; }
        public string ParentTableName { get; set; }
        public List<string> ChildTableNames { get; set; }
        public int UserId { get; set; }
    }

    public class SendToTemplateSetDataResponse
    {        
        public string Error { get; set; }
        public List<string> Message { get; set; }
        public string ZipPath { get; set; }
    }
}
