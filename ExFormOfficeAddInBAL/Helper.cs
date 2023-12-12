using ExFormOfficeAddInDAL;
using ExFormOfficeAddInEntities;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace ExFormOfficeAddInBAL
{
    public class Helper
    {
        public static User LogIn(string userName, string password, string account)
        {
            var dt = new DataTable("User");
            User user = null;
            CommonSql CommonSql = new CommonSql();
            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlComm = new SqlCommand("usp_Login", conn))
                    {
                        sqlComm.Parameters.AddWithValue("@UserName", userName);
                        sqlComm.Parameters.AddWithValue("@Password", password);
                        sqlComm.Parameters.AddWithValue("@AccountName", account);
                        sqlComm.CommandType = CommandType.StoredProcedure;

                        using (var da = new SqlDataAdapter())
                        {
                            da.SelectCommand = sqlComm;
                            da.Fill(dt);
                        }
                    }
                }
                if (dt.Rows.Count == 1)
                {
                    user = new User();
                    user.UserId = Convert.ToInt32(dt.Rows[0]["UserId"]);
                    user.FullName = Convert.ToString(dt.Rows[0]["FullName"]);
                    user.UserName = Convert.ToString(dt.Rows[0]["UserName"]);
                    user.UserType = Convert.ToChar(dt.Rows[0]["UserType"]);
                    user.CompanyId = Convert.ToInt32(dt.Rows[0]["CompanyId"]);
                    user.CompanyName = Convert.ToString(dt.Rows[0]["CompanyName"]);
                    user.Email = Convert.ToString(dt.Rows[0]["Email"]);
                    user.IsActive = Convert.ToBoolean(dt.Rows[0]["IsActive"]);
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
            return user;

        }
        public static byte[] GetLicenseStream()
        {
            var dt = new DataTable();
            CommonSql CommonSql = new CommonSql();
            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlComm = new SqlCommand("usp_GetLicenseStream", conn))
                    {
                        sqlComm.CommandType = CommandType.StoredProcedure;

                        using (var da = new SqlDataAdapter())
                        {
                            da.SelectCommand = sqlComm;
                            da.Fill(dt);
                        }
                    }
                }
                if (dt.Rows.Count == 1)
                {
                    if (dt.Rows[0]["DecryptedLicense"] != DBNull.Value)
                        return (byte[])dt.Rows[0]["DecryptedLicense"];
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
            return null;

        }
        public static List<TemplateFolder> GetTemplateFolderByCompanyId(int companyId)
        {
            var dt = new DataTable();
            var lstTemplateFolder = new List<TemplateFolder>();
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlComm = new SqlCommand("usp_GetTemplateFolderByCompanyId", conn))
                    {
                        sqlComm.Parameters.AddWithValue("@CompanyId", companyId);
                        sqlComm.CommandType = CommandType.StoredProcedure;

                        using (var da = new SqlDataAdapter())
                        {
                            da.SelectCommand = sqlComm;
                            da.Fill(dt);
                        }
                    }
                }

                if (dt.Rows.Count > 0)
                {
                    var rows = dt.Rows;
                    var rowCount = dt.Rows.Count;
                    for (var i = 0; i < rowCount; i++)
                    {
                        var templateFoder = new TemplateFolder();
                        templateFoder.FolderId = Convert.ToString(rows[i]["FolderId"]);
                        templateFoder.FolderName = Convert.ToString(rows[i]["FolderName"]);
                        templateFoder.ParentFolderId = Convert.ToString(rows[i]["ParentFolderId"]);
                        lstTemplateFolder.Add(templateFoder);
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }

            return lstTemplateFolder;
        }
        public static TemplateFolder GetDemoFolder()
        {
            var dt = new DataTable();
            var templateFoder = new TemplateFolder();
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlComm = new SqlCommand("usp_GetDemoFolder", conn))
                    {
                        sqlComm.CommandType = CommandType.StoredProcedure;

                        using (var da = new SqlDataAdapter())
                        {
                            da.SelectCommand = sqlComm;
                            da.Fill(dt);
                        }
                    }
                }

                if (dt.Rows.Count > 0)
                {
                    var rows = dt.Rows;

                    templateFoder.FolderId = Convert.ToString(rows[0]["FolderId"]);
                    templateFoder.FolderName = Convert.ToString(rows[0]["FolderName"]);
                    templateFoder.ParentFolderId = Convert.ToString(rows[0]["ParentFolderId"]);
                    templateFoder.IsDemoFolder = true;
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }

            return templateFoder;
        }
        public static List<Template> GetTemplateByFolderId(int? folderId)
        {
            CommonSql CommonSql = new CommonSql();
            var dt = new DataTable("Template");
            var lstTemplate = new List<Template>();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlComm = new SqlCommand("usp_GetTemplateByFolderId", conn))
                    {
                        sqlComm.Parameters.AddWithValue("@FolderId", folderId);
                        sqlComm.CommandType = CommandType.StoredProcedure;

                        using (var da = new SqlDataAdapter())
                        {
                            da.SelectCommand = sqlComm;
                            da.Fill(dt);
                        }
                    }
                }
                if (dt.Rows.Count > 0)
                {
                    var rows = dt.Rows;
                    var rowCount = dt.Rows.Count;
                    for (var i = 0; i < rowCount; i++)
                    {
                        var template = new Template();
                        template.TemplateId = Convert.ToInt32(rows[i]["TemplateId"]);
                        template.TemplateName = Convert.ToString(rows[i]["TemplateName"]);
                        template.Description = Convert.ToString(rows[i]["Description"]);
                        template.PdfCount = Convert.ToInt32(rows[i]["PdfCount"]);
                        template.IsDemoTemplate = rows[i]["IsDemoTemplate"] == DBNull.Value ? false : Convert.ToBoolean(rows[i]["IsDemoTemplate"]);
                        lstTemplate.Add(template);
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }

            return lstTemplate;
        }

        public static int CreateTemplate(PdfTemplate pdfTemplate)
        {
            CommonSql CommonSql = new CommonSql();
            var templateId = -1;
            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_CreateTemplateSet", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@CompanyId", pdfTemplate.CompanyId);
                        sqlCmd.Parameters.AddWithValue("@TemplateName", pdfTemplate.TemplateName);
                        sqlCmd.Parameters.AddWithValue("@Description", pdfTemplate.Description);
                        sqlCmd.Parameters.AddWithValue("@TemplateFileZip", pdfTemplate.TemplateFileZip);
                        sqlCmd.Parameters.AddWithValue("@IsActive", pdfTemplate.IsActive);
                        sqlCmd.Parameters.AddWithValue("@CreatedOn", pdfTemplate.CreatedOn);
                        sqlCmd.Parameters.AddWithValue("@CreatedBy", pdfTemplate.CreatedBy);
                        sqlCmd.Parameters.AddWithValue("@ExcelZip", pdfTemplate.ExcelZip);
                        sqlCmd.Parameters.AddWithValue("@TemplateFile", pdfTemplate.TemplateFile);
                        sqlCmd.Parameters.AddWithValue("@TemplateFileFieldMapping", pdfTemplate.TemplateFileFieldMapping);
                        sqlCmd.Parameters.AddWithValue("@TemplateFolderId", pdfTemplate.TemplateFolderId);
                        sqlCmd.Parameters.AddWithValue("@FolderName", pdfTemplate.FolderName);
                        sqlCmd.Parameters.AddWithValue("@SubFolderName", pdfTemplate.SubFolderName);
                        sqlCmd.Parameters.AddWithValue("@FileNamePart", pdfTemplate.FileNamePart);
                        sqlCmd.Parameters.AddWithValue("@ExcelVersion", pdfTemplate.ExcelVersion);
                        sqlCmd.Parameters.AddWithValue("@IsDemoTemplate", pdfTemplate.IsDemoTemplate);
                        sqlCmd.Parameters.AddWithValue("@ExcelBytes", pdfTemplate.ExcelBytes);
                        sqlCmd.Parameters.AddWithValue("@FileExtension", pdfTemplate.FileExtension);

                        sqlCmd.Parameters.Add("@TemplateId", SqlDbType.Int);
                        sqlCmd.Parameters["@TemplateId"].Direction = ParameterDirection.Output;
                        sqlCmd.ExecuteNonQuery();
                        templateId = Convert.ToInt32(sqlCmd.Parameters["@TemplateId"].Value);
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
            return templateId;
        }

        public static void UploadExcelFile(ExcelFile excelFile)
        {
            CommonSql CommonSql = new CommonSql();
            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_UpdateExcelBytes", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@TemplateId", excelFile.TemplateId);
                        sqlCmd.Parameters.AddWithValue("@ExcelBytes", excelFile.fileBytes);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
        }
        public static void UpdateTemplate(EditPdfTemplate editPdfTemplate)
        {
            CommonSql CommonSql = new CommonSql();
            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_UpdateTemplateSet", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@DeletedTemplateFileIds", editPdfTemplate.DeletedTemplateFileIds);
                        if (editPdfTemplate.UpdatedBy.HasValue)
                            sqlCmd.Parameters.AddWithValue("@UpdatedBy", editPdfTemplate.UpdatedBy);
                        else
                            sqlCmd.Parameters.AddWithValue("@UpdatedBy", DBNull.Value);
                        sqlCmd.Parameters.AddWithValue("@Template", editPdfTemplate.Template);
                        sqlCmd.Parameters.AddWithValue("@TemplateFile", editPdfTemplate.TemplateFile);
                        sqlCmd.Parameters.AddWithValue("@TemplateFileFieldMapping", editPdfTemplate.TemplateFileFieldMapping);
                        sqlCmd.Parameters.AddWithValue("@TemplateFileZip", editPdfTemplate.TemplateFileZip);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
        }
        public static DataTable GetTemplateByTemplateId(int templateId)
        {
            CommonSql CommonSql = new CommonSql();
            var dt = new DataTable();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_GetTemplateById", conn))
                    {
                        sqlCmd.Parameters.AddWithValue("@TemplateId", templateId);
                        sqlCmd.CommandType = CommandType.StoredProcedure;

                        using (var da = new SqlDataAdapter())
                        {
                            da.SelectCommand = sqlCmd;
                            da.Fill(dt);
                        }
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
            return dt;
        }
        public static void DeleteFolder(int folderId)
        {
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_DeleteTemplateFolder", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@FolderId", folderId);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
        }
        public static void CreateTemplateFolder(int companyId, int? parentFolderId, string folderName)
        {
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_CreateTemplateFolder", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@CompanyId", companyId);

                        if (parentFolderId.HasValue)
                            sqlCmd.Parameters.AddWithValue("@ParentFolderId", parentFolderId);
                        else
                            sqlCmd.Parameters.AddWithValue("@ParentFolderId", DBNull.Value);

                        sqlCmd.Parameters.AddWithValue("@FolderName", folderName);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
        }

        public static void RenameFolder(int folderId, string folderName)
        {
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_UpdateTemplateFolderName", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@FolderId", folderId);
                        sqlCmd.Parameters.AddWithValue("@Name", folderName);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
        }
        public static void CreateTemplateCopy(int templateId, int companyId, int createdBy)
        {
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_CreateTemplateCopy", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@TemplateId", templateId);
                        sqlCmd.Parameters.AddWithValue("@CompanyId", companyId);
                        sqlCmd.Parameters.AddWithValue("@CreatedBy", createdBy);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
        }
        public static void RenameTemplate(int templateId, string templateName)
        {
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_UpdateTemplateName", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@TemplateId", templateId);
                        sqlCmd.Parameters.AddWithValue("@Name", templateName);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
        }

        public static void DeleteTemplate(int templateId)
        {
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_DeleteTemplate", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@TemplateId", templateId);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
        }

        public static void DeleteTemplateFile(int templateFileId)
        {
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_DeleteTemplateFile", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@TemplateFileId", templateFileId);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
        }

        public static TemplateInfo GetTemplateById(int templateId)
        {
            var dt = new DataTable();
            var dt1 = new DataTable();
            var templateInfo = new TemplateInfo();
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_GetTemplateById", conn))
                    {
                        sqlCmd.Parameters.AddWithValue("@TemplateId", templateId);
                        sqlCmd.CommandType = CommandType.StoredProcedure;

                        using (var da = new SqlDataAdapter())
                        {
                            da.SelectCommand = sqlCmd;
                            da.Fill(dt);
                        }
                    }
                }

                if (dt.Rows.Count > 0)
                {
                    var rows = dt.Rows;
                    var rowCount = dt.Rows.Count;
                    for (var i = 0; i < rowCount; i++)
                    {
                        templateInfo.TemplateId = Convert.ToInt32(rows[i]["TemplateId"]);
                        templateInfo.CompanyId = Convert.ToInt32(rows[i]["CompanyId"]);
                        templateInfo.TemplateName = Convert.ToString(rows[i]["TemplateName"]);
                        templateInfo.Description = Convert.ToString(rows[i]["Description"]);

                        templateInfo.SubFolderName = Convert.ToString(rows[i]["SubFolderName"]);
                        templateInfo.FileNamePart = Convert.ToString(rows[i]["FileNamePart"]);
                        templateInfo.IsActive = Convert.ToBoolean(rows[i]["IsActive"]);
                    }
                }

                CommonSql CommonSql1 = new CommonSql();
                using (var conn = CommonSql1.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_GetTemplateFiles", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@TemplateId", templateId);

                        using (var da = new SqlDataAdapter())
                        {
                            da.SelectCommand = sqlCmd;
                            da.Fill(dt1);
                        }
                    }
                }

                if (dt1.Rows.Count > 0)
                {
                    var rows = dt1.Rows;
                    var rowCount = dt1.Rows.Count;
                    templateInfo.Files = new List<TemplateFile>();

                    for (var i = 0; i < rowCount; i++)
                    {
                        var templateFile = new TemplateFile();
                        templateFile.TemplateFileId = Convert.ToInt32(rows[i]["TemplateFileId"]);
                        templateFile.MappedPercentage = GetTemplateFileFieldsMappedPercentage(templateFile.TemplateFileId);
                        templateFile.DynamicFieldsCount = GetDynamicFieldsCount(templateFile.TemplateFileId);
                        templateFile.FileName = Convert.ToString(rows[i]["FileName"]);
                        templateFile.FilePath = Convert.ToString(rows[i]["FilePath"]);
                        templateFile.IsXFA = Convert.ToBoolean(rows[i]["IsXFA"]);
                        templateInfo.Files.Add(templateFile);
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
            return templateInfo;
        }

        public static DataTable GetTemplateMappedFields(int templateId)
        {
            var dt = new DataTable();
            CommonSql CommonSql = new CommonSql();
            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_GetTemplateMappedFields", conn))
                    {
                        sqlCmd.Parameters.AddWithValue("@TemplateId", templateId);
                        sqlCmd.CommandType = CommandType.StoredProcedure;

                        using (var da = new SqlDataAdapter())
                        {
                            da.SelectCommand = sqlCmd;
                            da.Fill(dt);
                        }
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
            return dt;
        }

        public static List<TemplateFile> GetTemplateFiles(int templateId)
        {
            var dt = new DataTable();
            var templateFiles = new List<TemplateFile>();
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {

                    using (var sqlCmd = new SqlCommand("usp_GetTemplateFiles", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@TemplateId", templateId);

                        using (var da = new SqlDataAdapter())
                        {
                            da.SelectCommand = sqlCmd;
                            da.Fill(dt);
                        }
                    }
                }


                if (dt.Rows.Count > 0)
                {
                    var rows = dt.Rows;
                    var rowCount = dt.Rows.Count;

                    for (var i = 0; i < rowCount; i++)
                    {
                        var templateFile = new TemplateFile();
                        templateFile.TemplateFileId = Convert.ToInt32(rows[i]["TemplateFileId"]);
                        //templateFile.MappedPercentage = GetTemplateFileFieldsMappedPercentage(templateFile.TemplateFileId);
                        templateFile.FileName = Convert.ToString(rows[i]["FileName"]);
                        templateFile.FilePath = Convert.ToString(rows[i]["FilePath"]);
                        templateFile.IsXFA = Convert.ToBoolean(rows[i]["IsXFA"]);
                        if (rows[i]["PdfBytes"] != DBNull.Value)
                            templateFile.PdfBytes = (byte[])rows[i]["PdfBytes"];
                        templateFiles.Add(templateFile);
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
            return templateFiles;
        }

        public static IsExcelVersionExistResponse IsExcelVersionExist(string excelVersion)
        {
            var dt = new DataTable();
            var response = new IsExcelVersionExistResponse();
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {

                    using (var sqlCmd = new SqlCommand("usp_IsExcelVersionExist", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@ExcelVersion", excelVersion);

                        using (var da = new SqlDataAdapter())
                        {
                            da.SelectCommand = sqlCmd;
                            da.Fill(dt);
                        }
                    }
                }


                if (dt.Rows.Count > 0)
                {
                    response.IsExcelVersionExist = true;
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
            return response;
        }

        public static DataTable GetParentChildTableMapping(int templateFileId)
        {
            var dt = new DataTable();
            CommonSql CommonSql = new CommonSql();
            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_GetParentChildTableMapping", conn))
                    {
                        sqlCmd.Parameters.AddWithValue("@TemplateFileId", templateFileId);
                        sqlCmd.CommandType = CommandType.StoredProcedure;

                        using (var da = new SqlDataAdapter())
                        {
                            da.SelectCommand = sqlCmd;
                            da.Fill(dt);
                        }
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
            return dt;
        }
        public static string GetTemplateFileFieldsMappedPercentage(int templateFileId)
        {
            CommonSql CommonSql = new CommonSql();
            var mappedPercentage = "0.00%";

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_TemplateFileFieldsMappedPercentage", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@TemplateFileId", templateFileId);

                        sqlCmd.Parameters.Add("@Percentage", SqlDbType.Float);
                        sqlCmd.Parameters["@Percentage"].Direction = ParameterDirection.Output;
                        sqlCmd.ExecuteNonQuery();
                        mappedPercentage = Convert.ToString(sqlCmd.Parameters["@Percentage"].Value) + "%";
                        if (mappedPercentage.Length > 4)
                            mappedPercentage = mappedPercentage.Substring(0, 4) + "%";
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
            return mappedPercentage;
        }

        public static List<TemplateFileFieldsMapping> GetTemplateFieldsByTemplateFileId(int templateFileId, bool isAdmin)
        {
            var dt = new DataTable();
            List<TemplateFileFieldsMapping> lstTemplateFileFieldsMapping = new List<TemplateFileFieldsMapping>();
            CommonSql CommonSql = new CommonSql();
            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_GetTemplateFieldsByTemplateFileId", conn))
                    {
                        sqlCmd.Parameters.AddWithValue("@TemplateFileId", templateFileId);
                        sqlCmd.Parameters.AddWithValue("@IsAdmin", isAdmin);
                        sqlCmd.CommandType = CommandType.StoredProcedure;

                        using (var da = new SqlDataAdapter())
                        {
                            da.SelectCommand = sqlCmd;
                            da.Fill(dt);
                        }
                    }
                }
                if (dt.Rows.Count > 0)
                {
                    var rows = dt.Rows;
                    var rowCount = dt.Rows.Count;
                    for (var i = 0; i < rowCount; i++)
                    {
                        var templateFileFieldsMapping = new TemplateFileFieldsMapping();
                        templateFileFieldsMapping.TemplateFileMappingId = Convert.ToInt32(rows[i]["TemplateFileMappingId"]);
                        templateFileFieldsMapping.TemplateId = Convert.ToInt32(rows[i]["TemplateId"]);
                        templateFileFieldsMapping.TemplateFileId = Convert.ToInt32(rows[i]["TemplateFileId"]);
                        templateFileFieldsMapping.PDFFieldName = Convert.ToString(rows[i]["PDFFieldName"]);
                        templateFileFieldsMapping.ExcelFieldName = Convert.ToString(rows[i]["ExcelFieldName"]);
                        templateFileFieldsMapping.XPath = Convert.ToString(rows[i]["XPath"]);
                        templateFileFieldsMapping.IsMapped = Convert.ToBoolean(rows[i]["IsMapped"]);
                        templateFileFieldsMapping.FilePath = Convert.ToString(rows[i]["FilePath"]);
                        templateFileFieldsMapping.SheetName = Convert.ToString(rows[i]["SheetName"]);
                        templateFileFieldsMapping.ExcelTableName = Convert.ToString(rows[i]["ExcelTableName"]);
                        templateFileFieldsMapping.FieldId = Convert.ToString(rows[i]["FieldId"]);
                        templateFileFieldsMapping.ParentFieldId = Convert.ToString(rows[i]["ParentFieldId"]);
                        templateFileFieldsMapping.IsDynamic = Convert.ToBoolean(rows[i]["IsDynamic"]);
                        templateFileFieldsMapping.HasChildFields = Convert.ToBoolean(rows[i]["HasChildFields"]);

                        lstTemplateFileFieldsMapping.Add(templateFileFieldsMapping);
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
            return lstTemplateFileFieldsMapping;
        }

        public static bool IsTemplateFieldsMapped(int templateId)
        {
            bool isMapped = false;
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_IsTemplateFieldsMapped", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@TemplateId", templateId);

                        sqlCmd.Parameters.Add("@IsMapped", SqlDbType.Bit);
                        sqlCmd.Parameters["@IsMapped"].Direction = ParameterDirection.Output;
                        sqlCmd.ExecuteNonQuery();
                        isMapped = Convert.ToBoolean(sqlCmd.Parameters["@IsMapped"].Value);
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
            return isMapped;
        }

        public static void RemoveAllMappedFields(int templateId)
        {
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_RemoveAllMappedFields", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@TemplateId", templateId);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
        }
        public static int GetDynamicFieldsCount(int templateFileId)
        {
            var count = 0;
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_GetDynamicFieldsCount", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@TemplateFileId", templateFileId);

                        sqlCmd.Parameters.Add("@Count", SqlDbType.Int);
                        sqlCmd.Parameters["@Count"].Direction = ParameterDirection.Output;
                        sqlCmd.ExecuteNonQuery();
                        count = Convert.ToInt32(sqlCmd.Parameters["@Count"].Value);
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
            return count;
        }

        public static int GetParentChildTableRelationshipCount(int templateFileId)
        {
            var count = 0;
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_GetParentChildTableRelationshipCount", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@TemplateFileId", templateFileId);

                        sqlCmd.Parameters.Add("@Count", SqlDbType.Int);
                        sqlCmd.Parameters["@Count"].Direction = ParameterDirection.Output;
                        sqlCmd.ExecuteNonQuery();
                        count = Convert.ToInt32(sqlCmd.Parameters["@Count"].Value);
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
            return count;
        }
        public static bool IsTemplateFileFieldsMapped(int templateFileId)
        {
            bool isMapped = false;
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_IsTemplateFileFieldsMapped", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@TemplateFileId", templateFileId);

                        sqlCmd.Parameters.Add("@IsMapped", SqlDbType.Bit);
                        sqlCmd.Parameters["@IsMapped"].Direction = ParameterDirection.Output;
                        sqlCmd.ExecuteNonQuery();
                        isMapped = Convert.ToBoolean(sqlCmd.Parameters["@IsMapped"].Value);
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
            return isMapped;
        }
        public static void RemoveFileMappedFields(int templateFileId)
        {
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_RemoveFileMappedFields", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@TemplateFileId", templateFileId);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
        }
        public static DataTable GetTemplateFieldsByTemplateId(int templateId)
        {
            var dt = new DataTable();
            var fieldId = string.Empty;
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_GetTemplateFieldsByTemplateId", conn))
                    {
                        sqlCmd.Parameters.AddWithValue("@TemplateId", templateId);
                        sqlCmd.CommandType = CommandType.StoredProcedure;

                        using (var da = new SqlDataAdapter())
                        {
                            da.SelectCommand = sqlCmd;
                            da.Fill(dt);
                        }
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
            return dt;
        }
        public static void MapParentTableFields(DataTable parentField)
        {
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_MapParentTableFields", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@ParentField", parentField);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
        }
        public static DataTable GetExcelVersion(int templateId)
        {
            var dt = new DataTable();
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_GetExcelVersion", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@TemplateId", templateId);

                        using (var da = new SqlDataAdapter())
                        {
                            da.SelectCommand = sqlCmd;
                            da.Fill(dt);
                        }
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
            return dt;
        }
        public static void UpdateTemplateExcelVersion(int templateId, int updatedBy, string version, byte[] excelZip, byte[] excelBytes, string FileExtension)
        {
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_UpdateTemplateExcelVersion", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@TemplateId", templateId);
                        sqlCmd.Parameters.AddWithValue("@UpdatedBy", updatedBy);
                        sqlCmd.Parameters.AddWithValue("@ExcelVersion", version);
                        sqlCmd.Parameters.AddWithValue("@ExcelZip", excelZip);
                        sqlCmd.Parameters.AddWithValue("@ExcelBytes", excelBytes);
                        sqlCmd.Parameters.AddWithValue("@FileExtension", FileExtension);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
        }
        public static string GetParentTableByFileId(int templateFileId)
        {
            var parentTable = string.Empty;
            var dt = new DataTable();
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlComm = new SqlCommand("usp_GetParentTableByFileId", conn))
                    {
                        sqlComm.Parameters.AddWithValue("@TemplateFileId", templateFileId);
                        sqlComm.CommandType = CommandType.StoredProcedure;

                        using (var da = new SqlDataAdapter())
                        {
                            da.SelectCommand = sqlComm;
                            da.Fill(dt);
                        }
                    }
                }
                if (dt.Rows.Count == 1)
                {
                    parentTable = Convert.ToString(dt.Rows[0]["ExcelTableName"]);
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
            return parentTable;

        }
        public static DataTable GetChildTablesByFileId(int templateFileId)
        {
            var dt = new DataTable();
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlComm = new SqlCommand("usp_GetChildTablesByFileId", conn))
                    {
                        sqlComm.Parameters.AddWithValue("@TemplateFileId", templateFileId);
                        sqlComm.CommandType = CommandType.StoredProcedure;

                        using (var da = new SqlDataAdapter())
                        {
                            da.SelectCommand = sqlComm;
                            da.Fill(dt);
                        }
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
            return dt;
        }
        public static DataTable GetMappedFieldParameterByFieldId(string fieldId)
        {
            var dt = new DataTable();
            CommonSql CommonSql = new CommonSql();
            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_GetMappedFieldParameterByFieldId", conn))
                    {
                        sqlCmd.Parameters.AddWithValue("@FieldId", fieldId);
                        sqlCmd.CommandType = CommandType.StoredProcedure;

                        using (var da = new SqlDataAdapter())
                        {
                            da.SelectCommand = sqlCmd;
                            da.Fill(dt);
                        }
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
            return dt;
        }

        public static DataTable GetSendDataParam(int templateId)
        {
            var dt = new DataTable();
            CommonSql CommonSql = new CommonSql();
            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_GetSendDataParam", conn))
                    {
                        sqlCmd.Parameters.AddWithValue("@TemplateId", templateId);
                        sqlCmd.CommandType = CommandType.StoredProcedure;

                        using (var da = new SqlDataAdapter())
                        {
                            da.SelectCommand = sqlCmd;
                            da.Fill(dt);
                        }
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
            return dt;
        }
        public static DataTable GetTableFields(int fileId, string tableName)
        {
            var dt = new DataTable();
            CommonSql CommonSql = new CommonSql();
            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_GetTableFields", conn))
                    {
                        sqlCmd.Parameters.AddWithValue("@FileId", fileId);
                        sqlCmd.Parameters.AddWithValue("@TableName", tableName);
                        sqlCmd.CommandType = CommandType.StoredProcedure;

                        using (var da = new SqlDataAdapter())
                        {
                            da.SelectCommand = sqlCmd;
                            da.Fill(dt);
                        }
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
            return dt;
        }
        public static void AddParentChildTableMapping(int templateFileId, string parentTableName, string parentTableColumn, string childTableName, string childTableColumn)
        {
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_AddParentChildTableMapping", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@TemplateFileId", templateFileId);
                        sqlCmd.Parameters.AddWithValue("@ParentTableName", parentTableName);
                        sqlCmd.Parameters.AddWithValue("@ParentTableColumn", parentTableColumn);
                        sqlCmd.Parameters.AddWithValue("@ChildTableName", childTableName);
                        sqlCmd.Parameters.AddWithValue("@ChildTableColumn", childTableColumn);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
        }
        public static void UpdateFileNamingOption(NamingOption fileNamingOption)
        {
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_UpdateFileNamingOption", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@TemplateId", fileNamingOption.TemplateId);
                        sqlCmd.Parameters.AddWithValue("@Fields", fileNamingOption.Fields);
                        sqlCmd.Parameters.AddWithValue("@UpdatedBy", fileNamingOption.UpdatedBy);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
        }
        public static void UpdateFolderNamingOption(NamingOption folderNamingOption)
        {
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_UpdateFolderNamingOption", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@TemplateId", folderNamingOption.TemplateId);
                        sqlCmd.Parameters.AddWithValue("@Fields", folderNamingOption.Fields);
                        sqlCmd.Parameters.AddWithValue("@UpdatedBy", folderNamingOption.UpdatedBy);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
        }
        public static DataTable GetExistingParentChildMapping(int templateFileId, string parentTable, string childTable)
        {
            var dt = new DataTable();
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_GetExistingParentChildMapping", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@TemplateFileId", templateFileId);
                        sqlCmd.Parameters.AddWithValue("@ParentTable", parentTable);
                        sqlCmd.Parameters.AddWithValue("@ChildTable", childTable);

                        using (var da = new SqlDataAdapter())
                        {
                            da.SelectCommand = sqlCmd;
                            da.Fill(dt);
                        }
                    }
                }
                return dt;
            }
            finally
            {
                CommonSql.CloseConnection();
            }
        }

        public static DataTable GetParentTableFields(int templateId)
        {
            var dt = new DataTable();
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_GetParentTableFields", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@TemplateId", templateId);

                        using (var da = new SqlDataAdapter())
                        {
                            da.SelectCommand = sqlCmd;
                            da.Fill(dt);
                        }
                    }
                }
                return dt;
            }
            finally
            {
                CommonSql.CloseConnection();
            }
        }
        public static void UpdateExcelColumnMapping(int templateId, string oldValue, string newValue)
        {
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_UpdateExcelColumnMapping", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@TemplateId", templateId);
                        sqlCmd.Parameters.AddWithValue("@OldValue", oldValue);
                        sqlCmd.Parameters.AddWithValue("@NewValue", newValue);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
        }
        public static byte[] GetTemplateExcelFileZip(int templateId)
        {
            var dt = new DataTable();
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_GetTemplateExcelFileZip", conn))
                    {
                        sqlCmd.Parameters.AddWithValue("@TemplateId", templateId);
                        sqlCmd.CommandType = CommandType.StoredProcedure;

                        using (var da = new SqlDataAdapter())
                        {
                            da.SelectCommand = sqlCmd;
                            da.Fill(dt);
                        }
                    }
                }
                if (dt.Rows.Count > 0 && dt.Rows[0]["ExcelZip"] != DBNull.Value)
                {
                    return (byte[])dt.Rows[0]["ExcelZip"];
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
            return null;
        }
        public static ExcelFileInfo GetTemplateExcelBytes(int templateId)
        {
            var dt = new DataTable();
            ExcelFileInfo excelFile = new ExcelFileInfo();
            CommonSql CommonSql = new CommonSql();

            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_GetTemplateExcelBytes", conn))
                    {
                        sqlCmd.Parameters.AddWithValue("@TemplateId", templateId);
                        sqlCmd.CommandType = CommandType.StoredProcedure;

                        using (var da = new SqlDataAdapter())
                        {
                            da.SelectCommand = sqlCmd;
                            da.Fill(dt);
                        }
                    }
                }
                if (dt.Rows.Count > 0)
                {
                    if (dt.Rows[0]["ExcelBytes"] != DBNull.Value)
                        excelFile.ExcelBytes = (byte[])dt.Rows[0]["ExcelBytes"];
                    excelFile.TemplateName = Convert.ToString(dt.Rows[0]["TemplateName"]);
                    excelFile.FileExtension = Convert.ToString(dt.Rows[0]["FileExtension"]);
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
            return excelFile;
        }
        public static DataTable GetDynamicField(int templateFileId, string excelTableName)
        {
            var dt = new DataTable();
            var fieldId = string.Empty;
            CommonSql CommonSql = new CommonSql();
            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_GetDynamicField", conn))
                    {
                        sqlCmd.Parameters.AddWithValue("@TemplateFileId", templateFileId);
                        sqlCmd.Parameters.AddWithValue("@ExcelTableName", excelTableName);
                        sqlCmd.CommandType = CommandType.StoredProcedure;

                        using (var da = new SqlDataAdapter())
                        {
                            da.SelectCommand = sqlCmd;
                            da.Fill(dt);
                        }
                    }
                }
                if (dt.Rows.Count == 1)
                    fieldId = Convert.ToString(dt.Rows[0]["FieldId"]);
            }
            finally
            {
                CommonSql.CloseConnection();
            }
            return dt;
        }
        public static DataTable GetDynamicChildFieldsByFieldId(string fieldId)
        {
            var dt = new DataTable();
            CommonSql CommonSql = new CommonSql();
            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_GetDynamicChildFieldsByFieldId", conn))
                    {
                        sqlCmd.Parameters.AddWithValue("@FieldId", fieldId);
                        sqlCmd.CommandType = CommandType.StoredProcedure;

                        using (var da = new SqlDataAdapter())
                        {
                            da.SelectCommand = sqlCmd;
                            da.Fill(dt);
                        }
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
            return dt;
        }

        public static void RemoveMappedField(int templateFileMappingId, bool isStaticField, string parentFieldName)
        {
            CommonSql CommonSql = new CommonSql();
            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_RemoveMappedField", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@TemplateFileMappingId", templateFileMappingId);
                        sqlCmd.Parameters.AddWithValue("@IsStaticField", isStaticField);
                        sqlCmd.Parameters.AddWithValue("@ParentFieldName", parentFieldName);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
        }
        public static void RemoveDynamicFieldMapping(int templateFileMappingId, bool isDynamicField)
        {
            CommonSql CommonSql = new CommonSql();
            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_RemoveDynamicFieldMapping", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@TemplateFileMappingId", templateFileMappingId);
                        sqlCmd.Parameters.AddWithValue("@IsDynamicField", isDynamicField);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
        }
        public static void AddAcroMappedField(MapFieldParam mapFieldParam)
        {
            CommonSql CommonSql = new CommonSql();
            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_AddAcroMappedFieldNew", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@TemplateFileMappingId", mapFieldParam.TemplateFileMappingId);
                        sqlCmd.Parameters.AddWithValue("@SheetName", mapFieldParam.SheetName);
                        sqlCmd.Parameters.AddWithValue("@ExcelTableName", mapFieldParam.TableName);
                        sqlCmd.Parameters.AddWithValue("@ExcelColumnName", mapFieldParam.ColumnName);
                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
        }
        public static void AddXFAMappedField(MapFieldParam mapFieldParam)
        {
            CommonSql CommonSql = new CommonSql();
            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlCmd = new SqlCommand("usp_AddMappedFieldNew", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@TemplateFileMappingId", mapFieldParam.TemplateFileMappingId);
                        sqlCmd.Parameters.AddWithValue("@SheetName", mapFieldParam.SheetName);
                        sqlCmd.Parameters.AddWithValue("@ExcelTableName", mapFieldParam.TableName);
                        sqlCmd.Parameters.AddWithValue("@ExcelColumnName", mapFieldParam.ColumnName);
                        sqlCmd.Parameters.AddWithValue("@IsDynamicElement", mapFieldParam.IsDynamicElement);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
        }

        public static DataTable GetTemplateNamesPart(int templateId)
        {
            var dt = new DataTable();
            CommonSql CommonSql = new CommonSql();
            try
            {
                using (var conn = CommonSql.GetConnection())
                {
                    using (var sqlComm = new SqlCommand("usp_GetTemplateNamesPart", conn))
                    {
                        sqlComm.Parameters.AddWithValue("@TemplateId", templateId);
                        sqlComm.CommandType = CommandType.StoredProcedure;

                        using (var da = new SqlDataAdapter())
                        {
                            da.SelectCommand = sqlComm;
                            da.Fill(dt);
                        }
                    }
                }
            }
            finally
            {
                CommonSql.CloseConnection();
            }
            return dt;

        }
    }
}
