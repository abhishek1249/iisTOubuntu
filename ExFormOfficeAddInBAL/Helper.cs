using ExFormOfficeAddInDAL;
using ExFormOfficeAddInEntities;
using MySql.Data.MySqlClient;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Text.Json.Serialization;

namespace ExFormOfficeAddInBAL
{
    public class Helper
    {
        //public static User LogIn(string userName, string password, string account)
        //{
        //    var dt = new DataTable("User");
        //    User user = null;
        //    //CommonSql CommonSql = new CommonSql();
        //    MySqlConnector mySqlConnector = new MySqlConnector();
        //    try
        //    {
        //        using (var conn = CommonSql.GetConnection())
        //        {
        //            using (var sqlComm = new SqlCommand("usp_Login", conn))
        //            {
        //                sqlComm.Parameters.AddWithValue("@UserName", userName);
        //                sqlComm.Parameters.AddWithValue("@Password", password);
        //                sqlComm.Parameters.AddWithValue("@AccountName", account);
        //                sqlComm.CommandType = CommandType.StoredProcedure;

        //                using (var da = new SqlDataAdapter())
        //                {
        //                    da.SelectCommand = sqlComm;
        //                    da.Fill(dt);
        //                }
        //            }
        //        }
        //        if (dt.Rows.Count == 1)
        //        {
        //            user = new User();
        //            user.UserId = Convert.ToInt32(dt.Rows[0]["UserId"]);
        //            user.FullName = Convert.ToString(dt.Rows[0]["FullName"]);
        //            user.UserName = Convert.ToString(dt.Rows[0]["UserName"]);
        //            user.UserType = Convert.ToChar(dt.Rows[0]["UserType"]);
        //            user.CompanyId = Convert.ToInt32(dt.Rows[0]["CompanyId"]);
        //            user.CompanyName = Convert.ToString(dt.Rows[0]["CompanyName"]);
        //            user.Email = Convert.ToString(dt.Rows[0]["Email"]);
        //            user.IsActive = Convert.ToBoolean(dt.Rows[0]["IsActive"]);
        //        }
        //    }
        //    finally
        //    {
        //        CommonSql.CloseConnection();
        //    }
        //    return user;

        //}

        public static User LogIn(string userName, string password, string account)
        {            
            var ds = new DataSet();
            User user = null;            
            MySqlConnector mySqlConnector = new MySqlConnector();
            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var mySqlComm = new MySqlCommand("usp_Login", conn))
                    {
                        mySqlComm.Parameters.AddWithValue("p_UserName", userName);
                        mySqlComm.Parameters.AddWithValue("p_Password", password);
                        mySqlComm.Parameters.AddWithValue("p_AccountName", account);
                        mySqlComm.CommandType = CommandType.StoredProcedure;

                        using (var da = new MySqlDataAdapter())
                        {
                            da.SelectCommand = mySqlComm;
                            da.Fill(ds);
                        }
                    }
                }
                if (ds.Tables[1].Rows.Count == 1)
                {
                    user = new User();
                    user.UserId = Convert.ToInt32(ds.Tables[1].Rows[0]["UserId"]);
                    user.FullName = Convert.ToString(ds.Tables[1].Rows[0]["FullName"]);
                    user.UserName = Convert.ToString(ds.Tables[1].Rows[0]["UserName"]);
                    user.UserType = Convert.ToChar(ds.Tables[1].Rows[0]["UserType"]);
                    user.CompanyId = Convert.ToInt32(ds.Tables[1].Rows[0]["CompanyId"]);
                    user.CompanyName = Convert.ToString(ds.Tables[1].Rows[0]["CompanyName"]);
                    user.Email = Convert.ToString(ds.Tables[1].Rows[0]["Email"]);
                    user.IsActive = Convert.ToBoolean(ds.Tables[1].Rows[0]["IsActive"]);
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
            return user;

        }
        public static byte[] GetLicenseStream()
        {
            var dt = new DataTable();
            MySqlConnector mySqlConnector = new MySqlConnector();
            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var mySqlComm = new MySqlCommand("usp_GetLicenseStream", conn))
                    {
                        mySqlComm.CommandType = CommandType.StoredProcedure;

                        using (var da = new MySqlDataAdapter())
                        {
                            da.SelectCommand = mySqlComm;
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
                mySqlConnector.CloseConnection();
            }
            return null;

        }
        public static List<TemplateFolder> GetTemplateFolderByCompanyId(int companyId)
        {
            var dt = new DataTable();
            var lstTemplateFolder = new List<TemplateFolder>();
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var mysqlComm = new MySqlCommand("usp_GetTemplateFolderByCompanyId", conn))
                    {
                        mysqlComm.Parameters.AddWithValue("p_CompanyId", companyId);
                        mysqlComm.CommandType = CommandType.StoredProcedure;

                        using (var da = new MySqlDataAdapter())
                        {
                            da.SelectCommand = mysqlComm;
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
                mySqlConnector.CloseConnection();
            }

            return lstTemplateFolder;
        }
        public static TemplateFolder GetDemoFolder()
        {
            var dt = new DataTable();
            var templateFoder = new TemplateFolder();
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlComm = new MySqlCommand("usp_GetDemoFolder", conn))
                    {
                        sqlComm.CommandType = CommandType.StoredProcedure;

                        using (var da = new MySqlDataAdapter())
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
                mySqlConnector.CloseConnection();
            }

            return templateFoder;
        }
        public static List<Template> GetTemplateByFolderId(int? folderId)
        {
            MySqlConnector mySqlConnector = new MySqlConnector();
            var dt = new DataTable("Template");
            var lstTemplate = new List<Template>();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlComm = new MySqlCommand("usp_GetTemplateByFolderId", conn))
                    {
                        sqlComm.Parameters.AddWithValue("p_FolderId", folderId);
                        sqlComm.CommandType = CommandType.StoredProcedure;

                        using (var da = new MySqlDataAdapter())
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
                mySqlConnector.CloseConnection();
            }

            return lstTemplate;
        }

        public static int CreateTemplate(PdfTemplate pdfTemplate)
        {
            MySqlConnector mySqlConnector = new MySqlConnector();
            var templateId = -1;
            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_CreateTemplateSet", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_CompanyId", pdfTemplate.CompanyId);
                        sqlCmd.Parameters.AddWithValue("p_TemplateName", pdfTemplate.TemplateName);
                        sqlCmd.Parameters.AddWithValue("p_Description", pdfTemplate.Description);
                        sqlCmd.Parameters.AddWithValue("p_TemplateFileZip", pdfTemplate.TemplateFileZip);
                        sqlCmd.Parameters.AddWithValue("p_IsActive", pdfTemplate.IsActive);
                        sqlCmd.Parameters.AddWithValue("p_CreatedOn", pdfTemplate.CreatedOn);
                        sqlCmd.Parameters.AddWithValue("p_CreatedBy", pdfTemplate.CreatedBy);
                        sqlCmd.Parameters.AddWithValue("p_ExcelZip", pdfTemplate.ExcelZip);
                       
                        sqlCmd.Parameters.AddWithValue("p_TemplateFolderId", pdfTemplate.TemplateFolderId);
                        sqlCmd.Parameters.AddWithValue("p_FolderName", pdfTemplate.FolderName);
                        sqlCmd.Parameters.AddWithValue("p_SubFolderName", pdfTemplate.SubFolderName);
                        sqlCmd.Parameters.AddWithValue("p_FileNamePart", pdfTemplate.FileNamePart);
                        sqlCmd.Parameters.AddWithValue("p_ExcelVersion", pdfTemplate.ExcelVersion);
                        sqlCmd.Parameters.AddWithValue("p_IsDemoTemplate", pdfTemplate.IsDemoTemplate);
                        sqlCmd.Parameters.AddWithValue("p_ExcelBytes", pdfTemplate.ExcelBytes);
                        sqlCmd.Parameters.AddWithValue("p_FileExtension", pdfTemplate.FileExtension);

                       
                        sqlCmd.Parameters.Add("p_TemplateFile", MySqlDbType.JSON);
                        sqlCmd.Parameters["p_TemplateFile"].Value = JsonConvert.SerializeObject(pdfTemplate.TemplateFile);

                        sqlCmd.Parameters.Add("p_TemplateFileFieldMapping", MySqlDbType.JSON);
                        sqlCmd.Parameters["p_TemplateFileFieldMapping"].Value = JsonConvert.SerializeObject(pdfTemplate.TemplateFileFieldMapping);

                        sqlCmd.Parameters.Add("p_TemplateId", MySqlDbType.Int32);
                        sqlCmd.Parameters["p_TemplateId"].Direction = ParameterDirection.Output;
                        sqlCmd.ExecuteNonQuery();
                        templateId = Convert.ToInt32(sqlCmd.Parameters["p_TemplateId"].Value);
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
            return templateId;
        }

        public static void UploadExcelFile(ExcelFile excelFile)
        {
            MySqlConnector mySqlConnector = new MySqlConnector();
            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_UpdateExcelBytes", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_TemplateId", excelFile.TemplateId);
                        sqlCmd.Parameters.AddWithValue("p_ExcelBytes", excelFile.fileBytes);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
        }
        public static void UpdateTemplate(EditPdfTemplate editPdfTemplate)
        {
            MySqlConnector mySqlConnector = new MySqlConnector();
            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_UpdateTemplateSet", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_DeletedTemplateFileIds", editPdfTemplate.DeletedTemplateFileIds);
                        if (editPdfTemplate.UpdatedBy.HasValue)
                            sqlCmd.Parameters.AddWithValue("p_UpdatedBy", editPdfTemplate.UpdatedBy);
                        else
                            sqlCmd.Parameters.AddWithValue("p_UpdatedBy", DBNull.Value);

                        sqlCmd.Parameters.Add("p_Template",MySqlDbType.JSON);
                        sqlCmd.Parameters["p_Template"].Value = JsonConvert.SerializeObject(editPdfTemplate.Template);

                        sqlCmd.Parameters.Add("p_TemplateFile", MySqlDbType.JSON);
                        sqlCmd.Parameters["p_TemplateFile"].Value = JsonConvert.SerializeObject(editPdfTemplate.TemplateFile);

                        sqlCmd.Parameters.Add("p_TemplateFileFieldMapping", MySqlDbType.JSON);
                        sqlCmd.Parameters["p_TemplateFileFieldMapping"].Value = JsonConvert.SerializeObject(editPdfTemplate.TemplateFileFieldMapping);

                        
                        sqlCmd.Parameters.AddWithValue("p_TemplateFileZip", editPdfTemplate.TemplateFileZip);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
        }
        public static DataTable GetTemplateByTemplateId(int templateId)
        {
            MySqlConnector mySqlConnector = new MySqlConnector();
            var dt = new DataTable();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_GetTemplateById", conn))
                    {
                        sqlCmd.Parameters.AddWithValue("p_TemplateId", templateId);
                        sqlCmd.CommandType = CommandType.StoredProcedure;

                        using (var da = new MySqlDataAdapter())
                        {
                            da.SelectCommand = sqlCmd;
                            da.Fill(dt);
                        }
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
            return dt;
        }
        public static void DeleteFolder(int folderId)
        {
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_DeleteTemplateFolder", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_FolderId", folderId);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
        }
        public static void CreateTemplateFolder(int companyId, int? parentFolderId, string folderName)
        {
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_CreateTemplateFolder", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_CompanyId", companyId);

                        if (parentFolderId.HasValue)
                            sqlCmd.Parameters.AddWithValue("p_ParentFolderId", parentFolderId);
                        else
                            sqlCmd.Parameters.AddWithValue("p_ParentFolderId", DBNull.Value);

                        sqlCmd.Parameters.AddWithValue("p_FolderName", folderName);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
        }

        public static void RenameFolder(int folderId, string folderName)
        {
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_UpdateTemplateFolderName", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_FolderId", folderId);
                        sqlCmd.Parameters.AddWithValue("p_Name", folderName);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
        }
        public static void CreateTemplateCopy(int templateId, int companyId, int createdBy)
        {
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_CreateTemplateCopy", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_TemplateId", templateId);
                        sqlCmd.Parameters.AddWithValue("p_CompanyId", companyId);
                        sqlCmd.Parameters.AddWithValue("p_CreatedBy", createdBy);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
        }
        public static void RenameTemplate(int templateId, string templateName)
        {
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_UpdateTemplateName", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_TemplateId", templateId);
                        sqlCmd.Parameters.AddWithValue("p_Name", templateName);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
        }

        public static void DeleteTemplate(int templateId)
        {
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_DeleteTemplate", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_TemplateId", templateId);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
        }

        public static void DeleteTemplateFile(int templateFileId)
        {
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_DeleteTemplateFile", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_TemplateFileId", templateFileId);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
        }

        public static TemplateInfo GetTemplateById(int templateId)
        {
            var dt = new DataTable();
            var dt1 = new DataTable();
            var templateInfo = new TemplateInfo();
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_GetTemplateById", conn))
                    {
                        sqlCmd.Parameters.AddWithValue("p_TemplateId", templateId);
                        sqlCmd.CommandType = CommandType.StoredProcedure;

                        using (var da = new MySqlDataAdapter())
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

                MySqlConnector mySqlConnector1 = new MySqlConnector();
                using (var conn = mySqlConnector1.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_GetTemplateFiles", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_TemplateId", templateId);

                        using (var da = new MySqlDataAdapter())
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
                mySqlConnector.CloseConnection();
            }
            return templateInfo;
        }

        public static DataTable GetTemplateMappedFields(int templateId)
        {
            var dt = new DataTable();
            MySqlConnector mySqlConnector = new MySqlConnector();
            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_GetTemplateMappedFields", conn))
                    {
                        sqlCmd.Parameters.AddWithValue("p_TemplateId", templateId);
                        sqlCmd.CommandType = CommandType.StoredProcedure;

                        using (var da = new MySqlDataAdapter())
                        {
                            da.SelectCommand = sqlCmd;
                            da.Fill(dt);
                        }
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
            return dt;
        }

        public static List<TemplateFile> GetTemplateFiles(int templateId)
        {
            var dt = new DataTable();
            var templateFiles = new List<TemplateFile>();
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {

                    using (var sqlCmd = new MySqlCommand("usp_GetTemplateFiles", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_TemplateId", templateId);

                        using (var da = new MySqlDataAdapter())
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
                mySqlConnector.CloseConnection();
            }
            return templateFiles;
        }

        public static IsExcelVersionExistResponse IsExcelVersionExist(string excelVersion)
        {
            var dt = new DataTable();
            var response = new IsExcelVersionExistResponse();
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {

                    using (var sqlCmd = new MySqlCommand("usp_IsExcelVersionExist", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_ExcelVersion", excelVersion);

                        using (var da = new MySqlDataAdapter())
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
                mySqlConnector.CloseConnection();
            }
            return response;
        }

        public static DataTable GetParentChildTableMapping(int templateFileId)
        {
            var dt = new DataTable();
            MySqlConnector mySqlConnector = new MySqlConnector();
            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_GetParentChildTableMapping", conn))
                    {
                        sqlCmd.Parameters.AddWithValue("p_TemplateFileId", templateFileId);
                        sqlCmd.CommandType = CommandType.StoredProcedure;

                        using (var da = new MySqlDataAdapter())
                        {
                            da.SelectCommand = sqlCmd;
                            da.Fill(dt);
                        }
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
            return dt;
        }
        public static string GetTemplateFileFieldsMappedPercentage(int templateFileId)
        {
            MySqlConnector mySqlConnector = new MySqlConnector();
            var mappedPercentage = "0.00%";

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_TemplateFileFieldsMappedPercentage", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_TemplateFileId", templateFileId);

                        sqlCmd.Parameters.Add("@Percentage", MySqlDbType.Float);
                        sqlCmd.Parameters["p_Percentage"].Direction = ParameterDirection.Output;
                        sqlCmd.ExecuteNonQuery();
                        mappedPercentage = Convert.ToString(sqlCmd.Parameters["p_Percentage"].Value) + "%";
                        if (mappedPercentage.Length > 4)
                            mappedPercentage = mappedPercentage.Substring(0, 4) + "%";
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
            return mappedPercentage;
        }

        public static List<TemplateFileFieldsMapping> GetTemplateFieldsByTemplateFileId(int templateFileId, bool isAdmin)
        {
            var dt = new DataTable();
            List<TemplateFileFieldsMapping> lstTemplateFileFieldsMapping = new List<TemplateFileFieldsMapping>();
            MySqlConnector mySqlConnector = new MySqlConnector();
            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_GetTemplateFieldsByTemplateFileId", conn))
                    {
                        sqlCmd.Parameters.AddWithValue("p_TemplateFileId", templateFileId);
                        sqlCmd.Parameters.AddWithValue("p_IsAdmin", isAdmin);
                        sqlCmd.CommandType = CommandType.StoredProcedure;

                        using (var da = new MySqlDataAdapter())
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
                mySqlConnector.CloseConnection();
            }
            return lstTemplateFileFieldsMapping;
        }

        public static bool IsTemplateFieldsMapped(int templateId)
        {
            bool isMapped = false;
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_IsTemplateFieldsMapped", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_TemplateId", templateId);

                        sqlCmd.Parameters.Add("p_IsMapped", MySqlDbType.Bit);
                        sqlCmd.Parameters["p_IsMapped"].Direction = ParameterDirection.Output;
                        sqlCmd.ExecuteNonQuery();
                        isMapped = Convert.ToBoolean(sqlCmd.Parameters["p_IsMapped"].Value);
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
            return isMapped;
        }

        public static void RemoveAllMappedFields(int templateId)
        {
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_RemoveAllMappedFields", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_TemplateId", templateId);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
        }
        public static int GetDynamicFieldsCount(int templateFileId)
        {
            var count = 0;
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_GetDynamicFieldsCount", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_TemplateFileId", templateFileId);

                        sqlCmd.Parameters.Add("p_Count", MySqlDbType.Int32);
                        sqlCmd.Parameters["p_Count"].Direction = ParameterDirection.Output;
                        sqlCmd.ExecuteNonQuery();
                        count = Convert.ToInt32(sqlCmd.Parameters["p_Count"].Value);
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
            return count;
        }

        public static int GetParentChildTableRelationshipCount(int templateFileId)
        {
            var count = 0;
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_GetParentChildTableRelationshipCount", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_TemplateFileId", templateFileId);

                        sqlCmd.Parameters.Add("p_Count", MySqlDbType.Int32);
                        sqlCmd.Parameters["p_Count"].Direction = ParameterDirection.Output;
                        sqlCmd.ExecuteNonQuery();
                        count = Convert.ToInt32(sqlCmd.Parameters["p_Count"].Value);
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
            return count;
        }
        public static bool IsTemplateFileFieldsMapped(int templateFileId)
        {
            bool isMapped = false;
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_IsTemplateFileFieldsMapped", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_TemplateFileId", templateFileId);

                        sqlCmd.Parameters.Add("p_IsMapped", MySqlDbType.Bit);
                        sqlCmd.Parameters["p_IsMapped"].Direction = ParameterDirection.Output;
                        sqlCmd.ExecuteNonQuery();
                        isMapped = Convert.ToBoolean(sqlCmd.Parameters["p_IsMapped"].Value);
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
            return isMapped;
        }
        public static void RemoveFileMappedFields(int templateFileId)
        {
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_RemoveFileMappedFields", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_TemplateFileId", templateFileId);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
        }
        public static DataTable GetTemplateFieldsByTemplateId(int templateId)
        {
            var dt = new DataTable();
            var fieldId = string.Empty;
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_GetTemplateFieldsByTemplateId", conn))
                    {
                        sqlCmd.Parameters.AddWithValue("p_TemplateId", templateId);
                        sqlCmd.CommandType = CommandType.StoredProcedure;

                        using (var da = new MySqlDataAdapter())
                        {
                            da.SelectCommand = sqlCmd;
                            da.Fill(dt);
                        }
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
            return dt;
        }
        public static void MapParentTableFields(DataTable parentField)
        {
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_MapParentTableFields", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.Add("p_ParentField", MySqlDbType.JSON);
                        sqlCmd.Parameters["p_ParentField"].Value = JsonConvert.SerializeObject(parentField);
                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
        }
        public static DataTable GetExcelVersion(int templateId)
        {
            var dt = new DataTable();
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_GetExcelVersion", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_TemplateId", templateId);

                        using (var da = new MySqlDataAdapter())
                        {
                            da.SelectCommand = sqlCmd;
                            da.Fill(dt);
                        }
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
            return dt;
        }
        public static void UpdateTemplateExcelVersion(int templateId, int updatedBy, string version, byte[] excelZip, byte[] excelBytes, string FileExtension)
        {
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_UpdateTemplateExcelVersion", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_TemplateId", templateId);
                        sqlCmd.Parameters.AddWithValue("p_UpdatedBy", updatedBy);
                        sqlCmd.Parameters.AddWithValue("p_ExcelVersion", version);
                        sqlCmd.Parameters.AddWithValue("p_ExcelZip", excelZip);
                        if(excelBytes != null)
                        {
                            sqlCmd.Parameters.AddWithValue("p_ExcelBytes", excelBytes);
                        }
                        else
                        {
                            sqlCmd.Parameters.AddWithValue("p_ExcelBytes", DBNull.Value);
                        }
                        
                        sqlCmd.Parameters.AddWithValue("p_FileExtension", FileExtension);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
        }
        public static string GetParentTableByFileId(int templateFileId)
        {
            var parentTable = string.Empty;
            var dt = new DataTable();
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlComm = new MySqlCommand("usp_GetParentTableByFileId", conn))
                    {
                        sqlComm.Parameters.AddWithValue("p_TemplateFileId", templateFileId);
                        sqlComm.CommandType = CommandType.StoredProcedure;

                        using (var da = new MySqlDataAdapter())
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
                mySqlConnector.CloseConnection();
            }
            return parentTable;

        }
        public static DataTable GetChildTablesByFileId(int templateFileId)
        {
            var dt = new DataTable();
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlComm = new MySqlCommand("usp_GetChildTablesByFileId", conn))
                    {
                        sqlComm.Parameters.AddWithValue("p_TemplateFileId", templateFileId);
                        sqlComm.CommandType = CommandType.StoredProcedure;

                        using (var da = new MySqlDataAdapter())
                        {
                            da.SelectCommand = sqlComm;
                            da.Fill(dt);
                        }
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
            return dt;
        }
        public static DataTable GetMappedFieldParameterByFieldId(string fieldId)
        {
            var dt = new DataTable();
            MySqlConnector mySqlConnector = new MySqlConnector();
            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_GetMappedFieldParameterByFieldId", conn))
                    {
                        sqlCmd.Parameters.AddWithValue("p_FieldId", fieldId);
                        sqlCmd.CommandType = CommandType.StoredProcedure;

                        using (var da = new MySqlDataAdapter())
                        {
                            da.SelectCommand = sqlCmd;
                            da.Fill(dt);
                        }
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
            return dt;
        }

        public static DataTable GetSendDataParam(int templateId)
        {
            var dt = new DataTable();
            MySqlConnector mySqlConnector = new MySqlConnector();
            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_GetSendDataParam", conn))
                    {
                        sqlCmd.Parameters.AddWithValue("p_TemplateId", templateId);
                        sqlCmd.CommandType = CommandType.StoredProcedure;

                        using (var da = new MySqlDataAdapter())
                        {
                            da.SelectCommand = sqlCmd;
                            da.Fill(dt);
                        }
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
            return dt;
        }
        public static DataTable GetTableFields(int fileId, string tableName)
        {
            var dt = new DataTable();
            MySqlConnector mySqlConnector = new MySqlConnector();
            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_GetTableFields", conn))
                    {
                        sqlCmd.Parameters.AddWithValue("p_FileId", fileId);
                        sqlCmd.Parameters.AddWithValue("p_TableName", tableName);
                        sqlCmd.CommandType = CommandType.StoredProcedure;

                        using (var da = new MySqlDataAdapter())
                        {
                            da.SelectCommand = sqlCmd;
                            da.Fill(dt);
                        }
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
            return dt;
        }
        public static void AddParentChildTableMapping(int templateFileId, string parentTableName, string parentTableColumn, string childTableName, string childTableColumn)
        {
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_AddParentChildTableMapping", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_TemplateFileId", templateFileId);
                        sqlCmd.Parameters.AddWithValue("p_ParentTableName", parentTableName);
                        sqlCmd.Parameters.AddWithValue("p_ParentTableColumn", parentTableColumn);
                        sqlCmd.Parameters.AddWithValue("p_ChildTableName", childTableName);
                        sqlCmd.Parameters.AddWithValue("p_ChildTableColumn", childTableColumn);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
        }
        public static void UpdateFileNamingOption(NamingOption fileNamingOption)
        {
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_UpdateFileNamingOption", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_TemplateId", fileNamingOption.TemplateId);
                        sqlCmd.Parameters.AddWithValue("p_Fields", fileNamingOption.Fields);
                        sqlCmd.Parameters.AddWithValue("p_UpdatedBy", fileNamingOption.UpdatedBy);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
        }
        public static void UpdateFolderNamingOption(NamingOption folderNamingOption)
        {
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_UpdateFolderNamingOption", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_TemplateId", folderNamingOption.TemplateId);
                        sqlCmd.Parameters.AddWithValue("p_Fields", folderNamingOption.Fields);
                        sqlCmd.Parameters.AddWithValue("p_UpdatedBy", folderNamingOption.UpdatedBy);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
        }
        public static DataTable GetExistingParentChildMapping(int templateFileId, string parentTable, string childTable)
        {
            var dt = new DataTable();
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_GetExistingParentChildMapping", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_TemplateFileId", templateFileId);
                        sqlCmd.Parameters.AddWithValue("p_ParentTable", parentTable);
                        sqlCmd.Parameters.AddWithValue("p_ChildTable", childTable);

                        using (var da = new MySqlDataAdapter())
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
                mySqlConnector.CloseConnection();
            }
        }

        public static DataTable GetParentTableFields(int templateId)
        {
            var dt = new DataTable();
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_GetParentTableFields", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_TemplateId", templateId);

                        using (var da = new MySqlDataAdapter())
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
                mySqlConnector.CloseConnection();
            }
        }
        public static void UpdateExcelColumnMapping(int templateId, string oldValue, string newValue)
        {
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_UpdateExcelColumnMapping", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_TemplateId", templateId);
                        sqlCmd.Parameters.AddWithValue("p_OldValue", oldValue);
                        sqlCmd.Parameters.AddWithValue("p_NewValue", newValue);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
        }
        public static byte[] GetTemplateExcelFileZip(int templateId)
        {
            var dt = new DataTable();
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_GetTemplateExcelFileZip", conn))
                    {
                        sqlCmd.Parameters.AddWithValue("p_TemplateId", templateId);
                        sqlCmd.CommandType = CommandType.StoredProcedure;

                        using (var da = new MySqlDataAdapter())
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
                mySqlConnector.CloseConnection();
            }
            return null;
        }
        public static ExcelFileInfo GetTemplateExcelBytes(int templateId)
        {
            var dt = new DataTable();
            ExcelFileInfo excelFile = new ExcelFileInfo();
            MySqlConnector mySqlConnector = new MySqlConnector();

            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_GetTemplateExcelBytes", conn))
                    {
                        sqlCmd.Parameters.AddWithValue("p_TemplateId", templateId);
                        sqlCmd.CommandType = CommandType.StoredProcedure;

                        using (var da = new MySqlDataAdapter())
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
                mySqlConnector.CloseConnection();
            }
            return excelFile;
        }
        public static DataTable GetDynamicField(int templateFileId, string excelTableName)
        {
            var dt = new DataTable();
            var fieldId = string.Empty;
            MySqlConnector mySqlConnector = new MySqlConnector();
            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_GetDynamicField", conn))
                    {
                        sqlCmd.Parameters.AddWithValue("p_TemplateFileId", templateFileId);
                        sqlCmd.Parameters.AddWithValue("p_ExcelTableName", excelTableName);
                        sqlCmd.CommandType = CommandType.StoredProcedure;

                        using (var da = new MySqlDataAdapter())
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
                mySqlConnector.CloseConnection();
            }
            return dt;
        }
        public static DataTable GetDynamicChildFieldsByFieldId(string fieldId)
        {
            var dt = new DataTable();
            MySqlConnector mySqlConnector = new MySqlConnector();
            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_GetDynamicChildFieldsByFieldId", conn))
                    {
                        sqlCmd.Parameters.AddWithValue("p_FieldId", fieldId);
                        sqlCmd.CommandType = CommandType.StoredProcedure;

                        using (var da = new MySqlDataAdapter())
                        {
                            da.SelectCommand = sqlCmd;
                            da.Fill(dt);
                        }
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
            return dt;
        }

        public static void RemoveMappedField(int templateFileMappingId, bool isStaticField, string parentFieldName)
        {
            MySqlConnector mySqlConnector = new MySqlConnector();
            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_RemoveMappedField", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_TemplateFileMappingId", templateFileMappingId);
                        sqlCmd.Parameters.AddWithValue("p_IsStaticField", isStaticField);
                        sqlCmd.Parameters.AddWithValue("p_ParentFieldName", parentFieldName);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
        }
        public static void RemoveDynamicFieldMapping(int templateFileMappingId, bool isDynamicField)
        {
            MySqlConnector mySqlConnector = new MySqlConnector();
            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_RemoveDynamicFieldMapping", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_TemplateFileMappingId", templateFileMappingId);
                        sqlCmd.Parameters.AddWithValue("p_IsDynamicField", isDynamicField);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
        }
        public static void AddAcroMappedField(MapFieldParam mapFieldParam)
        {
            MySqlConnector mySqlConnector = new MySqlConnector();
            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_AddAcroMappedFieldNew", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_TemplateFileMappingId", mapFieldParam.TemplateFileMappingId);
                        sqlCmd.Parameters.AddWithValue("p_SheetName", mapFieldParam.SheetName);
                        sqlCmd.Parameters.AddWithValue("p_ExcelTableName", mapFieldParam.TableName);
                        sqlCmd.Parameters.AddWithValue("p_ExcelColumnName", mapFieldParam.ColumnName);
                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
        }
        public static void AddXFAMappedField(MapFieldParam mapFieldParam)
        {
            MySqlConnector mySqlConnector = new MySqlConnector();
            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlCmd = new MySqlCommand("usp_AddMappedFieldNew", conn))
                    {
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("p_TemplateFileMappingId", mapFieldParam.TemplateFileMappingId);
                        sqlCmd.Parameters.AddWithValue("p_SheetName", mapFieldParam.SheetName);
                        sqlCmd.Parameters.AddWithValue("p_ExcelTableName", mapFieldParam.TableName);
                        sqlCmd.Parameters.AddWithValue("p_ExcelColumnName", mapFieldParam.ColumnName);
                        sqlCmd.Parameters.AddWithValue("p_IsDynamicElement", mapFieldParam.IsDynamicElement);

                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
        }

        public static DataTable GetTemplateNamesPart(int templateId)
        {
            var dt = new DataTable();
            MySqlConnector mySqlConnector = new MySqlConnector();
            try
            {
                using (var conn = mySqlConnector.GetConnection())
                {
                    using (var sqlComm = new MySqlCommand("usp_GetTemplateNamesPart", conn))
                    {
                        sqlComm.Parameters.AddWithValue("p_TemplateId", templateId);
                        sqlComm.CommandType = CommandType.StoredProcedure;

                        using (var da = new MySqlDataAdapter())
                        {
                            da.SelectCommand = sqlComm;
                            da.Fill(dt);
                        }
                    }
                }
            }
            finally
            {
                mySqlConnector.CloseConnection();
            }
            return dt;

        }
    }
}
