using Aspose.Pdf;
using ExFormOfficeAddInBAL;
using ExFormOfficeAddInEntities;
using Ionic.Zip;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;
using System.Web.UI.WebControls;
using System.Xml;
using System.Xml.Linq;
using AsposePdf = Aspose.Pdf;

namespace ExFormOfficeAddInExcelUIWeb.Controllers
{
    public class TemplateController : ApiController
    {
        string _previousParentId = null;
        string _previousParentName = null;
        DataSet _dsNewSetPdf;
        DataSet _dsPdf;
        bool isDynamic = false;

        public TemplateController()
        {
            var licenseStream = Helper.GetLicenseStream();

            if (licenseStream != null)
                ExFormsAsposePdf.LoadLicenseFromStream(licenseStream);
        }

        [Route("api/Template/SaveAutomapFields")]
        [HttpPost()]
        public string SaveAutomapFields(MappingFields fields)
        {
            var dtParentFields = new DataTable();
            dtParentFields.Clear();
            CreateParentFieldsDt(dtParentFields);
            try
            {
                foreach (var field in fields.ParentFields)
                {
                    var dataRow = dtParentFields.NewRow();
                    dataRow["TemplateFileMappingId"] = field.TemplateFileMappingId;
                    dataRow["SheetName"] = field.SheetName;
                    dataRow["ExcelTableName"] = field.ExcelTableName;
                    dataRow["ExcelFieldName"] = field.ExcelFieldName;
                    dataRow["IsMapped"] = field.IsMapped;
                    dtParentFields.Rows.Add(dataRow);
                }
                foreach (var field in fields.ChildFields)
                {
                    var dataRow = dtParentFields.NewRow();
                    dataRow["TemplateFileMappingId"] = field.TemplateFileMappingId;
                    dataRow["SheetName"] = field.SheetName;
                    dataRow["ExcelTableName"] = field.ExcelTableName;
                    dataRow["ExcelFieldName"] = field.ExcelFieldName;
                    dataRow["IsMapped"] = field.IsMapped;
                    dtParentFields.Rows.Add(dataRow);
                }

                var templateFields = GetTemplateFields(fields.TemplateId);

                foreach (var field in fields.DynamicFieldIds)
                {
                    foreach (var childField in fields.ChildFields)
                    {
                        if (childField.ParentFieldId == field)
                        {
                            var dynamicFieldItem = templateFields.AsEnumerable().Where(row => row.Field<string>("FieldId") == field);
                            if (dynamicFieldItem.Any() && dynamicFieldItem.Count() == 1)
                            {
                                var dynamicFieldItemdt = dynamicFieldItem.CopyToDataTable();
                                var dataRow = dtParentFields.NewRow();
                                dataRow["TemplateFileMappingId"] = Convert.ToInt32(dynamicFieldItemdt.Rows[0]["TemplateFileMappingId"]); ;
                                dataRow["SheetName"] = childField.SheetName;
                                dataRow["ExcelTableName"] = childField.ExcelTableName;
                                dataRow["ExcelFieldName"] = null;
                                dataRow["IsMapped"] = true;
                                dtParentFields.Rows.Add(dataRow);
                            }
                            break;
                        }
                    }
                }

                Helper.MapParentTableFields(dtParentFields);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            return "success";
        }

        [Route("api/Template/CreateTemplate")]
        [HttpPost()]
        public string CreateTemplate(JTemplate pdfTemplate)
        {
            var dtMappedPdfFields = new DataTable();
            try
            {
                CreateMappedPdfTable(dtMappedPdfFields);
                var drMapped = dtMappedPdfFields.NewRow();
                drMapped["PDFFieldName"] = "test";
                drMapped["FilePath"] = "test";
                drMapped["IsMapped"] = false;
                drMapped["SheetName"] = "test";
                drMapped["ExcelTableName"] = "test";
                drMapped["FieldId"] = "test";
                drMapped["ParentFieldId"] = "test";
                drMapped["IsDynamic"] = false;
                drMapped["HasChildFields"] = false;
                drMapped["XPath"] = "test";
                dtMappedPdfFields.Rows.Add(drMapped);

                var _dtSelectedPdfFiles = new DataTable();
                CreateTemplateFileTable(ref _dtSelectedPdfFiles);
                foreach (var file in pdfTemplate.TemplateFile)
                {
                    var files = file.Split(',');
                    var dataRow = _dtSelectedPdfFiles.NewRow();
                    dataRow["FileName"] = files[0];
                    dataRow["FilePath"] = files[1];
                    dataRow["IsXFA"] = false;
                    _dtSelectedPdfFiles.Rows.Add(dataRow);
                }

                var pdfTemplateSet = new PdfTemplate()
                {
                    TemplateName = pdfTemplate.TemplateName,
                    Description = pdfTemplate.Description,
                    CompanyId = pdfTemplate.CompanyId,
                    TemplateFileZip = new MemoryStream().ToArray(),
                    IsActive = pdfTemplate.IsActive,
                    CreatedOn = DateTime.Now,
                    CreatedBy = pdfTemplate.CreatedBy,
                    ExcelZip = new MemoryStream().ToArray(),
                    TemplateFile = _dtSelectedPdfFiles,
                    TemplateFileFieldMapping = dtMappedPdfFields,
                    TemplateFolderId = pdfTemplate.TemplateFolderId,
                    FolderName = pdfTemplate.FolderName,
                    SubFolderName = pdfTemplate.SubFolderName,
                    FileNamePart = pdfTemplate.FileNamePart,
                    ExcelVersion = pdfTemplate.ExcelVersion
                };
                Helper.CreateTemplate(pdfTemplateSet);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            return "success";
        }

        [Route("api/Template/CreateTemplateSet")]
        [HttpPost()]
        public CreateTemplateResponse CreateTemplateSet()
        {
            var dtMappedPdfFields = new DataTable();
            CreateTemplateResponse response = new CreateTemplateResponse();
            response.TemplateId = -1;

            try
            {
                CreateMappedPdfTable(dtMappedPdfFields);

                var dtSelectedPdfFiles = new DataTable();
                CreateTemplateFileTable(ref dtSelectedPdfFiles);
                ZipFile templateZip = new ZipFile();
                ZipFile excelZip = new ZipFile();
                var templateMemoryStream = new MemoryStream();
                var excelMemoryStream = new MemoryStream();
                var pdfTemplateSet = new PdfTemplate();
                pdfTemplateSet.ExcelBytes = new MemoryStream().ToArray();

                var httpContext = HttpContext.Current;

                if (httpContext.Request.Files.Count > 0)
                {
                    for (int j = 0; j < httpContext.Request.Files.Count; j++)
                    {
                        Pdf PdfFile = new Pdf();
                        HttpPostedFile httpPostedFile = httpContext.Request.Files[j];

                        if (httpPostedFile != null && httpPostedFile.ContentType == "application/pdf")
                        {
                            AsposePdf.Document pdfDocument = new AsposePdf.Document(httpPostedFile.InputStream);

                            if (pdfDocument.Form.Type == AsposePdf.Forms.FormType.Dynamic)
                                PdfFile.IsXFA = true;

                            PdfFile.Folder = httpPostedFile.FileName;
                            PdfFile.FileName = httpPostedFile.FileName.Substring(httpPostedFile.FileName.LastIndexOf("\\") + 1);
                            var userId = Convert.ToInt32(httpContext.Request["CreatedBy"]);
                            GenerateNewSetPdfFields(PdfFile, pdfDocument, httpPostedFile.InputStream, userId);

                            var dataRow = dtSelectedPdfFiles.NewRow();
                            dataRow["FileName"] = PdfFile.FileName;
                            dataRow["FilePath"] = PdfFile.Folder;
                            dataRow["IsXFA"] = PdfFile.IsXFA;
                            using (var binaryReader = new BinaryReader(httpPostedFile.InputStream))
                            {
                                dataRow["PdfBytes"] = binaryReader.ReadBytes(httpPostedFile.ContentLength);
                            }
                            dtSelectedPdfFiles.Rows.Add(dataRow);
                            DataTable dt = new DataTable();
                            dt = _dsNewSetPdf.Tables[j];
                            for (int i = 0; i < dt.Rows.Count; i++)
                            {
                                var drMapped = dtMappedPdfFields.NewRow();
                                drMapped["PDFFieldName"] = dt.Rows[i]["PDFField Name"];
                                drMapped["FilePath"] = dt.Rows[i]["FilePath"];
                                drMapped["IsMapped"] = dt.Rows[i]["Is Mapped"];
                                drMapped["SheetName"] = dt.Rows[i]["Sheet Name"];
                                drMapped["ExcelTableName"] = dt.Rows[i]["Excel Table Name"];
                                drMapped["FieldId"] = dt.Rows[i]["FieldId"];
                                drMapped["ParentFieldId"] = dt.Rows[i]["ParentFieldId"];
                                drMapped["IsDynamic"] = dt.Rows[i]["IsDynamic"];
                                drMapped["HasChildFields"] = dt.Rows[i]["HasChildFields"];
                                drMapped["XPath"] = dt.Rows[i]["XPath"];
                                dtMappedPdfFields.Rows.Add(drMapped);
                            }
                        }
                    }

                    foreach (var key in httpContext.Request.Form.AllKeys)
                    {
                        switch (key)
                        {
                            case "TemplateFile":
                                pdfTemplateSet.TemplateFile = dtSelectedPdfFiles;
                                break;
                            case "TemplateName":
                                pdfTemplateSet.TemplateName = httpContext.Request[key];
                                break;
                            case "Description":
                                pdfTemplateSet.Description = httpContext.Request[key];
                                break;
                            case "CompanyId":
                                pdfTemplateSet.CompanyId = Convert.ToInt32(httpContext.Request[key]);
                                break;
                            case "TemplateFileZip":
                                pdfTemplateSet.TemplateFileZip = new MemoryStream().ToArray();
                                break;
                            case "IsActive":
                                pdfTemplateSet.IsActive = false;
                                break;
                            case "CreatedOn":
                                pdfTemplateSet.CreatedOn = DateTime.Now;
                                break;
                            case "ExcelZip":
                                pdfTemplateSet.ExcelZip = new MemoryStream().ToArray();
                                break;
                            case "TemplateFileFieldMapping":
                                pdfTemplateSet.TemplateFileFieldMapping = dtMappedPdfFields;
                                break;

                            case "TemplateFolderId":
                                pdfTemplateSet.TemplateFolderId = Convert.ToInt32(httpContext.Request[key]);
                                break;
                            case "FolderName":
                                pdfTemplateSet.FolderName = httpContext.Request[key];
                                break;
                            case "SubFolderName":
                                pdfTemplateSet.SubFolderName = httpContext.Request[key];
                                break;
                            case "FileNamePart":
                                pdfTemplateSet.FileNamePart = httpContext.Request[key];
                                break;
                            case "ExcelVersion":
                                pdfTemplateSet.ExcelVersion = httpContext.Request[key];
                                break;
                            case "CreatedBy":
                                pdfTemplateSet.CreatedBy = Convert.ToInt32(httpContext.Request[key]);
                                break;
                            case "IsDemo":
                                pdfTemplateSet.IsDemoTemplate = Convert.ToBoolean(httpContext.Request[key]);
                                break;
                            case "FileExtension":
                                pdfTemplateSet.FileExtension = httpContext.Request[key];
                                break;
                            default:
                                break;
                        }
                    }
                }

                templateZip.Save(templateMemoryStream);
                pdfTemplateSet.TemplateFile = dtSelectedPdfFiles;
                excelZip.Save(excelMemoryStream);

                pdfTemplateSet.TemplateFileZip = templateMemoryStream.ToArray();
                pdfTemplateSet.ExcelZip = excelMemoryStream.ToArray();

                response.TemplateId = Helper.CreateTemplate(pdfTemplateSet);
            }
            catch (Exception ex)
            {
                response.Error = ex.Message;
            }
            return response;
        }

        [Route("api/Template/UpdateExcelVersion")]
        [HttpPost()]
        public string UpdateExcelVersion(ExcelVersion excelVersion)
        {
            try
            {
                Helper.UpdateTemplateExcelVersion(excelVersion.TemplateId, excelVersion.UserId, excelVersion.ExcelVersionId, new MemoryStream().ToArray(), excelVersion.fileBytes, excelVersion.FileExtension);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            return "success";
        }
        private void GenerateNewSetPdfFields(Pdf pdfFile, AsposePdf.Document pdfDocument, Stream pdfStream, int userId)
        {
            try
            {
                if (_dsNewSetPdf == null)
                    _dsNewSetPdf = new DataSet();
                else
                {
                    foreach (DataTable dt in _dsNewSetPdf.Tables)
                    {
                        if (dt.TableName.Contains(pdfFile.Folder))
                            return;
                    }
                }
                var mappedFieldsDataSourcedt = new DataTable();
                mappedFieldsDataSourcedt.Clear();
                CreateMappedFieldsDataSourceDt(ref mappedFieldsDataSourcedt);


                if (pdfFile.IsXFA)
                {
                    XmlNode nodeDataSet = pdfDocument.Form.XFA.Datasets;

                    string[] sam = pdfDocument.Form.XFA.FieldNames;

                    if (nodeDataSet.ChildNodes.Count == 1)
                    {
                        //if (pdfDocument.Form.XFA.XDP != null)
                        //{

                        //    AddFieldsMapping(pdfDocument, nodeDataSet.ChildNodes.Count, nodeDataSet.ChildNodes[0], mappedFieldsDataSourcedt, pdfFile);
                        //}
                        AddFieldsMapping(pdfDocument, nodeDataSet.ChildNodes.Count, nodeDataSet.ChildNodes[0], mappedFieldsDataSourcedt, pdfFile);
                    }
                    else
                    {
                        AddFieldsMapping(pdfDocument, nodeDataSet.ChildNodes.Count, nodeDataSet.ChildNodes[1], mappedFieldsDataSourcedt, pdfFile);
                    }
                }
                else
                {
                    AsposePdf.Facades.Form form = new AsposePdf.Facades.Form(pdfStream);
                    var ss = pdfDocument.Form.Fields.ToList();
                    string[] sam = pdfDocument.Form.XFA.FieldNames;

                    var folderPath = HttpContext.Current.Request.MapPath("~/Content/" + userId);
                    if (!Directory.Exists(folderPath))
                    {
                        Directory.CreateDirectory(folderPath);
                    }
                    var path = $"~/Content/{userId}/Template.xml";
                    var xmlFilePath = HttpContext.Current.Request.MapPath(path);

                    using (FileStream xmlFileStream = new FileStream(xmlFilePath, FileMode.Create, FileAccess.ReadWrite))
                    {
                        form.ExportXml(xmlFileStream);
                    }

                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.Load(xmlFilePath);

                    var parentDataRow = mappedFieldsDataSourcedt.NewRow();
                    parentDataRow["PDFField Name"] = xmlDoc.DocumentElement.Name;
                    parentDataRow["FilePath"] = pdfFile.Folder;
                    var parentId = Guid.NewGuid().ToString("N");
                    parentDataRow["FieldId"] = parentId;
                    parentDataRow["ParentFieldId"] = null;
                    parentDataRow["Is Mapped"] = false;
                    parentDataRow["IsDynamic"] = false;
                    parentDataRow["HasChildFields"] = xmlDoc.DocumentElement.HasChildNodes;
                    parentDataRow["XPath"] = GetXPathToNode(xmlDoc.DocumentElement);
                    mappedFieldsDataSourcedt.Rows.Add(parentDataRow);

                    if (xmlDoc.DocumentElement.HasChildNodes)
                    {
                        var nodeList = xmlDoc.DocumentElement.ChildNodes;
                        for (var i = 0; i <= xmlDoc.DocumentElement.ChildNodes.Count - 1; i++)
                        {
                            var childNode = xmlDoc.DocumentElement.ChildNodes[i];
                            var xpath = GetXPathToNode(childNode);
                            var dataRow = mappedFieldsDataSourcedt.NewRow();
                            if (childNode.Attributes != null && childNode.Attributes["name"] != null)
                                dataRow["PDFField Name"] = childNode.Attributes["name"].Value;
                            else
                                dataRow["PDFField Name"] = childNode.Name;
                            dataRow["FilePath"] = pdfFile.Folder;
                            dataRow["FieldId"] = _previousParentId = Guid.NewGuid().ToString("N");
                            dataRow["ParentFieldId"] = parentId;
                            dataRow["Is Mapped"] = false;

                            // dataRow["IsDynamic"] = false;

                            if (pdfDocument.Form.XFA.XDP != null)
                            {
                                pdfFile.IsXFA = true;
                                bool ISDynamic = Dynamicfield(pdfDocument, xpath, childNode);
                                dataRow["IsDynamic"] = ISDynamic; //find childNode.Name fieldname 
                            }
                            else
                            {
                                dataRow["IsDynamic"] = false;
                            }


                            dataRow["HasChildFields"] = childNode.HasChildNodes && childNode.ChildNodes[0].Name.ToLower() != "value" ? true : false;
                            dataRow["XPath"] = xpath;
                            mappedFieldsDataSourcedt.Rows.Add(dataRow);

                            if (childNode.HasChildNodes && childNode.ChildNodes[0].Name.ToLower() != "value")
                            {
                                AddStaticFieldsMapping(pdfDocument, childNode, mappedFieldsDataSourcedt, pdfFile, _previousParentId);
                            }
                        }
                    }
                }
                mappedFieldsDataSourcedt.TableName = pdfFile.Folder;
                _dsNewSetPdf.Tables.Add(mappedFieldsDataSourcedt);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void AddStaticFieldsMapping(AsposePdf.Document pdfDocument, XmlNode xmlNode, DataTable mappedFieldsDataSourcedt, Pdf pdfFile, string parentId, bool isDynamicParent=false)
        {
            _previousParentId = parentId;
            _previousParentName = xmlNode.ParentNode.Name;
            for (var i = 0; i <= xmlNode.ChildNodes.Count - 1; i++)
            {
                var childNode = xmlNode.ChildNodes[i];
                var parent = childNode.ParentNode;

                if (childNode.NodeType != XmlNodeType.Text && childNode.Name.ToLower() != "value")
                {
                    if (parent.Name != _previousParentName)
                    {
                        for (int j = 0; j < mappedFieldsDataSourcedt.Rows.Count; j++)
                        {
                            if (Convert.ToString(mappedFieldsDataSourcedt.Rows[j]["PDFField Name"]) == parent.Name)
                            {
                                _previousParentId = Convert.ToString(mappedFieldsDataSourcedt.Rows[j]["FieldId"]);
                            }
                        }
                    }
                    var xpath = GetXPathToNode(childNode);
                    var dataRow = mappedFieldsDataSourcedt.NewRow();
                    if (childNode.Attributes != null && childNode.Attributes["name"] != null)
                        dataRow["PDFField Name"] = childNode.Attributes["name"].Value;
                    dataRow["PDFField Name"] = childNode.Name;
                    dataRow["FilePath"] = pdfFile.Folder;
                    dataRow["FieldId"] = Guid.NewGuid().ToString("N");
                    dataRow["ParentFieldId"] = _previousParentId;
                    dataRow["Is Mapped"] = false;
                    //bool ISDynamic = Dynamicfield(pdfDocument, xpath, childNode);
                    if(!isDynamicParent)
                        isDynamicParent = Dynamicfield(pdfDocument, xpath, childNode);
                    dataRow["IsDynamic"] = isDynamicParent ;//false;
                    dataRow["HasChildFields"] = childNode.HasChildNodes;
                    dataRow["XPath"] = xpath;
                    mappedFieldsDataSourcedt.Rows.Add(dataRow);

                    AddStaticFieldsMapping(pdfDocument, childNode, mappedFieldsDataSourcedt, pdfFile, _previousParentId, isDynamicParent);
                }
            }
        }

        public bool Dynamicfield(AsposePdf.Document pdfDocument, string path, XmlNode xmlNode_childNode)
        {
            if (pdfDocument.Form.XFA.XDP != null)
            {
                #region Dynamic Express
                XmlDocument XmlDoc = pdfDocument.Form.XFA.XDP;
                //XmlNodeList xmlNodes = XmlDoc.GetElementsByTagName("subform");
                //foreach (XmlNode node in xmlNodes)
                //{
                //    string str = node.Attributes["name"].Value;
                //    if (str == xmlNode_childNode.Name)
                //    {
                //        foreach (XmlNode child in node.ChildNodes)
                //        {
                //            if (child.Name == "occur" && (child.Attributes["max"].Value == "-1" || Convert.ToInt32(child.Attributes["max"].Value) > 0))
                //            {
                //                return true;
                //            }
                //        }
                //    }
                //}
                //
                //return false;

                XDocument xdoc;
                using (var nodeReader = new XmlNodeReader(XmlDoc))
                {
                    nodeReader.MoveToContent();
                    xdoc = XDocument.Load(nodeReader);
                }

                var fields = xdoc.Root.Descendants();
                var flag = false;

                foreach (var node in fields)
                {
                    var dynamic = node.Attribute("name");
                    if (dynamic != null && dynamic.Value == xmlNode_childNode.Name)
                    {
                        foreach (var n in node.Elements())
                        {
                            if (n.Name.LocalName == "assist" && n.Attribute("role").Value == "TH")
                            {
                                break;
                            }
                            if (n.Name.LocalName == "occur" && (n.Attribute("max").Value == "-1" || Convert.ToInt32(n.Attribute("max").Value) > 0))
                            {
                                flag = true;
                            }
                        }
                    }
                }
                return flag;
                #endregion
            }
            else
            {
                return false;
            }
        }

        public bool DynamicfieldIsXFA(AsposePdf.Document pdfDocument, string path, XmlNode xmlNode_childNode)
        {
            if (pdfDocument.Form.XFA.XDP != null)
            {
                #region Dynamic Express
                XmlDocument XmlDoc = pdfDocument.Form.XFA.XDP;
                XmlNodeList xmlNodes = XmlDoc.GetElementsByTagName("subform");
                XDocument xdoc;
                using (var nodeReader = new XmlNodeReader(XmlDoc))
                {
                    nodeReader.MoveToContent();
                    xdoc = XDocument.Load(nodeReader);
                }

                var fields = xdoc.Root.Descendants();
                var flag = false;

                foreach (var node in fields)
                {
                    var dynamic = node.Attribute("name");
                    if (dynamic != null && dynamic.Value == xmlNode_childNode.Name)
                    { 
                        foreach(var n in node.Elements())
                        {
                            if(n.Name.LocalName == "assist" && n.Attribute("role")!=null && n.Attribute("role").Value == "TH")
                            {
                                break;
                            }
                            if (n.Name.LocalName == "occur" && (n.Attribute("max").Value == "-1" || Convert.ToInt32(n.Attribute("max").Value) > 0))
                            {
                                flag = true;
                            }
                        }
                    }
                }
                //foreach (XmlNode node in xmlNodes)
                //{
                //    if (node.Attributes.GetNamedItem("name") != null)
                //    {
                //        string str = node.Attributes.GetNamedItem("name").Value;
                //        if (str == xmlNode_childNode.Name)
                //        {
                //            foreach (XmlNode child in node.ChildNodes)
                //            {
                //                if (child.Name == "occur" && (child.Attributes["max"].Value == "-1" || Convert.ToInt32(child.Attributes["max"].Value) > 0))
                //                {
                //                    flag = true;
                //                }
                //            }
                //        }
                //    }
                //}
                #endregion
                return flag;
            }
            else
            {
                return false;
            }
        }

        private void AddFieldsMapping(AsposePdf.Document pdfDocument, int childNodeCount, XmlNode xmlNode, DataTable mappedFieldsDataSourcedt, Pdf pdfFile,bool isParentDynamic=false)
        {
            XmlNode newNode;
            XmlNodeList nodeList;
            int i;

            if (xmlNode.HasChildNodes)
            {
                nodeList = xmlNode.ChildNodes;
                var isDynamic = false;

                for (i = 0; i <= xmlNode.ChildNodes.Count - 1; i++)
                {
                    newNode = xmlNode.ChildNodes[i];

                    if (newNode.NodeType != XmlNodeType.Text)
                    {
                        if (newNode.Attributes.Count > 0 && newNode.Attributes["dd:maxOccur"] != null && (newNode.Attributes["dd:maxOccur"].Value == "-1" || newNode.Attributes["dd:maxOccur"].Value == "unbounded" || Convert.ToInt32(newNode.Attributes["dd:maxOccur"].Value) > 0))
                            isParentDynamic = true;
                        var parent = newNode.ParentNode;
                        if (parent.Name == "dd:dataDescription" || parent.Name == "xfa:data")
                            _previousParentId = null;
                        else if (parent.Name != _previousParentName)
                        {
                            for (int j = 0; j < mappedFieldsDataSourcedt.Rows.Count; j++)
                            {
                                if (Convert.ToString(mappedFieldsDataSourcedt.Rows[j]["PDFField Name"]) == parent.Name)
                                {
                                    _previousParentId = Convert.ToString(mappedFieldsDataSourcedt.Rows[j]["FieldId"]);
                                }
                            }
                        }
                        var xpath = GetXPath(newNode);
                        if (childNodeCount == 1)
                            xpath = xpath.Substring(xpath.LastIndexOf("xfa:data") + "xfa:data".Length);
                        else
                            xpath = xpath.Substring(xpath.IndexOf("dataDescription") + "dataDescription".Length);
                        var dataRow = mappedFieldsDataSourcedt.NewRow();
                        dataRow["XPath"] = xpath;
                        dataRow["PDFField Name"] = newNode.Name;
                        dataRow["FilePath"] = pdfFile.Folder;
                        dataRow["FieldId"] = Guid.NewGuid().ToString("N");
                        dataRow["ParentFieldId"] = _previousParentId;
                        dataRow["Is Mapped"] = false;
                        //bool ISDynamic = DynamicfieldIsXFA(pdfDocument, xpath, xmlNode);
                        if(!isParentDynamic)
                            isParentDynamic = DynamicfieldIsXFA(pdfDocument, xpath, newNode);
                        dataRow["IsDynamic"] = isParentDynamic; //find childNode.Name fieldname 
                        //dataRow["IsDynamic"] = isDynamic;
                        dataRow["HasChildFields"] = newNode.HasChildNodes && newNode.ChildNodes[0].NodeType != XmlNodeType.Text ? true : false;
                        mappedFieldsDataSourcedt.Rows.Add(dataRow);
                        _previousParentName = parent.Name;
                        isDynamic = false;
                        AddFieldsMapping(pdfDocument, childNodeCount, newNode, mappedFieldsDataSourcedt, pdfFile, isParentDynamic);
                    }
                }
            }
        }
        private string GetXPath(XmlNode node)
        {
            if (node.NodeType == XmlNodeType.Attribute)
                return String.Format("{0}/@{1}", GetXPathToNode(((XmlAttribute)node).OwnerElement), node.Name);
            if (node.ParentNode == null)
                return "";

            int indexInParent = 1;
            XmlNode siblingNode = node.PreviousSibling;
            while (siblingNode != null)
            {
                if (siblingNode.Name == node.Name)
                {
                    indexInParent++;
                }
                siblingNode = siblingNode.PreviousSibling;
            }

            return String.Format("//{0}//{1}", GetXPath(node.ParentNode), node.Name);
        }
        private void CreateMappedFieldsDataSourceDt(ref DataTable mappedFieldsDataSourceDt)
        {
            mappedFieldsDataSourceDt.Columns.Add("TemplateFileMappingId", typeof(int));
            mappedFieldsDataSourceDt.Columns.Add("TemplateId", typeof(int));
            mappedFieldsDataSourceDt.Columns.Add("TemplateFileId", typeof(int));
            mappedFieldsDataSourceDt.Columns.Add("PDFField Name", typeof(string));
            mappedFieldsDataSourceDt.Columns.Add("Sheet Name", typeof(string));
            mappedFieldsDataSourceDt.Columns.Add("Excel Table Name", typeof(string));
            mappedFieldsDataSourceDt.Columns.Add("Is Mapped", typeof(bool));
            mappedFieldsDataSourceDt.Columns.Add("FilePath", typeof(string));
            mappedFieldsDataSourceDt.Columns.Add("FieldId", typeof(string));
            mappedFieldsDataSourceDt.Columns.Add("ParentFieldId", typeof(string));
            mappedFieldsDataSourceDt.Columns.Add("IsDynamic", typeof(bool));
            mappedFieldsDataSourceDt.Columns.Add("HasChildFields", typeof(bool));
            mappedFieldsDataSourceDt.Columns.Add("XPath", typeof(string));
        }
        private string GetXPathToNode(XmlNode node)
        {
            if (node.NodeType == XmlNodeType.Attribute)
                return String.Format("{0}/@{1}", GetXPathToNode(((XmlAttribute)node).OwnerElement), node.Name);
            if (node.ParentNode == null)
                return "";

            int indexInParent = 1;
            XmlNode siblingNode = node.PreviousSibling;
            while (siblingNode != null)
            {
                if (siblingNode.Name == node.Name)
                {
                    indexInParent++;
                }
                siblingNode = siblingNode.PreviousSibling;
            }

            return String.Format("{0}/{1}[{2}]", GetXPathToNode(node.ParentNode), node.Name, indexInParent);
        }
        private void CreateTemplateFieldsDataSourceDt(DataTable dtTemplateField)
        {
            dtTemplateField.Columns.Add("TemplateFileMappingId", typeof(int));
            dtTemplateField.Columns.Add("TemplateId", typeof(int));
            dtTemplateField.Columns.Add("TemplateFileId", typeof(int));
            dtTemplateField.Columns.Add("PDFField Name", typeof(string));
            dtTemplateField.Columns.Add("Is Mapped", typeof(bool));
            dtTemplateField.Columns.Add("FilePath", typeof(string));
            dtTemplateField.Columns.Add("Sheet Name", typeof(string));
            dtTemplateField.Columns.Add("Excel Table Name", typeof(string));
            dtTemplateField.Columns.Add("FieldId", typeof(string));
            dtTemplateField.Columns.Add("ParentFieldId", typeof(string));
            dtTemplateField.Columns.Add("IsDynamic", typeof(bool));
            dtTemplateField.Columns.Add("HasChildFields", typeof(bool));
            dtTemplateField.Columns.Add("XPath", typeof(string));
        }
        private void AddFieldsMapping(int childNodeCount, XmlNode xmlNode, DataTable mappedFieldsDataSourcedt, PdfFile pdfFile)
        {
            XmlNode newNode;
            XmlNodeList nodeList;
            int i;

            if (xmlNode.HasChildNodes)
            {
                nodeList = xmlNode.ChildNodes;
                var isDynamic = false;

                for (i = 0; i <= xmlNode.ChildNodes.Count - 1; i++)
                {
                    newNode = xmlNode.ChildNodes[i];

                    if (newNode.NodeType != XmlNodeType.Text)
                    {
                        if (newNode.Attributes.Count > 0 && newNode.Attributes["dd:maxOccur"] != null && newNode.Attributes["dd:maxOccur"].Value == "-1")
                            isDynamic = true;
                        var parent = newNode.ParentNode;
                        if (parent.Name == "dd:dataDescription" || parent.Name == "xfa:data")
                            _previousParentId = null;
                        else if (parent.Name != _previousParentName)
                        {
                            for (int j = 0; j < mappedFieldsDataSourcedt.Rows.Count; j++)
                            {
                                if (Convert.ToString(mappedFieldsDataSourcedt.Rows[j]["PDFField Name"]) == parent.Name)
                                {
                                    _previousParentId = Convert.ToString(mappedFieldsDataSourcedt.Rows[j]["FieldId"]);
                                }
                            }
                        }
                        var xpath = GetXPath(newNode);
                        if (childNodeCount == 1)
                            xpath = xpath.Substring(xpath.LastIndexOf("xfa:data") + "xfa:data".Length);
                        else
                            xpath = xpath.Substring(xpath.IndexOf("dataDescription") + "dataDescription".Length);
                        var dataRow = mappedFieldsDataSourcedt.NewRow();
                        dataRow["XPath"] = xpath;
                        dataRow["PDFField Name"] = newNode.Name;
                        dataRow["FilePath"] = pdfFile.FilePath;
                        dataRow["FieldId"] = Guid.NewGuid().ToString("N");
                        dataRow["ParentFieldId"] = _previousParentId;
                        dataRow["Is Mapped"] = false;
                        dataRow["IsDynamic"] = isDynamic;
                        dataRow["HasChildFields"] = newNode.HasChildNodes && newNode.ChildNodes[0].NodeType != XmlNodeType.Text ? true : false;
                        mappedFieldsDataSourcedt.Rows.Add(dataRow);
                        _previousParentName = parent.Name;
                        isDynamic = false;
                        AddFieldsMapping(childNodeCount, newNode, mappedFieldsDataSourcedt, pdfFile);
                    }
                }
            }
        }
        private void GeneratePdfFields(Pdf pdfFile, AsposePdf.Document pdfDocument, Stream pdfStream, int userId)
        {
            try
            {
                //if (!pdfFile.TemplateFileId.HasValue)
                //{
                if (_dsPdf == null)
                    _dsPdf = new DataSet();
                else
                {
                    foreach (DataTable dt in _dsPdf.Tables)
                    {
                        if (dt.TableName.Contains(pdfFile.Folder))
                            return;
                    }
                }

                var dtMappedPdfFieldDataSource = new DataTable();
                dtMappedPdfFieldDataSource.Clear();
                CreateTemplateFieldsDataSourceDt(dtMappedPdfFieldDataSource);

                if (pdfFile.IsXFA)
                {
                    XmlNode nodeDataSet = pdfDocument.Form.XFA.Datasets;
                    if (nodeDataSet.ChildNodes.Count == 1)
                        AddFieldsMapping(pdfDocument, nodeDataSet.ChildNodes.Count, nodeDataSet.ChildNodes[0], dtMappedPdfFieldDataSource, pdfFile);
                    else
                        AddFieldsMapping(pdfDocument, nodeDataSet.ChildNodes.Count, nodeDataSet.ChildNodes[1], dtMappedPdfFieldDataSource, pdfFile);
                }
                else
                {
                    AsposePdf.Facades.Form form = new AsposePdf.Facades.Form(pdfStream);

                    var folderPath = HttpContext.Current.Request.MapPath("~/Content/" + userId);
                    if (!Directory.Exists(folderPath))
                    {
                        Directory.CreateDirectory(folderPath);
                    }
                    var path = $"~/Content/{userId}/Template.xml";
                    var xmlFilePath = HttpContext.Current.Request.MapPath(path);

                    using (FileStream xmlFileStream = new FileStream(xmlFilePath, FileMode.Create, FileAccess.ReadWrite))
                    {
                        form.ExportXml(xmlFileStream);
                    }

                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.Load(xmlFilePath);

                    var parentDataRow = dtMappedPdfFieldDataSource.NewRow();
                    parentDataRow["PDFField Name"] = xmlDoc.DocumentElement.Name;
                    parentDataRow["FilePath"] = pdfFile.Folder;
                    var parentId = Guid.NewGuid().ToString("N");
                    parentDataRow["FieldId"] = parentId;
                    parentDataRow["ParentFieldId"] = null;
                    parentDataRow["Is Mapped"] = false;
                    parentDataRow["IsDynamic"] = false;
                    parentDataRow["HasChildFields"] = xmlDoc.DocumentElement.HasChildNodes;
                    parentDataRow["XPath"] = GetXPathToNode(xmlDoc.DocumentElement);
                    dtMappedPdfFieldDataSource.Rows.Add(parentDataRow);

                    if (xmlDoc.DocumentElement.HasChildNodes)
                    {
                        var nodeList = xmlDoc.DocumentElement.ChildNodes;
                        for (var i = 0; i <= xmlDoc.DocumentElement.ChildNodes.Count - 1; i++)
                        {
                            var childNode = xmlDoc.DocumentElement.ChildNodes[i];
                            var xpath = GetXPathToNode(childNode);
                            var dataRow = dtMappedPdfFieldDataSource.NewRow();
                            if (childNode.Attributes != null && childNode.Attributes["name"] != null)
                                dataRow["PDFField Name"] = childNode.Attributes["name"].Value;
                            dataRow["FilePath"] = pdfFile.Folder;
                            dataRow["FieldId"] = Guid.NewGuid().ToString("N");
                            dataRow["ParentFieldId"] = parentId;
                            dataRow["Is Mapped"] = false;
                            //dataRow["IsDynamic"] = false;
                            if (pdfDocument.Form.XFA.XDP != null)
                            {
                                pdfFile.IsXFA = true;
                                bool ISDynamic = Dynamicfield(pdfDocument, xpath, childNode);
                                dataRow["IsDynamic"] = ISDynamic; //find childNode.Name fieldname 
                            }
                            else
                            {
                                dataRow["IsDynamic"] = false;
                            }

                            //dataRow["HasChildFields"] = false;
                            dataRow["HasChildFields"] = childNode.HasChildNodes && childNode.ChildNodes[0].Name.ToLower() != "value" ? true : false;
                            dataRow["XPath"] = xpath;
                            dtMappedPdfFieldDataSource.Rows.Add(dataRow);
                            if (childNode.HasChildNodes && childNode.ChildNodes[0].Name.ToLower() != "value")
                            {
                                AddStaticFieldsMapping(pdfDocument, childNode, dtMappedPdfFieldDataSource, pdfFile, _previousParentId);
                            }
                        }
                    }
                }

                dtMappedPdfFieldDataSource.TableName = pdfFile.Folder;
                _dsPdf.Tables.Add(dtMappedPdfFieldDataSource);
                //}
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [Route("api/Template/UpdateTemplate")]
        [HttpPost()]
        public string UpdateTemplate()
        {
            var dtMappedPdfFields = new DataTable();
            var dtTemplate = new DataTable();
            var dtTemplateFile = new DataTable();
            ZipFile templateZip = new ZipFile();
            var templateMemoryStream = new MemoryStream();
            var templateId = 0;
            var templateName = string.Empty;
            var description = string.Empty;
            var subFolderName = string.Empty;
            var fileNamePart = string.Empty;
            var deletedTemplateFileIds = string.Empty;
            var updatedBy = 0;

            try
            {
                CreateTemplateFieldsDt(dtMappedPdfFields);
                CreateTemplateFileDt(dtTemplateFile);
                CreateEditTemplateDt(dtTemplate);

                var httpContext = HttpContext.Current;

                if (httpContext.Request.Files.Count > 0)
                {
                    string[] excludeFilesArray = null;
                    var exclueFiles = httpContext.Request["ExcludeFiles"];

                    if (!string.IsNullOrWhiteSpace(exclueFiles))
                        excludeFilesArray = exclueFiles.Split(',');

                    for (int j = 0; j < httpContext.Request.Files.Count; j++)
                    {
                        //PdfFile pdfFile = new PdfFile();
                        Pdf pdfFile = new Pdf();
                        pdfFile.FileName = null;
                        HttpPostedFile httpPostedFile = httpContext.Request.Files[j];

                        if (httpPostedFile != null && httpPostedFile.ContentType == "application/pdf")
                        {
                            var isFileExclude = false;
                            pdfFile.FileName = httpPostedFile.FileName.Substring(httpPostedFile.FileName.LastIndexOf("\\") + 1);

                            if (excludeFilesArray != null)
                            {
                                foreach (var excludeFile in excludeFilesArray)
                                {
                                    if (pdfFile.FileName == excludeFile)
                                    {
                                        isFileExclude = true;
                                        break;
                                    }
                                }
                            }

                            if (isFileExclude)
                                continue;

                            AsposePdf.Document pdfDocument = new AsposePdf.Document(httpPostedFile.InputStream);

                            if (pdfDocument.Form.Type == AsposePdf.Forms.FormType.Dynamic)
                                pdfFile.IsXFA = true;

                            pdfFile.Folder = httpPostedFile.FileName;
                            pdfFile.FileName = httpPostedFile.FileName.Substring(httpPostedFile.FileName.LastIndexOf("\\") + 1);

                            var userId = Convert.ToInt32(httpContext.Request["UpdatedBy"]);
                            GenerateNewSetPdfFields(pdfFile, pdfDocument, httpPostedFile.InputStream, userId);

                            var dataRow = dtTemplateFile.NewRow();
                            dataRow["FileName"] = pdfFile.FileName;
                            dataRow["FilePath"] = pdfFile.Folder;
                            dataRow["IsXFA"] = pdfFile.IsXFA;
                            if (httpContext.Request.Form.Get("TemplateId") == null)
                                dataRow["TemplateId"] = DBNull.Value;
                            else
                                dataRow["TemplateId"] = httpContext.Request.Form.Get("TemplateId");

                            using (var binaryReader = new BinaryReader(httpPostedFile.InputStream))
                            {
                                dataRow["PdfBytes"] = binaryReader.ReadBytes(httpPostedFile.ContentLength);
                            }

                            dtTemplateFile.Rows.Add(dataRow);
                        }
                    }
                }

                templateZip.Save(templateMemoryStream);

                if (_dsNewSetPdf != null)
                {
                    foreach (DataTable dt in _dsNewSetPdf.Tables)
                    {
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            var dataRow = dtMappedPdfFields.NewRow();
                            dataRow["TemplateFileMappingId"] = dt.Rows[i]["TemplateFileMappingId"];
                            dataRow["TemplateId"] = dt.Rows[i]["TemplateId"];
                            dataRow["TemplateFileId"] = dt.Rows[i]["TemplateFileId"];
                            dataRow["PDFFieldName"] = dt.Rows[i]["PDFField Name"];
                            dataRow["IsMapped"] = dt.Rows[i]["Is Mapped"];
                            dataRow["FilePath"] = dt.Rows[i]["FilePath"];
                            dataRow["SheetName"] = dt.Rows[i]["Sheet Name"];
                            dataRow["ExcelTableName"] = dt.Rows[i]["Excel Table Name"];
                            dataRow["FieldId"] = dt.Rows[i]["FieldId"];
                            dataRow["ParentFieldId"] = dt.Rows[i]["ParentFieldId"];
                            dataRow["IsDynamic"] = dt.Rows[i]["IsDynamic"];
                            dataRow["HasChildFields"] = dt.Rows[i]["HasChildFields"];
                            dataRow["XPath"] = dt.Rows[i]["XPath"];
                            dtMappedPdfFields.Rows.Add(dataRow);
                        }
                    }
                }

                foreach (var key in httpContext.Request.Form.AllKeys)
                {
                    switch (key)
                    {
                        case "TemplateId":
                            templateId = Convert.ToInt32(httpContext.Request[key]);
                            break;
                        case "TemplateName":
                            templateName = httpContext.Request[key];
                            break;
                        case "Description":
                            description = httpContext.Request[key];
                            break;
                        case "SubFolderName":
                            subFolderName = httpContext.Request[key];
                            break;
                        case "FileNamePart":
                            fileNamePart = httpContext.Request[key];
                            break;
                        case "DeletedTemplateFileIds":
                            deletedTemplateFileIds = httpContext.Request[key];
                            break;
                        case "UpdatedBy":
                            updatedBy = Convert.ToInt32(httpContext.Request[key]);
                            break;
                        default:
                            break;
                    }
                }

                dtTemplate = Helper.GetTemplateByTemplateId(templateId);
                dtTemplate.Rows[0]["TemplateName"] = templateName;
                dtTemplate.Rows[0]["Description"] = description;
                dtTemplate.Rows[0]["SubFolderName"] = subFolderName;
                dtTemplate.Rows[0]["FileNamePart"] = fileNamePart;
                dtTemplate.AcceptChanges();

                var editPdfTemplate = new EditPdfTemplate()
                {
                    DeletedTemplateFileIds = deletedTemplateFileIds,
                    UpdatedBy = updatedBy,
                    Template = dtTemplate,
                    TemplateFile = dtTemplateFile,
                    TemplateFileFieldMapping = dtMappedPdfFields,
                    TemplateFileZip = templateMemoryStream.ToArray()
                };

                Helper.UpdateTemplate(editPdfTemplate);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            return "success";
        }

        [Route("api/Template/DeleteTemplateFile")]
        [HttpPost()]
        public string DeleteTemplateFile([FromBody] int fileId)
        {
            try
            {
                Helper.DeleteTemplateFile(fileId);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            return "success";
        }

        [Route("api/Template/UploadExcelFile")]
        [HttpPost()]
        public string UploadExcelFile(ExcelFile excelFile)
        {
            try
            {
                Helper.UploadExcelFile(excelFile);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            return "success";
        }

        [Route("api/Template/BindFieldsMappingTreeView")]
        [HttpGet]
        public object BindFieldsMappingTreeView(string id, int fileId)
        {
            try
            {
                var mappedPdfFields = Helper.GetTemplateFieldsByTemplateFileId(fileId, true);

                List<JsTreeAttribute> core = new List<JsTreeAttribute>();
                if (id == "#")
                {
                    foreach (var mappedPdfField in mappedPdfFields)
                    {
                        if (string.IsNullOrEmpty(mappedPdfField.ParentFieldId))
                        {
                            JsTreeAttribute obj = new JsTreeAttribute()
                            {
                                templateFileMappingId = mappedPdfField.TemplateFileMappingId,
                                id = mappedPdfField.FieldId,
                                text = mappedPdfField.PDFFieldName,
                                children = true,
                                type = "Folder",
                                isMapped = mappedPdfField.IsMapped
                            };
                            core.Add(obj);
                            break;
                        }
                    }
                }
                else
                {
                    var field = mappedPdfFields.FirstOrDefault(mp => mp.FieldId == id);
                    var lstmappedField = mappedPdfFields.Where(mp => mp.ParentFieldId == id).ToList();
                    foreach (var mappedField in lstmappedField)
                    {
                        JsTreeAttribute obj = new JsTreeAttribute()
                        {
                            templateFileMappingId = mappedField.TemplateFileMappingId,
                            id = mappedField.FieldId,
                            text = mappedField.PDFFieldName,
                            children = mappedField.HasChildFields,
                            title = field.IsDynamic ? "Dynamic Field" : mappedField.IsDynamic ? "Dynamic Section" : mappedField.HasChildFields ? null : "Static Field",
                            /*icon = mappedField.IsMapped && mappedField.IsDynamic ? "glyphicon glyphicon-duplicate text-success" : !mappedField.IsMapped && mappedField.IsDynamic ? "glyphicon glyphicon-duplicate text-danger" : mappedField.HasChildFields ? null : field.IsDynamic && field.IsMapped ? "glyphicon glyphicon-file text-success" : field.IsDynamic && !field.IsMapped ? "glyphicon glyphicon-file text-danger" : !field.IsDynamic && field.IsMapped ? "glyphicon glyphicon-file text-success" : "jstree-file",*/
                            type = mappedField.IsDynamic ? "DF" : mappedField.HasChildFields ? "Folder" : field.IsDynamic ? "DE" : "SF",
                            isMapped = mappedField.IsMapped
                        };

                        if (obj.title == "Static Field")
                        {
                            if (mappedField.IsMapped)
                                obj.icon = "glyphicon glyphicon-file text-success";
                            else
                                obj.icon = "glyphicon glyphicon-file text-danger";
                        }
                        else if (obj.title == "Dynamic Field")
                        {
                            if (mappedField.IsMapped)
                                obj.icon = "glyphicon glyphicon-duplicate text-success";
                            //obj.icon = "glyphicon glyphicon-file text-success";
                            else
                                obj.icon = "glyphicon glyphicon-duplicate text-danger";
                            //obj.icon = "glyphicon glyphicon-file text-danger";
                        }
                        else if (obj.title == "Dynamic Section")
                        {
                            if (mappedField.IsMapped)
                                obj.icon = "glyphicon glyphicon-align-justify text-success"; //"glyphicon glyphicon-duplicate text-success";
                            else
                                obj.icon = "glyphicon glyphicon-align-justify text-danger"; //"glyphicon glyphicon-duplicate text-danger";

                            if (mappedField.HasChildFields == false)
                            {
                                obj.type = "Folder";
                                obj.icon = null;
                                obj.title = null;
                            }
                        }
                        core.Add(obj);
                    }
                }
                return Json<List<JsTreeAttribute>>(core);
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        [Route("api/Template/RemoveStaticFieldMapping")]
        [HttpPost()]
        public string RemoveStaticFieldMapping([FromBody] int templateFileMappingId)
        {
            try
            {
                Helper.RemoveMappedField(templateFileMappingId, true, string.Empty);
                return "success";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        [Route("api/Template/RemoveDynamicFieldMapping")]
        [HttpPost()]
        public string RemoveDynamicFieldMapping(DynamicParam dynamicParam)
        {
            try
            {
                Helper.RemoveDynamicFieldMapping(dynamicParam.TemplateFileMappingId, dynamicParam.IsDynamicField);
                return "success";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        [Route("api/Template/RemoveAllMappings")]
        [HttpPost()]
        public RemoveMappingsResponse RemoveAllMappings([FromBody] int templateId)
        {
            var response = new RemoveMappingsResponse();
            try
            {
                if (!Helper.IsTemplateFieldsMapped(templateId))
                    response.IsAnyFieldMapped = false;
                else
                {
                    Helper.RemoveAllMappedFields(templateId);
                    response.IsAnyFieldMapped = true;
                }
            }
            catch (Exception ex)
            {
                response.Error = ex.Message;
            }
            return response;
        }

        [Route("api/Template/MapXfaField")]
        [HttpPost()]
        public string MapXfaField(MapFieldParam mapFieldParam)
        {
            try
            {
                Helper.AddXFAMappedField(mapFieldParam);
                return "success";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        [Route("api/Template/MapAcroField")]
        [HttpPost()]
        public string MapAcroField(MapFieldParam mapFieldParam)
        {
            try
            {
                Helper.AddAcroMappedField(mapFieldParam);
                return "success";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        [Route("api/Template/RemoveFileMappings")]
        [HttpPost()]
        public RemoveMappingsResponse RemoveFileMappings([FromBody] int templateFileId)
        {
            var response = new RemoveMappingsResponse();
            try
            {
                if (!Helper.IsTemplateFileFieldsMapped(templateFileId))
                    response.IsAnyFieldMapped = false;
                else
                {
                    Helper.RemoveFileMappedFields(templateFileId);
                    response.IsAnyFieldMapped = true;
                }
            }
            catch (Exception ex)
            {
                response.Error = ex.Message;
            }
            return response;
        }

        [Route("api/Template/AutoMapFields")]
        [HttpGet()]
        public DataTable AutoMapFields(int templateId, int templateFileId = 0)
        {
            var templateFields = new DataTable();
            try
            {
                if(templateFileId <= 0)
                {
                    Helper.RemoveAllMappedFields(templateId);
                }
                
                templateFields = Helper.GetTemplateFieldsByTemplateId(templateId);
                return templateFields;
            }
            catch (Exception ex)
            {
                return null;
            }

        }

        [Route("api/Template/SyncMappedFields")]
        [HttpGet()]
        public DataTable SyncMappedFields(int templateId)
        {
            var templateFields = new DataTable();
            try
            {    
                templateFields = Helper.GetTemplateFieldsByTemplateId(templateId);

                var rows = templateFields.Select("IsMapped= " + false);
                foreach (var row in rows)
                { row.Delete(); }
                templateFields.AcceptChanges();

                return templateFields;
            }
            catch (Exception ex)
            {
                return null;
            }

        }

        [Route("api/Template/BackFromEditFieldsMapping")]
        [HttpGet]
        public EditMapFieldBackResponse BackFromEditFieldsMapping(int templateFileId)
        {
            var editMapFieldBackResponse = new EditMapFieldBackResponse();
            try
            {
                var isAnyFieldMapped = Helper.IsTemplateFileFieldsMapped(templateFileId);
                editMapFieldBackResponse.IsAnyFiedMapped = isAnyFieldMapped;

                var dynamicFieldsCount = Helper.GetDynamicFieldsCount(templateFileId);
                editMapFieldBackResponse.DynamicFieldsCount = dynamicFieldsCount;

                var parentTable = Helper.GetParentTableByFileId(templateFileId);
                editMapFieldBackResponse.ParentTable = parentTable;

                var childTables = Helper.GetChildTablesByFileId(templateFileId);
                editMapFieldBackResponse.ChildTablesCount = childTables.Rows.Count;

                var parentChildRelationshipCount = Helper.GetParentChildTableRelationshipCount(templateFileId);
                editMapFieldBackResponse.ParentChildRelationshipCount = parentChildRelationshipCount;
            }
            catch (Exception ex)
            {
                return null;
            }
            return editMapFieldBackResponse;
        }

        [Route("api/Template/GetSendDataParam")]
        [HttpGet]
        public SendDataParamResponse GetSendDataParam(int templateId)
        {
            var response = new SendDataParamResponse();
            response.Error = null;
            try
            {
                response.Params = Helper.GetSendDataParam(templateId);
            }
            catch (Exception ex)
            {
                response.Error = ex.Message;
            }
            return response;
        }

        [Route("api/Template/IsExcelVersionExist")]
        [HttpGet]
        public IsExcelVersionExistResponse IsExcelVersionExist(string excelVersion)
        {
            var response = new IsExcelVersionExistResponse();
            response.Error = null;
            try
            {
                response = Helper.IsExcelVersionExist(excelVersion);
            }
            catch (Exception ex)
            {
                response.Error = ex.Message;
            }
            return response;
        }

        [Route("api/Template/GetNamingOptions")]
        [HttpGet]
        public NamingOptions GetNamingOptions(int templateId)
        {
            var response = new NamingOptions();
            response.Error = null;
            try
            {
                response.Names = Helper.GetTemplateNamesPart(templateId);
            }
            catch (Exception ex)
            {
                response.Error = ex.Message;
            }
            return response;
        }

        [Route("api/Template/SendDataToTemplateSet")]
        [HttpPost]
        public SendToTemplateSetDataResponse SendDataToTemplateSet(SendToTemplateSetData sendData)
        {
            SendToTemplateSetDataResponse response = new SendToTemplateSetDataResponse();
            response.Message = new List<string>();

            try
            {
                var dtMappedFields = Helper.GetTemplateMappedFields(sendData.TemplateId);
                if (dtMappedFields == null || dtMappedFields.Rows.Count == 0)
                {
                    response.Error = "No data mapping is done for the selected Templates Set .\nPlease do the mapping first.";
                    return response;
                }

                var templateFiles = Helper.GetTemplateFiles(sendData.TemplateId);
                var templateName = Convert.ToString(dtMappedFields.Rows[0]["TemplateName"]);
                var subFolder = Convert.ToString(dtMappedFields.Rows[0]["SubFolderName"]);
                var fileNamePart = Convert.ToString(dtMappedFields.Rows[0]["FileNamePart"]);
                var excelColumnNameValue = string.Empty;
                var dynamicParent = string.Empty;
                var sessionFolderName = string.Empty;
                var clientDirectory = string.Empty;
                ZipFile templateZip = new ZipFile();

                if (string.IsNullOrWhiteSpace(sessionFolderName))
                {
                    var subdir = HttpContext.Current.Request.MapPath($"~/{DateTime.Now.ToString("yyyyMMddHHmmss")}FormFilling");

                    if (!Directory.Exists(subdir))
                    {
                        Directory.CreateDirectory(subdir);
                    }
                    sessionFolderName = subdir;
                }
                foreach (var templateFile in templateFiles)
                {
                    var counter = 0;
                    var isPdfMapped = false;
                    var parentTableName = string.Empty;
                    var ds = new DataSet();
                    var pdfOutStream = new MemoryStream();

                    var pdfFilePath = HttpContext.Current.Request.MapPath("~/Content/" + templateFile.FileName);
                    pdfOutStream.Write(templateFile.PdfBytes, 0, templateFile.PdfBytes.Length);

                    //var pdfStream = new MemoryStream(templateFile.PdfBytes);
                    Document pdfDocument = new Document(pdfOutStream);

                    var xmlDoc = new XmlDocument();

                    if (!templateFile.IsXFA)
                    {
                        foreach (AsposePdf.Forms.Field field in pdfDocument.Form)
                        {
                            field.Value = string.Empty;
                        }
                    }
                    pdfDocument.Save(pdfFilePath);
                    pdfDocument = new Document(pdfFilePath);
                    var folderPath = HttpContext.Current.Request.MapPath("~/Content/" + sendData.UserId);
                    if (!Directory.Exists(folderPath))
                    {
                        Directory.CreateDirectory(folderPath);
                    }
                    var path = $"~/Content/{sendData.UserId}/Template.xml";
                    var xmlFilePath = HttpContext.Current.Request.MapPath(path);

                    using (FileStream xmlFileStream = new FileStream(xmlFilePath, FileMode.Create, FileAccess.ReadWrite))
                    {
                        using (var pdfStreamData = File.Open(pdfFilePath, FileMode.Open, FileAccess.ReadWrite))
                        {
                            AsposePdf.Facades.Form form = new AsposePdf.Facades.Form(pdfOutStream);
                            if (templateFile.IsXFA)
                                form.ExtractXfaData(xmlFileStream);
                            else
                                form.ExportXml(xmlFileStream);
                        }
                    }

                    xmlDoc.Load(xmlFilePath);

                    var fileMappedFields = new DataTable();
                    var fields = dtMappedFields.AsEnumerable()
                        .Where(row => row.Field<Int32>("TemplateFileId") == templateFile.TemplateFileId);

                    if (fields.Any())
                        fileMappedFields = fields.CopyToDataTable();
                    else
                        continue;

                    var rowIndex = 0;
                    DataTable pdt = new DataTable();

                    foreach (var ParentTableDataRow in sendData.ParentTableData)
                    {
                        var rowValues = ParentTableDataRow.Split('*');
                        if (rowIndex == 0)
                        {
                            foreach (var value in rowValues)
                            {
                                pdt.TableName = sendData.ParentTableName;
                                pdt.Columns.Add(value);
                            }
                        }
                        else
                        {
                            var columnIndex = 0;
                            var dr = pdt.NewRow();
                            foreach (var value in rowValues)
                            {
                                dr[columnIndex] = value;
                                columnIndex++;
                                if (columnIndex == pdt.Columns.Count)
                                {
                                    pdt.Rows.Add(dr);
                                    dr = pdt.NewRow();
                                    columnIndex = 0;
                                }
                            }

                        }
                        rowIndex++;
                    }
                    ds.Tables.Add(pdt);
                    rowIndex = 0;
                    DataTable cdt = null;
                    var childTablesCount = -1;
                    foreach (var ChildTables in sendData.ChildTableData)
                    {
                        childTablesCount++;
                        cdt = new DataTable();
                        rowIndex = 0;
                        foreach (var ChildTableDataRow in ChildTables)
                        {
                            var rowValues = ChildTableDataRow.Split('*');
                            if (rowIndex == 0)
                            {
                                foreach (var value in rowValues)
                                {
                                    cdt.TableName = sendData.ChildTableNames[childTablesCount];
                                    cdt.Columns.Add(value);
                                }
                            }
                            else
                            {
                                var columnIndex = 0;
                                var dr = cdt.NewRow();
                                foreach (var value in rowValues)
                                {
                                    dr[columnIndex] = value;
                                    columnIndex++;
                                    if (columnIndex == cdt.Columns.Count)
                                    {
                                        cdt.Rows.Add(dr);
                                        dr = cdt.NewRow();
                                        columnIndex = 0;
                                    }
                                }
                            }
                            rowIndex++;
                        }
                        ds.Tables.Add(cdt);
                    }

                    if (templateFile.IsXFA)
                    {
                        counter = 0;
                        var newds = new DataSet();
                        ClearFieldsValue(xmlDoc.ChildNodes[1]);
                        xmlDoc.Save(xmlFilePath);
                        var parentChildTableMapping = Helper.GetParentChildTableMapping(templateFile.TemplateFileId);
                        if (parentChildTableMapping.Rows.Count > 0)
                        {
                            for (var j = 0; j < parentChildTableMapping.Rows.Count; j++)
                            {
                                parentTableName = Convert.ToString(parentChildTableMapping.Rows[j]["ParentTableName"]);
                                var parentTableColumn = Convert.ToString(parentChildTableMapping.Rows[j]["ParentTableColumn"]);
                                var childTableName = Convert.ToString(parentChildTableMapping.Rows[j]["ChildTableName"]);

                                foreach (DataTable dt in ds.Tables)
                                {
                                    if (dt.TableName == parentTableName || dt.TableName == childTableName)
                                    {
                                        if (newds.Tables.Contains(dt.TableName))
                                            continue;
                                        isPdfMapped = true;
                                        newds.Tables.Add(dt.Copy());
                                    }
                                }
                            }

                            foreach (DataTable dt in newds.Tables)
                            {
                                if (dt.TableName == parentTableName)
                                {
                                    for (var k = 0; k < dt.Rows.Count; k++)
                                    {
                                        //ClearFieldsValue(xmlDoc.ChildNodes[1]);
                                        //xmlDoc.Save(xmlFilePath);

                                        var outputSubFolder = subFolder;
                                        var outputFileName = fileNamePart;
                                        var hasFileNamePartId = false;
                                        if (!string.IsNullOrWhiteSpace(fileNamePart) && fileNamePart.ToLower().Contains("id"))
                                            hasFileNamePartId = true;
                                        counter++;

                                        for (var j = 0; j < parentChildTableMapping.Rows.Count; j++)
                                        {
                                            var parentTableColumn = Convert.ToString(parentChildTableMapping.Rows[j]["ParentTableColumn"]);
                                            var childTableName = Convert.ToString(parentChildTableMapping.Rows[j]["ChildTableName"]);
                                            var childTableColumnName = Convert.ToString(parentChildTableMapping.Rows[j]["ChildTableColumn"]);

                                            for (var l = 0; l < dt.Columns.Count; l++)
                                            {
                                                excelColumnNameValue = Convert.ToString(dt.Rows[k][l]);
                                                var excelColumnName = dt.Columns[l].ColumnName;
                                                if (subFolder != null && subFolder.Contains(excelColumnName))
                                                    outputSubFolder = outputSubFolder.Replace(excelColumnName, excelColumnNameValue);
                                                if (fileNamePart != null && fileNamePart.Contains(excelColumnName))
                                                    outputFileName = outputFileName.Replace(excelColumnName, excelColumnNameValue);

                                                if (excelColumnName == parentTableColumn)
                                                {
                                                    var parentKey = excelColumnNameValue;

                                                    var dynamicField = Helper.GetDynamicField(templateFile.TemplateFileId, childTableName);
                                                    if (dynamicField.Rows.Count > 0)
                                                    {
                                                        var dynamicFieldId = Convert.ToString(dynamicField.Rows[0]["FieldId"]);
                                                        var dynamicFieldXPath = Convert.ToString(dynamicField.Rows[0]["XPath"]);
                                                        var dynamicFieldName = Convert.ToString(dynamicField.Rows[0]["PDFFieldName"]);

                                                        var dynamicElements = Helper.GetDynamicChildFieldsByFieldId(dynamicFieldId);

                                                        foreach (DataTable ncdt in newds.Tables)
                                                        {
                                                            if (ncdt.TableName == childTableName)
                                                            {
                                                                var field = ncdt.AsEnumerable()
                                                                    .Where(row => row.Field<string>(childTableColumnName) == parentKey);
                                                                if (field.Any())
                                                                {
                                                                    var fieldDt = field.CopyToDataTable();
                                                                    var xmlNode = xmlDoc.SelectSingleNode(dynamicFieldXPath);
                                                                    var parentNode = xmlNode.ParentNode;
                                                                    var childNodes = xmlDoc.SelectNodes(dynamicFieldXPath);
                                                                    foreach (XmlNode node in childNodes)
                                                                        parentNode.RemoveChild(node);

                                                                    for (var m = 0; m < fieldDt.Rows.Count; m++)
                                                                    {
                                                                        var childKey = string.Empty;
                                                                        XmlElement parentElement = xmlDoc.CreateElement(dynamicFieldName);
                                                                        for (int o = 0; o < dynamicElements.Rows.Count; o++)
                                                                        {
                                                                            var isMapped = Convert.ToBoolean(dynamicElements.Rows[o]["IsMapped"]);
                                                                            var childExcelFieldName = Convert.ToString(dynamicElements.Rows[o]["ExcelFieldName"]);
                                                                            XmlElement element = xmlDoc.CreateElement(Convert.ToString(dynamicElements.Rows[o]["PDFFieldName"]));

                                                                            for (var n = 0; n < fieldDt.Columns.Count; n++)
                                                                            {
                                                                                var childColumnName = fieldDt.Columns[n].ColumnName;
                                                                                var childColumnNameValue = Convert.ToString(fieldDt.Rows[m][n]);
                                                                                if (childTableColumnName == childColumnName)
                                                                                    childKey = childColumnNameValue;

                                                                                if (childExcelFieldName == childColumnName)
                                                                                {
                                                                                    var cellValue = childColumnNameValue;
                                                                                    if (isMapped)
                                                                                        element.InnerText = Convert.ToString(cellValue);
                                                                                }
                                                                            }
                                                                            if (childKey == parentKey)
                                                                                parentElement.AppendChild(element);
                                                                        }
                                                                        if (childKey == parentKey)
                                                                            parentNode.AppendChild(parentElement);
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        for (int p = 0; p < fileMappedFields.Rows.Count; p++)
                                        {
                                            var isDynamic = Convert.ToBoolean(fileMappedFields.Rows[p]["IsDynamic"]);
                                            var fieldId = Convert.ToString(fileMappedFields.Rows[p]["FieldId"]);
                                            var parentFieldId = Convert.ToString(fileMappedFields.Rows[p]["ParentFieldId"]);

                                            if (isDynamic)
                                            {
                                                dynamicParent = fieldId;
                                                continue;
                                            }
                                            if (dynamicParent == parentFieldId)
                                                continue;

                                            var tableName = Convert.ToString(fileMappedFields.Rows[p]["ExcelTableName"]);
                                            var sheetName = Convert.ToString(fileMappedFields.Rows[p]["SheetName"]);
                                            var pdfFieldName = Convert.ToString(fileMappedFields.Rows[p]["PDFFieldName"]);
                                            var excelFieldName = Convert.ToString(fileMappedFields.Rows[p]["ExcelFieldName"]);
                                            var xpath = Convert.ToString(fileMappedFields.Rows[p]["XPath"]);

                                            for (var l = 0; l < dt.Columns.Count; l++)
                                            {
                                                excelColumnNameValue = Convert.ToString(dt.Rows[k][l]);
                                                var excelColumnName = dt.Columns[l].ColumnName;
                                                if (subFolder != null && subFolder.Contains(excelColumnName))
                                                    outputSubFolder = outputSubFolder.Replace(excelColumnName, excelColumnNameValue);
                                                if (fileNamePart != null && fileNamePart.Contains(excelColumnName))
                                                    outputFileName = outputFileName.Replace(excelColumnName, excelColumnNameValue);
                                                var column = excelColumnName.Split(new[] { "--" }, StringSplitOptions.None);
                                                if (excelFieldName == column[0])
                                                {
                                                    var xmlNode = xmlDoc.SelectSingleNode(xpath + "[" + (column.Count() > 1 ? column[1] : "1") + "]");
                                                    xmlNode.InnerText = excelColumnNameValue;
                                                }
                                            }
                                        }

                                        if (isPdfMapped)
                                        {
                                            outputSubFolder = outputSubFolder.Replace("{", "").Replace("}", "").Trim();
                                            outputFileName = outputFileName.Replace("{", "").Replace("}", "").Trim();

                                            if (string.IsNullOrWhiteSpace(outputSubFolder))
                                            {
                                                response.Message.Add($"No output subfolder name data found in the table for {templateFile.FileName}.");
                                                continue;
                                            }
                                            if (string.IsNullOrWhiteSpace(outputFileName))
                                            {
                                                response.Message.Add($"No output file naming options data found in the table for {templateFile.FileName}.");
                                                continue;
                                            }
                                            if (!IsNameValid(outputSubFolder))
                                            {
                                                response.Message.Add($"Output subfolder name is not valid for {templateFile.FileName}.");
                                                continue;
                                            }
                                            if (!IsNameValid(outputFileName))
                                            {
                                                response.Message.Add($"Output file name is not valid for {templateFile.FileName}.");
                                                continue;
                                            }

                                            xmlDoc.Save(xmlFilePath);
                                            using (FileStream xmlStream = new FileStream(xmlFilePath, FileMode.Open, FileAccess.Read))
                                            {
                                                using (var pdfStream = File.Open(pdfFilePath, FileMode.Open, FileAccess.ReadWrite))
                                                {
                                                    AsposePdf.Facades.Form form = new AsposePdf.Facades.Form(pdfStream);
                                                    if (templateFile.IsXFA)
                                                        form.SetXfaData(xmlStream);
                                                    else
                                                        form.ImportXml(xmlStream);
                                                    form.Save(pdfStream);
                                                }
                                            }

                                            if (!Directory.Exists($"{sessionFolderName}/{outputSubFolder}"))
                                            {
                                                clientDirectory = $"{sessionFolderName}/{outputSubFolder}";
                                                try
                                                {
                                                    Directory.CreateDirectory(clientDirectory);
                                                }
                                                catch (PathTooLongException)
                                                {
                                                    response.Message.Add($"The subfolder path:\n\"{clientDirectory}\"\n is too long for {templateFile.FileName}.");
                                                    continue;
                                                }
                                            }
                                            clientDirectory = $"{sessionFolderName}/{outputSubFolder}";
                                            if (outputFileName.Contains("OriginalFileName"))
                                            {
                                                var file = templateFile.FileName.Substring(0, templateFile.FileName.LastIndexOf("."));
                                                outputFileName = outputFileName.Replace("OriginalFileName", file);
                                            }
                                            if (!hasFileNamePartId)
                                                outputFileName = outputFileName + counter;
                                            var filePath = clientDirectory + $@"\{outputFileName}.pdf";
                                            try
                                            {
                                                if (!File.Exists(filePath))
                                                {
                                                    File.Copy(pdfFilePath, filePath);
                                                    templateZip.AddFile(filePath, clientDirectory);
                                                }
                                            }
                                            catch (PathTooLongException)
                                            {
                                                response.Message.Add($"The file path:\n\"{filePath}\"\n is too long for {templateFile.FileName}.");
                                                continue;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            parentTableName = Helper.GetParentTableByFileId(templateFile.TemplateFileId);
                            DataTable dt = new DataTable();
                            foreach (DataTable ndt in ds.Tables)
                            {
                                if (ndt.TableName == parentTableName)
                                {
                                    isPdfMapped = true;
                                    dt = ndt;
                                    dt.TableName = ndt.TableName;
                                    break;
                                }
                            }
                            for (var k = 0; k < dt.Rows.Count; k++)
                            {
                                var outputSubFolder = subFolder;
                                var outputFileName = fileNamePart;
                                var hasFileNamePartId = false;
                                if (!string.IsNullOrWhiteSpace(fileNamePart) && fileNamePart.ToLower().Contains("id"))
                                    hasFileNamePartId = true;
                                counter++;

                                for (int p = 0; p < fileMappedFields.Rows.Count; p++)
                                {
                                    var isDynamic = Convert.ToBoolean(fileMappedFields.Rows[p]["IsDynamic"]);
                                    var fieldId = Convert.ToString(fileMappedFields.Rows[p]["FieldId"]);
                                    var parentFieldId = Convert.ToString(fileMappedFields.Rows[p]["ParentFieldId"]);
                                    var tableName = Convert.ToString(fileMappedFields.Rows[p]["ExcelTableName"]);
                                    var sheetName = Convert.ToString(fileMappedFields.Rows[p]["SheetName"]);
                                    var pdfFieldName = Convert.ToString(fileMappedFields.Rows[p]["PDFFieldName"]);
                                    var excelFieldName = Convert.ToString(fileMappedFields.Rows[p]["ExcelFieldName"]);
                                    var xpath = Convert.ToString(fileMappedFields.Rows[p]["XPath"]);

                                    for (var l = 0; l < dt.Columns.Count; l++)
                                    {
                                        excelColumnNameValue = Convert.ToString(dt.Rows[k][l]);
                                        var excelColumnName = dt.Columns[l].ColumnName;
                                        if (subFolder != null && subFolder.Contains(excelColumnName))
                                            outputSubFolder = outputSubFolder.Replace(excelColumnName, excelColumnNameValue);
                                        if (fileNamePart != null && fileNamePart.Contains(excelColumnName))
                                            outputFileName = outputFileName.Replace(excelColumnName, excelColumnNameValue);
                                        var column = excelColumnName.Split(new[] { "--" }, StringSplitOptions.None);
                                        if (excelFieldName == column[0])
                                        {
                                            var xmlNode = xmlDoc.SelectSingleNode(xpath + "[" + (column.Count() > 1 ? column[1] : "1")+"]");
                                            xmlNode.InnerText = excelColumnNameValue;
                                        }
                                    }
                                }

                                if (isPdfMapped)
                                {
                                    outputSubFolder = outputSubFolder.Replace("{", "").Replace("}", "").Trim();
                                    outputFileName = outputFileName.Replace("{", "").Replace("}", "").Trim();

                                    if (string.IsNullOrWhiteSpace(outputSubFolder))
                                    {
                                        response.Message.Add($"No output subfolder name data found in the table for {templateFile.FileName}.");
                                        continue;
                                    }
                                    if (string.IsNullOrWhiteSpace(outputFileName))
                                    {
                                        response.Message.Add($"No output file naming options data found in the table for {templateFile.FileName}.");
                                        continue;
                                    }
                                    if (!IsNameValid(outputSubFolder))
                                    {
                                        response.Message.Add($"Output subfolder name is not valid for {templateFile.FileName}.");
                                        continue;
                                    }
                                    if (!IsNameValid(outputFileName))
                                    {
                                        response.Message.Add($"Output file name is not valid for {templateFile.FileName}.");
                                        continue;
                                    }


                                    xmlDoc.Save(xmlFilePath);

                                    using (FileStream xmlStream = new FileStream(xmlFilePath, FileMode.Open, FileAccess.Read))
                                    {
                                        using (var pdfStream = File.Open(pdfFilePath, FileMode.Open, FileAccess.ReadWrite))
                                        {
                                            AsposePdf.Facades.Form form = new AsposePdf.Facades.Form(pdfStream);
                                            if (templateFile.IsXFA)
                                                form.SetXfaData(xmlStream);
                                            else
                                                form.ImportXml(xmlStream);
                                            form.Save(pdfStream);
                                            /*pdfDocument.Save(pdfFilePath);
                                            using (var fileStream = new FileStream(pdfFilePath, FileMode.Create, FileAccess.Write))
                                            {
                                                pdfOutStream.CopyTo(fileStream);
                                            }*/
                                        }
                                    }

                                    if (!Directory.Exists($"{sessionFolderName}/{outputSubFolder}"))
                                    {
                                        clientDirectory = $"{sessionFolderName}/{outputSubFolder}";
                                        try
                                        {
                                            Directory.CreateDirectory(clientDirectory);
                                        }
                                        catch (PathTooLongException)
                                        {
                                            response.Message.Add($"The subfolder path:\n\"{clientDirectory}\"\n is too long for {templateFile.FileName}.");
                                            continue;
                                        }
                                    }
                                    clientDirectory = $"{sessionFolderName}/{outputSubFolder}";
                                    if (outputFileName.Contains("OriginalFileName"))
                                    {
                                        var file = templateFile.FileName.Substring(0, templateFile.FileName.LastIndexOf("."));
                                        outputFileName = outputFileName.Replace("OriginalFileName", file);
                                    }
                                    if (!hasFileNamePartId)
                                        outputFileName = outputFileName + counter;
                                    var filePath = clientDirectory + $@"\{outputFileName}.pdf";
                                    try
                                    {
                                        if (!File.Exists(filePath))
                                        {
                                            File.Copy(pdfFilePath, filePath);
                                            templateZip.AddFile(filePath, clientDirectory);
                                        }
                                    }
                                    catch (PathTooLongException)
                                    {
                                        response.Message.Add($"The file path:\n\"{filePath}\"\n is too long for {templateFile.FileName}.");
                                        continue;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        parentTableName = Helper.GetParentTableByFileId(templateFile.TemplateFileId);
                        DataTable dt = new DataTable();
                        foreach (DataTable ndt in ds.Tables)
                        {
                            if (ndt.TableName == parentTableName)
                            {
                                isPdfMapped = true;
                                dt = ndt;
                                dt.TableName = ndt.TableName;
                                break;
                            }
                        }
                        for (var k = 0; k < dt.Rows.Count; k++)
                        {
                            var outputSubFolder = subFolder;
                            var outputFileName = fileNamePart;
                            var hasFileNamePartId = false;
                            if (!string.IsNullOrWhiteSpace(fileNamePart) && fileNamePart.ToLower().Contains("id"))
                                hasFileNamePartId = true;
                            counter++;

                            for (int p = 0; p < fileMappedFields.Rows.Count; p++)
                            {
                                var isDynamic = Convert.ToBoolean(fileMappedFields.Rows[p]["IsDynamic"]);
                                var fieldId = Convert.ToString(fileMappedFields.Rows[p]["FieldId"]);
                                var parentFieldId = Convert.ToString(fileMappedFields.Rows[p]["ParentFieldId"]);
                                var tableName = Convert.ToString(fileMappedFields.Rows[p]["ExcelTableName"]);
                                var sheetName = Convert.ToString(fileMappedFields.Rows[p]["SheetName"]);
                                var pdfFieldName = Convert.ToString(fileMappedFields.Rows[p]["PDFFieldName"]);
                                var excelFieldName = Convert.ToString(fileMappedFields.Rows[p]["ExcelFieldName"]);
                                var xpath = Convert.ToString(fileMappedFields.Rows[p]["XPath"]);

                                for (var l = 0; l < dt.Columns.Count; l++)
                                {
                                    excelColumnNameValue = Convert.ToString(dt.Rows[k][l]);
                                    var excelColumnName = dt.Columns[l].ColumnName;

                                    if (subFolder != null && subFolder.Contains(excelColumnName))
                                        outputSubFolder = outputSubFolder.Replace(excelColumnName, excelColumnNameValue);
                                    if (fileNamePart != null && fileNamePart.Contains(excelColumnName))
                                        outputFileName = outputFileName.Replace(excelColumnName, excelColumnNameValue);
                                    if (excelFieldName == excelColumnName)
                                    {
                                        var xmlNode = xmlDoc.SelectSingleNode(xpath);

                                        XmlElement elem = xmlDoc.CreateElement("value");
                                        elem.InnerText = excelColumnNameValue;
                                        if (xmlNode.HasChildNodes)
                                            xmlNode.RemoveChild(xmlNode.ChildNodes[0]);
                                        xmlNode.AppendChild(elem);
                                    }
                                }
                            }

                            if (isPdfMapped)
                            {
                                outputSubFolder = outputSubFolder.Replace("{", "").Replace("}", "").Trim();
                                outputFileName = outputFileName.Replace("{", "").Replace("}", "").Trim();

                                if (string.IsNullOrWhiteSpace(outputSubFolder))
                                {
                                    response.Message.Add($"No output subfolder name data found in the table for {templateFile.FileName}.");
                                    continue;
                                }
                                if (string.IsNullOrWhiteSpace(outputFileName))
                                {
                                    response.Message.Add($"No output file naming options data found in the table for {templateFile.FileName}.");
                                    continue;
                                }
                                if (!IsNameValid(outputSubFolder))
                                {
                                    response.Message.Add($"Output subfolder name is not valid for {templateFile.FileName}.");
                                    continue;
                                }
                                if (!IsNameValid(outputFileName))
                                {
                                    response.Message.Add($"Output file name is not valid for {templateFile.FileName}.");
                                    continue;
                                }

                                xmlDoc.Save(xmlFilePath);

                                using (FileStream xmlStream = new FileStream(xmlFilePath, FileMode.Open, FileAccess.Read))
                                {
                                    using (var pdfStream = File.Open(pdfFilePath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                                    {
                                        AsposePdf.Facades.Form form = new AsposePdf.Facades.Form(pdfStream);
                                        if (templateFile.IsXFA)
                                            form.SetXfaData(xmlStream);
                                        else
                                            form.ImportXml(xmlStream);
                                        form.Save(pdfStream);
                                    }
                                }

                                if (!Directory.Exists($"{sessionFolderName}/{outputSubFolder}"))
                                {
                                    clientDirectory = $"{sessionFolderName}/{outputSubFolder}";
                                    try
                                    {
                                        Directory.CreateDirectory(clientDirectory);
                                    }
                                    catch (PathTooLongException)
                                    {
                                        response.Message.Add($"The subfolder path:\n\"{clientDirectory}\"\n is too long for {templateFile.FileName}.");
                                        continue;
                                    }
                                }
                                clientDirectory = $"{sessionFolderName}/{outputSubFolder}";
                                if (outputFileName.Contains("OriginalFileName"))
                                {
                                    var file = templateFile.FileName.Substring(0, templateFile.FileName.LastIndexOf("."));
                                    outputFileName = outputFileName.Replace("OriginalFileName", file);
                                }
                                if (!hasFileNamePartId)
                                    outputFileName = outputFileName + counter;
                                var filePath = clientDirectory + $@"\{outputFileName}.pdf";
                                try
                                {
                                    if (!File.Exists(filePath))
                                    {
                                        File.Copy(pdfFilePath, filePath);
                                        templateZip.AddFile(filePath, clientDirectory);
                                    }
                                }
                                catch (PathTooLongException)
                                {
                                    response.Message.Add($"The file path:\n\"{filePath}\"\n is too long for {templateFile.FileName}.");
                                    continue;
                                }
                            }
                        }
                    }
                    if (File.Exists(pdfFilePath))
                        File.Delete(pdfFilePath);
                }

                response.ZipPath = $"/Content/ExFormsData-{Guid.NewGuid()}.zip";
                var zipFilePath = HttpContext.Current.Request.MapPath("~" + response.ZipPath);
                templateZip.Save(zipFilePath);

                if (Directory.Exists(sessionFolderName))
                {
                    Directory.Delete(sessionFolderName, true);
                }
            }
            catch (Exception ex)
            {
                response.Error = ex.Message;
            }
            return response;
        }

        [Route("api/Template/DeleteZipFolder")]
        [HttpPost]
        public string DeleteZipFolder([FromBody] string zipPath)
        {
            try
            {
                zipPath = HttpContext.Current.Request.MapPath("~" + zipPath);
                if (File.Exists(zipPath))
                {
                    File.Delete(zipPath);
                }
                return "success";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        [Route("api/Template/DeleteExcelTemplateFile")]
        [HttpPost]
        public string DeleteExcelTemplateFile([FromBody] string filePath)
        {
            try
            {
                filePath = HttpContext.Current.Request.MapPath("~" + filePath);
                var folderPath = filePath.Substring(0, filePath.LastIndexOf('\\'));
                if (Directory.Exists(folderPath))
                {
                    Directory.Delete(folderPath, true);
                }
                return "success";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        private bool IsNameValid(string name)
        {
            foreach (char ch in Path.GetInvalidFileNameChars())
            {
                if (name.Contains(ch))
                    return false;
            }
            return true;
        }

        private void ClearFieldsValue(XmlNode xmlNode)
        {
            for (var i = 0; i <= xmlNode.ChildNodes.Count - 1; i++)
            {
                XmlNode newNode = xmlNode.ChildNodes[i];
                //if (newNode.NextSibling != null && newNode.NextSibling.Name == newNode.Name)
                //{
                //    xmlNode.RemoveChild(newNode.NextSibling);
                //    i--;
                //    continue;
                //}
                if (newNode.NodeType == XmlNodeType.Text)
                    newNode.InnerText = string.Empty;
                else if (newNode.NodeType != XmlNodeType.Text)
                    ClearFieldsValue(newNode);
            }
        }

        [Route("api/Template/GetRangeParameter")]
        [HttpGet]
        public RangeParamResponse GetRangeParameter(string fieldId)
        {
            var rangeParamResponse = new RangeParamResponse();
            try
            {
                var mappedField = Helper.GetMappedFieldParameterByFieldId(fieldId);
                if (mappedField != null && mappedField.Rows.Count == 1)
                {
                    rangeParamResponse.ColumnName = Convert.ToString(mappedField.Rows[0]["ExcelFieldName"]);
                    rangeParamResponse.TableName = Convert.ToString(mappedField.Rows[0]["ExcelTableName"]);
                    rangeParamResponse.SheetName = Convert.ToString(mappedField.Rows[0]["SheetName"]);
                }
            }
            catch (Exception ex)
            {
                return null;
            }
            return rangeParamResponse;
        }

        [Route("api/Template/GetTableFields")]
        [HttpGet]
        public ParentChildFields GetTableFields(int fileId, string parentTable, string childTable)
        {
            var parentChildFields = new ParentChildFields();
            parentChildFields.ChildFields = new List<string>();
            parentChildFields.ParentFields = new List<string>();

            try
            {
                var childFields = Helper.GetTableFields(fileId, childTable);
                var parentFields = Helper.GetTableFields(fileId, parentTable);
                var parenChildColumns = Helper.GetExistingParentChildMapping(fileId, parentTable, childTable);
                if (parenChildColumns.Rows.Count == 1)
                {
                    parentChildFields.ChildField = Convert.ToString(parenChildColumns.Rows[0]["ChildTableColumn"]);
                    parentChildFields.ParentField = Convert.ToString(parenChildColumns.Rows[0]["ParentTableColumn"]);
                }

                if (childFields.Rows.Count > 0)
                {
                    var rows = childFields.Rows;
                    var rowCount = childFields.Rows.Count;
                    for (var i = 0; i < rowCount; i++)
                    {
                        var filedName = Convert.ToString(rows[i]["ExcelFieldName"]);
                        parentChildFields.ChildFields.Add(filedName);
                    }
                }
                if (parentFields.Rows.Count > 0)
                {
                    var rows = parentFields.Rows;
                    var rowCount = parentFields.Rows.Count;
                    for (var i = 0; i < rowCount; i++)
                    {
                        var filedName = Convert.ToString(rows[i]["ExcelFieldName"]);
                        parentChildFields.ParentFields.Add(filedName);
                    }
                }
            }
            catch (Exception ex)
            {
                return null;
            }
            return parentChildFields;
        }

        [Route("api/Template/GetParentTableFields")]
        [HttpGet]
        public List<string> GetParentTableFields(int templateId)
        {
            var response = new List<string>();

            try
            {
                var parentTableFields = Helper.GetParentTableFields(templateId);
                if (parentTableFields.Rows.Count > 0)
                {
                    var rows = parentTableFields.Rows;
                    var rowCount = parentTableFields.Rows.Count;
                    for (var i = 0; i < rowCount; i++)
                    {
                        var filedName = Convert.ToString(rows[i]["ExcelFieldName"]);
                        response.Add(filedName);
                    }
                }
            }
            catch (Exception ex)
            {
                return null;
            }
            return response;
        }

        [Route("api/Template/UpdateFileNamingOption")]
        [HttpPost()]
        public string UpdateFileNamingOption(NamingOption fileNamingOption)
        {
            try
            {
                Helper.UpdateFileNamingOption(fileNamingOption);
                return "success";
            }
            catch (Exception ex)
            {
                return null;
            }

        }

        [Route("api/Template/UpdateFolderNamingOption")]
        [HttpPost()]
        public string UpdateFolderNamingOption(NamingOption folderNamingOption)
        {
            try
            {
                Helper.UpdateFolderNamingOption(folderNamingOption);
                return "success";
            }
            catch (Exception ex)
            {
                return null;
            }

        }

        [Route("api/Template/SaveParentChildMapping")]
        [HttpPost()]
        public string SaveParentChildMapping(ParentChildMapping parentChildMapping)
        {
            try
            {
                Helper.AddParentChildTableMapping(parentChildMapping.TemplateFileId, parentChildMapping.ParentTable, parentChildMapping.ParentTableField, parentChildMapping.ChildTable, parentChildMapping.ChildTableField);
                return "success";
            }
            catch (Exception ex)
            {
                return null;
            }

        }

        [Route("api/Template/UpdateExcelColumnMapping")]
        [HttpPost()]
        public string UpdateExcelColumnMapping(ColumnValue column)
        {
            try
            {
                Helper.UpdateExcelColumnMapping(column.TemplateId, column.OldValue, column.NewValue);
                return "success";
            }
            catch (Exception ex)
            {
                return null;
            }

        }

        [Route("api/Template/GetExistingParentChildMapping")]
        [HttpGet]
        public DataTable GetExistingParentChildMapping(int templateFileId, string parentTable, string childTable)
        {
            var parenChildColumns = new DataTable();

            try
            {
                parenChildColumns = Helper.GetExistingParentChildMapping(templateFileId, parentTable, childTable);
            }
            catch (Exception ex)
            {
                return null;
            }
            return parenChildColumns;
        }

        [Route("api/Template/GetExcelVersion")]
        [HttpGet]
        public ExcelVersionResponse GetExcelVersion(int templateId)
        {
            ExcelVersionResponse res = new ExcelVersionResponse();
            try
            {
                var dt = Helper.GetExcelVersion(templateId);
                if (dt.Rows.Count == 1)
                {
                    res.ExcelVersion = Convert.ToString(dt.Rows[0]["ExcelVersion"]);
                    res.UpdatedOn = Convert.ToDateTime(dt.Rows[0]["UpdatedOn"]).ToString("MMMM dd yyyy");
                    res.UpdatedBy = Convert.ToString(dt.Rows[0]["UpdatedBy"]);
                }
            }
            catch (Exception ex)
            {
                res.Error = ex.Message;
            }
            return res;
        }

        [Route("api/Template/GetTableParentChildTables")]
        [HttpGet]
        public DataTable GetTableParentChildTables(int fileId)
        {
            var parentChildTables = new DataTable();
            try
            {
                var parentTable = Helper.GetParentTableByFileId(fileId);
                var childTables = Helper.GetChildTablesByFileId(fileId);
                CreateParentChildTableDt(parentChildTables);

                for (int i = 0; i < childTables.Rows.Count; i++)
                {
                    var dr = parentChildTables.NewRow();
                    dr["ParentTable"] = parentTable;
                    dr["ChildTable"] = childTables.Rows[i]["ExcelTableName"];
                    dr["IsMapped"] = childTables.Rows[i]["IsMapped"];
                    parentChildTables.Rows.Add(dr);
                }
            }
            catch (Exception ex)
            {
                return null;
            }
            return parentChildTables;
        }

        [Route("api/Template/DownloadExcelTemplate")]
        [HttpGet]
        public string DownloadExcelTemplate(int templateId, int userId)
        {
            try
            {
                var excelFile = Helper.GetTemplateExcelBytes(templateId);

                var excelStream = new MemoryStream(excelFile.ExcelBytes);
                var folderPath = HttpContext.Current.Request.MapPath("~/Content/" + userId);
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }
                var path = $"/Content/{userId}/Exf_Template-{excelFile.TemplateName}{excelFile.FileExtension}";
                var excelFilePath = HttpContext.Current.Request.MapPath("~" + path);
                if (File.Exists(excelFilePath))
                    File.Delete(excelFilePath);
                using (FileStream excelFileStream = new FileStream(excelFilePath, FileMode.Create, FileAccess.ReadWrite))
                {
                    excelStream.Seek(0, SeekOrigin.Begin);
                    excelStream.CopyTo(excelFileStream);
                }

                //HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK) { Content = new StreamContent(new MemoryStream(stream)) };
                //result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
                //result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                //result.Content.Headers.ContentDisposition.FileName = "ExcelTemplate.xlsx";
                //return result;
                return path;
            }
            catch (Exception)
            {
                return null;
                //return Request.CreateErrorResponse(HttpStatusCode.NotFound, "File Not Found");
            }
        }

        private DataTable GetTemplateFields(int templateId)
        {
            return Helper.GetTemplateFieldsByTemplateId(templateId);
        }
        public class JsTreeAttribute
        {
            public int templateFileMappingId;
            public string id;
            public string text;
            public object children;
            public string icon;
            public string type;
            public string title;
            public bool isMapped;
        }
        public class RemoveMappingsResponse
        {
            public bool IsAnyFieldMapped;
            public string Error;
        }

        [HttpGet]
        public TemplateInfo GetTemplateById(int id)
        {
            try
            {
                var templateInfo = Helper.GetTemplateById(id);
                return templateInfo;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        [HttpGet]
        public string GetTemplateFileFieldsMappedPercentage(int templateFileId)
        {
            try
            {
                var mappedPercentage = Helper.GetTemplateFileFieldsMappedPercentage(templateFileId);
                return mappedPercentage;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        private void CreateParentChildTableDt(DataTable dt)
        {
            dt.Columns.Add("ParentTable", typeof(string));
            dt.Columns.Add("ChildTable", typeof(string));
            dt.Columns.Add("IsMapped", typeof(bool));
        }

        private void CreateTemplateFileDt(DataTable dtTemplateFile)
        {
            dtTemplateFile.Columns.Add("TemplateFileId", typeof(int));
            dtTemplateFile.Columns.Add("TemplateId", typeof(int));
            dtTemplateFile.Columns.Add("FileName", typeof(string));
            dtTemplateFile.Columns.Add("FileDisplayName", typeof(string));
            dtTemplateFile.Columns.Add("FilePath", typeof(string));
            dtTemplateFile.Columns.Add("IsXFA", typeof(bool));
            dtTemplateFile.Columns.Add("PdfBytes", typeof(byte[]));
        }

        private void CreateEditTemplateDt(DataTable dtEditTemplate)
        {
            dtEditTemplate.Columns.Add("TemplateId", typeof(int));
            dtEditTemplate.Columns.Add("CompanyId", typeof(int));
            dtEditTemplate.Columns.Add("TemplateName", typeof(string));
            dtEditTemplate.Columns.Add("Description", typeof(string));
            dtEditTemplate.Columns.Add("TemplateFileZip", typeof(byte[]));
            dtEditTemplate.Columns.Add("IsActive", typeof(bool));
            dtEditTemplate.Columns.Add("CreatedOn", typeof(DateTime));
            dtEditTemplate.Columns.Add("CreatedBy", typeof(int));
            dtEditTemplate.Columns.Add("UpdatedOn", typeof(DateTime));
            dtEditTemplate.Columns.Add("UpdatedBy", typeof(int));
            dtEditTemplate.Columns.Add("SubFolderName", typeof(string));
            dtEditTemplate.Columns.Add("FileNamePart", typeof(string));

        }

        private void CreateTemplateFieldsDt(DataTable dtTemplateField)
        {
            dtTemplateField.Columns.Add("TemplateFileMappingId", typeof(int));
            dtTemplateField.Columns.Add("TemplateId", typeof(int));
            dtTemplateField.Columns.Add("TemplateFileId", typeof(int));
            dtTemplateField.Columns.Add("PDFFieldName", typeof(string));
            dtTemplateField.Columns.Add("IsMapped", typeof(bool));
            dtTemplateField.Columns.Add("FilePath", typeof(string));
            dtTemplateField.Columns.Add("SheetName", typeof(string));
            dtTemplateField.Columns.Add("ExcelTableName", typeof(string));
            dtTemplateField.Columns.Add("FieldId", typeof(string));
            dtTemplateField.Columns.Add("ParentFieldId", typeof(string));
            dtTemplateField.Columns.Add("IsDynamic", typeof(bool));
            dtTemplateField.Columns.Add("HasChildFields", typeof(bool));
            dtTemplateField.Columns.Add("XPath", typeof(string));
        }

        private void CreateMappedPdfTable(DataTable dtMappedFields)
        {
            dtMappedFields.Columns.Add("TemplateFileMappingId", typeof(int));
            dtMappedFields.Columns.Add("TemplateId", typeof(int));
            dtMappedFields.Columns.Add("TemplateFileId", typeof(int));
            dtMappedFields.Columns.Add("PDFFieldName", typeof(string));
            dtMappedFields.Columns.Add("IsMapped", typeof(bool));
            dtMappedFields.Columns.Add("FilePath", typeof(string));
            dtMappedFields.Columns.Add("SheetName", typeof(string));
            dtMappedFields.Columns.Add("ExcelTableName", typeof(string));
            dtMappedFields.Columns.Add("FieldId", typeof(string));
            dtMappedFields.Columns.Add("ParentFieldId", typeof(string));
            dtMappedFields.Columns.Add("IsDynamic", typeof(bool));
            dtMappedFields.Columns.Add("HasChildFields", typeof(bool));
            dtMappedFields.Columns.Add("XPath", typeof(string));
        }
        private void CreatePdfFileDt(DataTable dtPdfFile)
        {
            dtPdfFile.Columns.Add("TemplateFileId", typeof(int));
            dtPdfFile.Columns.Add("File Name", typeof(string));
            dtPdfFile.Columns.Add("Mapped", typeof(string));
            dtPdfFile.Columns.Add("FilePath", typeof(string));
            dtPdfFile.Columns.Add("IsXFA", typeof(bool));
        }
        private void CreateTemplateFileTable(ref DataTable dtPdfTable)
        {
            dtPdfTable.Columns.Add("TemplateId", typeof(int));
            dtPdfTable.Columns.Add("TemplateFileId", typeof(int));
            dtPdfTable.Columns.Add("FileName", typeof(string));
            dtPdfTable.Columns.Add("FileDisplayName", typeof(string));
            dtPdfTable.Columns.Add("FilePath", typeof(string));
            dtPdfTable.Columns.Add("IsXFA", typeof(bool));
            dtPdfTable.Columns.Add("PdfBytes", typeof(byte[]));
        }
        private void CreateParentFieldsDt(DataTable dtParentField)
        {
            dtParentField.Columns.Add("TemplateFileMappingId", typeof(int));
            dtParentField.Columns.Add("ExcelFieldName", typeof(string));
            dtParentField.Columns.Add("SheetName", typeof(string));
            dtParentField.Columns.Add("IsMapped", typeof(bool));
            dtParentField.Columns.Add("ExcelTableName", typeof(string));
            dtParentField.Columns.Add("XPath", typeof(string));
        }
    }
}
