using ExFormOfficeAddInEntities;
using Ionic.Zip;
using System;
using System.Collections.Generic;
using System.IO;
using System.Web;
using System.Web.Http;

namespace ExFormOfficeAddInExcelUIWeb.Controllers
{
    public class UploadFileController : ApiController
    {
        [Route("api/UploadFile/UploadFiles")]
        [HttpPost()]
        public void UploadFiles()
        {
            try
            {
                var session = HttpContext.Current.Session;
                if (session != null)
                {
                    if (session["UloadedFiles"] == null)
                    {
                        session["UloadedFiles"] = new List<HttpPostedFile>();
                    }
                }

                var httpContext = HttpContext.Current;
                
                // Check for any uploaded file  
                if (httpContext.Request.Files.Count > 0)
                {
                    var uploadedFiles = session["UloadedFiles"] as List<HttpPostedFile>;
                    //Loop through uploaded files  
                    for (int i = 0; i < httpContext.Request.Files.Count; i++)
                    {
                        var isFileUploaded = false;
                        HttpPostedFile httpPostedFile = httpContext.Request.Files[i];

                        if (httpPostedFile != null)
                        {                            
                            if (uploadedFiles.Count > 0)
                            {
                                foreach (var file in uploadedFiles)
                                {
                                    if (file.FileName == httpPostedFile.FileName)
                                    {
                                        isFileUploaded = true;
                                        break;
                                    }
                                }
                            }
                            else
                                isFileUploaded = true;
                            if (isFileUploaded)
                                uploadedFiles.Add(httpPostedFile);                            
                            // Construct file save path  
                            //var fileSavePath = Path.Combine(HostingEnvironment.MapPath(ConfigurationManager.AppSettings["fileUploadFolder"]), httpPostedFile.FileName);

                            //// Save the uploaded file  
                            //httpPostedFile.SaveAs(fileSavePath);
                        }
                        session["UloadedFiles"] = uploadedFiles;
                    }
                }

                //HttpPostedFile file = HttpContext.Current.Request.Files["file"];
                ///*if (HttpContext.Current.Session["files"] == null)
                //    HttpContext.Current.Session["files"] = new List<HttpPostedFile>();
                //else
                //{
                //    var files = HttpContext.Current.Session["files"] as List<HttpPostedFile>;
                //    files.Add(file);
                //}*/
                //byte[] fileData = null;
                //var templateMemoryStream = new MemoryStream();
                //ZipFile templateZip = new ZipFile();
                //using (var binaryReader = new BinaryReader(file.InputStream))
                //{
                //    fileData = binaryReader.ReadBytes(file.ContentLength);

                //    templateZip.AddFile(file.FileName, "Temlate");
                //    templateZip.Save(templateMemoryStream);
                //}
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /*private void GetNewTemplateStream(out MemoryStream templateMemoryStream)
        {
            templateMemoryStream = new MemoryStream();

            using (ZipFile templateZip = new ZipFile())
            {
                foreach (var pdf in _lstNewSetPdfFiles)
                {
                    if (File.Exists(pdf.Folder))
                        templateZip.AddFile(pdf.Folder, "Temlate");

                    var dataRow = _dtSelectedPdfFiles.NewRow();
                    dataRow["FileName"] = pdf.FileName;
                    dataRow["FilePath"] = pdf.Folder;
                    dataRow["IsXFA"] = pdf.IsXFA;
                    _dtSelectedPdfFiles.Rows.Add(dataRow);
                }

                templateZip.Save(templateMemoryStream);
            }
        }*/
    }
}
