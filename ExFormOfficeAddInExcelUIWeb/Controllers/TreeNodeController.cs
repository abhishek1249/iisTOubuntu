﻿using ExFormOfficeAddInBAL;
using ExFormOfficeAddInEntities;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Web.Http;

namespace ExFormOfficeAddInExcelUIWeb.Controllers
{
    public class TreeNodeController : ApiController
    {
        public class JsTreeModel
        {
            public string data;
            public JsTreeAttribute attr;
            // this was "open" but changing it to “leaf” adds “jstree-leaf” to the class   
            public string state = "leaf";
            public List<JsTreeModel> children;
        }

        public class JsTreeAttribute
        {
            public string id;
            public string text;
            public string parent;
            //public object children;
            public string icon;
            public string type;
            public string title;
            public bool IsDemo;
        }

        
        [Route("api/TreeNode/UpsertTemplateFolder")]
        [HttpPost()]
        public object UpsertTemplateFolder(UpsertTemplateFolderParam upsertTemplateFolder)
        {
            try
            {
                if (upsertTemplateFolder.ParentFolderId == "#")
                {
                    if (upsertTemplateFolder.FolderId.Contains("j1_1"))
                    {
                        Helper.CreateTemplateFolder(Convert.ToInt32(upsertTemplateFolder.CompanyId), Convert.ToInt32(upsertTemplateFolder.TeamId), null, upsertTemplateFolder.FolderName);
                    }
                    else
                    {
                        var folderId = upsertTemplateFolder.FolderId.Replace("F", "");
                        Helper.RenameFolder(Convert.ToInt32(folderId), upsertTemplateFolder.FolderName);
                    }
                }
                else if (upsertTemplateFolder.Type == "Folder")
                {
                    var parentFolderId = upsertTemplateFolder.ParentFolderId.Replace("F", "");
                    var folderId = upsertTemplateFolder.FolderId.Replace("F", "");
                    if (upsertTemplateFolder.Old == "New node")
                        Helper.CreateTemplateFolder(Convert.ToInt32(upsertTemplateFolder.CompanyId), Convert.ToInt32(upsertTemplateFolder.TeamId), Convert.ToInt32(parentFolderId), upsertTemplateFolder.FolderName);
                    else
                        Helper.RenameFolder(Convert.ToInt32(folderId), upsertTemplateFolder.FolderName);
                }
                else
                {
                    var folderId = upsertTemplateFolder.FolderId.Replace("T", "");
                    Helper.RenameTemplate(Convert.ToInt32(folderId), upsertTemplateFolder.FolderName);
                }

            }
            catch
            {
                return Json("false");
            }

            return Json("true");
        }

        [Route("api/TreeNode/Delete")]
        [HttpPost()]
        public object Delete(TemplateFolderDelete templateFolderDeleteParam)
        {
            try
            {
                if (templateFolderDeleteParam.Type == "Folder")
                {
                    var id = templateFolderDeleteParam.Id.Replace("F", "");
                    Helper.DeleteFolder(Convert.ToInt32(id));
                }
                else
                {
                    var id = templateFolderDeleteParam.Id.Replace("T", "");
                    Helper.DeleteTemplate(Convert.ToInt32(id));
                }
            }
            catch (Exception ex)
            {
                return Json("false");
            }

            return Json("true");
        }

        [Route("api/TreeNode/CreateDuplicateTemplate")]
        [HttpPost()]
        public object CreateDuplicateTemplate(DuplicateTemplateParam duplicateTemplate)
        {
            try
            {
                var Id = duplicateTemplate.Id.Replace("T", "");
                Helper.CreateTemplateCopy(Convert.ToInt32(Id), Convert.ToInt32(duplicateTemplate.CompanyId), Convert.ToInt32(duplicateTemplate.TeamId), Convert.ToInt32(duplicateTemplate.UserId));
            }
            catch
            {
                return Json("false");
            }

            return Json(duplicateTemplate.Id);
        }

        /*[HttpGet()]
        public object GetFolders(int? id, string companyId, string teamId)
        {
            List<JsTreeAttribute> core = new List<JsTreeAttribute>();

            if(id==null)
            {
                var demoFolder = Helper.GetDemoFolder();
                JsTreeAttribute obj = new JsTreeAttribute()
                {
                    id = "F" + demoFolder.FolderId,
                    text = demoFolder.FolderName,
                    children = true,
                    type = "Folder",
                    IsDemo=true,
                    title = $"For Demo Purpose Only."
                };
                core.Add(obj);                
            }
            
            var templateFolders = Helper.GetTemplateFolderByCompanyId(Convert.ToInt32(companyId), Convert.ToInt32(teamId));
            foreach (var templateFolder in templateFolders)
            {
                //Without Child
                if (templateFolder.ParentFolderId == "" && id == null)
                {
                    JsTreeAttribute obj = new JsTreeAttribute()
                    {
                        id = "F" + templateFolder.FolderId,
                        text = templateFolder.FolderName,
                        children = true,
                        type = "Folder"
                    };
                    core.Add(obj);
                }
                //With child
                else if (id != null)
                {
                    if (!string.IsNullOrEmpty(templateFolder.ParentFolderId))
                    {
                        if (Convert.ToInt32(templateFolder.ParentFolderId) == id)
                        {
                            JsTreeAttribute obj = new JsTreeAttribute()
                            {
                                id = "F" + templateFolder.FolderId,
                                text = templateFolder.FolderName,
                                children = true,
                                type = "Folder"
                            };
                            core.Add(obj);
                        }
                    }
                }
            }

            if (id != null)
            {
                var FileInFolder = Helper.GetTemplateByFolderId(id);
                List<JsTreeAttribute> FilesList = new List<JsTreeAttribute>();
                foreach (var File in FileInFolder)
                {
                    JsTreeAttribute Fileobj = new JsTreeAttribute()
                    {
                        id = "T" + Convert.ToString(File.TemplateId),
                        text = File.TemplateName,
                        icon = "jstree-file",
                        children = false,
                        type = "File",
                        IsDemo= File.IsDemoTemplate,
                        title = $"{File.Description} ( {File.PdfCount} PDF Forms )"
                    };
                    core.Add(Fileobj);
                }
            }
            
            return Json<List<JsTreeAttribute>>(core);
        }*/

        [HttpGet()]
        public object GetFolders(int? id, string companyId, string teamId)
        {
            List<JsTreeAttribute> core = new List<JsTreeAttribute>();

            if (id == null)
            {
                var demoFolder = Helper.GetDemoFolder();

                JsTreeAttribute obj = new JsTreeAttribute()
                {
                    id = "F" + demoFolder.FolderId,
                    text = demoFolder.FolderName,
                    //children = true,
                    parent = "#",
                    type = "Folder",
                    IsDemo = true,
                    title = $"For Demo Purpose Only."
                };

                core.Add(obj);
                core = GetTemplateFiles(Convert.ToInt32(demoFolder.FolderId), core);

            }

            var templateFolders = Helper.GetTemplateFolderByCompanyId(Convert.ToInt32(companyId), Convert.ToInt32(teamId));
            foreach (var templateFolder in templateFolders)
            {

                JsTreeAttribute obj = new JsTreeAttribute()
                {
                    id = "F" + templateFolder.FolderId,
                    text = templateFolder.FolderName,
                    //children = true,
                    parent = String.IsNullOrEmpty(templateFolder.ParentFolderId) ? "#" : "F" + templateFolder.ParentFolderId,
                    type = "Folder"
                };
                core.Add(obj);
                core = GetTemplateFiles(Convert.ToInt32(templateFolder.FolderId), core);
            }

            return Json<List<JsTreeAttribute>>(core);
        }

        public List<JsTreeAttribute> GetTemplateFiles(int FolderId, List<JsTreeAttribute> FilesList)
        {
            var FileInFolder = Helper.GetTemplateByFolderId(FolderId);
            //List<JsTreeAttribute> FilesList = new List<JsTreeAttribute>();
            foreach (var File in FileInFolder)
            {
                JsTreeAttribute Fileobj = new JsTreeAttribute()
                {
                    id = "T" + Convert.ToString(File.TemplateId),
                    text = File.TemplateName,
                    icon = "jstree-file",
                    //children = false,
                    parent = "F" + FolderId,
                    type = "File",
                    IsDemo = File.IsDemoTemplate,
                    title = $"{File.Description} ( {File.PdfCount} PDF Forms )"
                };
                FilesList.Add(Fileobj);
            }
            return FilesList;
        }
    }
}
