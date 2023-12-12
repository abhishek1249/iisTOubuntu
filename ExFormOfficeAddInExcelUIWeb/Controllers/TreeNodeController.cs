using ExFormOfficeAddInBAL;
using ExFormOfficeAddInEntities;
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
            public object children;
            public string icon;
            public string type;
            public string title;
            public bool IsDemo;
        }

        /*[HttpPost]
        public JsonResult GetTreeData()
        {
            //var list = new List();

            //var lstJsTree = BuildMYTree(list);
            //return Json(lstJsTree);

            JsTreeModel rootNode = new JsTreeModel();
            rootNode.attr = new JsTreeAttribute();
            rootNode.data = "Root";
            string rootPath = Request.MapPath(dataPath);
            rootNode.attr.id = rootPath;
            PopulateTree(rootPath, rootNode);
            AlreadyPopulated = true;
            return Json(rootNode);

        }*/

        [Route("api/TreeNode/UpsertTemplateFolder")]
        [HttpPost()]
        public object UpsertTemplateFolder(UpsertTemplateFolderParam upsertTemplateFolder)
        {
            try
            {
                if(upsertTemplateFolder.ParentFolderId=="#")
                {
                    Helper.CreateTemplateFolder(Convert.ToInt32(upsertTemplateFolder.CompanyId), null, upsertTemplateFolder.FolderName);
                }
                else if (upsertTemplateFolder.Type == "Folder")
                {
                    var parentFolderId = upsertTemplateFolder.ParentFolderId.Replace("F", "");
                    var folderId = upsertTemplateFolder.FolderId.Replace("F", "");
                    if (upsertTemplateFolder.Old == "New node")
                        Helper.CreateTemplateFolder(Convert.ToInt32(upsertTemplateFolder.CompanyId), Convert.ToInt32(parentFolderId), upsertTemplateFolder.FolderName);
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
            catch
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
                Helper.CreateTemplateCopy(Convert.ToInt32(Id), Convert.ToInt32(duplicateTemplate.CompanyId), Convert.ToInt32(duplicateTemplate.UserId));
            }
            catch
            {
                return Json("false");
            }

            return Json(duplicateTemplate.Id);
        }

        [HttpGet()]
        public object GetFolders(int? id, string companyId)
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
            
            var templateFolders = Helper.GetTemplateFolderByCompanyId(Convert.ToInt32(companyId));
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
        }
    }
}
