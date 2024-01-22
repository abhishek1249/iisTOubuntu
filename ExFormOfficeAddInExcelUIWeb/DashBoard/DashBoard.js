var pdffiles = [];
var selectedFiles = [];
var deletedFiles = [];
var template;
var editFileId;
var isExcelVersionUpdated = false;
var isFileMappingEdited = false;
var isXfa = false;
var parentTable;
var childTable;
var isParentChildRelationshipSaved = false;
var couterCheck = 0;
var templateId = -1;
var currentTemplateFileId = -1;
var arrFileMap = [];
var currentRemoveFileId = -1;
var isSubFolderNameOption = false;
var isFileNameOption = false;
var docdata = [];
var isExcelVersioMatching = false;
var isDemoSet = false;
var excludeFiles = '';
var rootNodeText = '';
var excelVersion = '';
var isNamingOptionsModified = false;
var teamId = "0";

(function () {
    "use strict";
    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function () {
        $(document).ready(function () {
            app.initialize();
            // Initialize the FabricUI notification mechanism and hide it
            var to = false;
            $('.ms-NavBar').NavBar();
            $("#logout").show();
            $("#loggingout").hide();
            $("#accordion").accordion();

            if (IsAdmin() || IsSuperAdmin()) {
                $("#divEditSet").show();
            }
            else {
                $("#divEditSet").hide();
            }
            $("#Logout-buttonn").click(LogoutBtn);

            $('#newSet').hide();
            $('#defalts').hide();
            $('#selectedSet').hide();
            $('#templateSets').hide();
            $('#editSet').hide();
            $('#editFieldsMapping').hide();
            $('#defaults').hide();
            $('#ParentChildTable').hide();
            $('#ParentChildTableRelationship').hide();

            $('#txtDefaultsSubFolderName').attr('readonly', 'true');
            $('#txtDefaultsFileName').attr('readonly', 'true');


            $("#btnSave").click(CreateTemplate);
            //$("#btnUpdateVersion").click(UpdateExcelVersion);
            $("#btnEditSet").click(EditSet);
            $("#btnTableRelationship").click(GetTableParentChildTables);
            $("#btnSaveExistingSet").click(SaveExistingSet);
            $("#btnNamingOptions").click(EditDefaults);
            $("#btnRemoveMappings").click(RemoveAllMappings);
            $("#btnRemoveFileMapping").click(confirmRemoveFileMappings);
            $("#btnAutomapFields").click(confirmAutoFieldsMappings);
            $("#btnSaveParentChildTableRelationship").click(SaveParentChildMapping);
            $("#btnMoveUp").click(MoveUp);
            $("#btnMoveDown").click(MoveDown);
            $("#btnSaveNameOptions").click(SaveNamingOptions);
            $("#btnAddCustomText").click(GoToAddCustomText);
            $("#btnCustomTextBack").click(GoToNamingOption);
            $("#btnSendData").click(confirmSendData);

            $("#btnInfoBack").click(BackToInfo);
            $("#btnDefaultsBack").click(BackFromNamingOptions);
            $("#btnDefaultsSubFolderNameEdit").click(GoToFolderNamingOptions);
            $("#btnDefaultsFileNameEdit").click(GoToFileNamingOptions);
            $("#btnNamingOptionsBack").click(GoToDefaults);
            $("#btnSetBack").click(BackToSet);
            $("#btnEditFieldsMappingBack").click(BackToEditSet);
            $("#btnNewSetBack").click(BackToSet);
            $("#btnSelectedSetBack").click(BackToSelectedSet);
            $("#btnTableRelationshipBack").click(BackToEditFieldsMapping);
            $("#btnParentChildTableRelationshipBack").click(BackToTableRelationship);
            $("#btnSaveCustomText").click(SaveCustomText);
            $('#btnDownloadExcel').on('click.open', function (e) {
                //e.preventDefault();
                DownloadExcelTemplate();
            });
            //$("#btnDownloadExcel").click(DownloadExcelTemplate);

            $("#btnPrivacyPolicy").on("click", function () {
                PanelOpen($(this).attr("InfoName"));
            });
            $("#btnFancyChartPanel").on("click", function () {
                PanelOpen($(this).attr("InfoName"));
            });
            $("#infoLearnMore").click(OpenInfoUrl);

            $("#setofFormsLearnMore").click(OpenSetofFormsUrl);
            $("#setofforms").click(SetOfForms);
            if (localStorage.getItem("CompanyName") === "") {
                $("#CompName").html("Super Admin's Forms List");
            } else {
                $("#CompName").html(localStorage.getItem("CompanyName") + "'s Forms List");
            }

            if (localStorage.getItem("isSaved") === "true") {
                localStorage.setItem("isSaved", false);
                $('#info').hide();
                $('#templateSets').show();
                $("#accordion").accordion({ active: 1 });
            }
            if (localStorage.getItem("isTemplateSetEdited") === "true") {
                localStorage.setItem("isTemplateSetEdited", false);
                $('#info').hide();
                $('#templateSets').show();
                $("#accordion").accordion({ active: 1 });
            }
            $('#data').jstree({
                'core': {
                    "multiple": false,
                    'data': {
                        "url": "/api/TreeNode/GetFolders",
                        "data": function (node) {
                            return { "id": node.id.replace(/F/ig, "").replace(/T/ig, ""), "companyId": localStorage.getItem("CompanyID"), "teamId": teamId };
                        },
                        "dataType": "json",
                        "type": "get",
                        "error": function (jqXHR, textStatus, errorThrown) { $('#data').html("<h3>There was an error while loading data for this tree</h3><p>" + jqXHR.responseText + "</p>"); }
                    },
                    "check_callback": true
                },
                "contextmenu": {
                    "items": customMenu
                },
                "plugins": [
                    "contextmenu", "search", "types"
                ],
                'types': {
                    '#': { /* options */ },
                    'Folder': { /* options */ },
                    'File': { /* options */ }
                }
            }).on('hover_node.jstree', function (e, data) {
                if (data.node.original.title !== null) {
                    $("#" + data.node.id).prop('title', data.node.original.title);
                }
            }).on('loaded.jstree', function () {
                $("#data").jstree("open_all");
            });

            function customMenu(node) {
                var tree = $("#tree").jstree(true);
                var items = {
                    "Expand": {
                        "separator_before": false,
                        "separator_after": false,
                        "label": "Expand",
                        "action": function (obj) {
                            $("#data").jstree().open_node(node.id, function () { ; }, false);
                        }
                    },
                    "Collapse": {
                        "separator_before": false,
                        "separator_after": false,
                        "label": "Collapse",
                        "action": function (obj) {
                            $("#data").jstree().close_node(node.id, function () { ; }, false);
                        }
                    },
                    "AddFolder": {
                        "separator_before": false,
                        "separator_after": false,
                        "label": "Add Folder",
                        "action": function (obj) {
                            var ref = $('#data').jstree(true),
                                sel = ref.get_selected();
                            if (!sel.length) { return false; }
                            sel = sel[0];
                            sel = ref.create_node("#", { "type": "Folder" });
                            if (sel) {
                                ref.edit(sel);
                            }
                        }
                    },
                    "AddSubFolder": {
                        "separator_before": false,
                        "separator_after": false,
                        "label": "Add Sub Folder",
                        "action": function (obj) {
                            $("#data").jstree().open_node(node.id, function () {

                                var ref = $('#data').jstree(true),
                                    sel = ref.get_selected();
                                if (!sel.length) { return false; }
                                sel = sel[0];
                                sel = ref.create_node(sel, { "type": "Folder" });
                                if (sel) {
                                    ref.edit(sel);
                                }

                            }, false);
                        }
                    },
                    "AddNewSet": {
                        "separator_before": false,
                        "separator_after": false,
                        "label": "Add New Set",
                        "action": function (obj) {
                            $('#txtOuptputFileName').prop('disabled', true);
                            $('#txtSubfolderName').prop('disabled', true);
                            $('#txtSetName').val("");
                            $('#txtDescription').val("");
                            $('#btnUpload').hide();
                            $('#templateSets').hide();
                            $('#newSet').show();
                            $("#accordion").accordion({ active: 5 });
                        }
                    },
                    "DeleteFolder": {
                        "separator_before": false,
                        "separator_after": false,
                        "label": "Delete Folder",
                        "action": function (obj) {
                            confirmRemove();
                        }
                    },
                    "RenameFolder": {
                        "separator_before": false,
                        "separator_after": false,
                        "label": "Rename Folder",
                        "action": function (obj) {
                            var ref = $('#data').jstree(true),
                                sel = ref.get_selected();
                            if (!sel.length) { return false; }
                            sel = sel[0];
                            ref.edit(sel);
                        }
                    },
                    "SelectSet": {
                        "separator_before": false,
                        "separator_after": false,
                        "label": "Select Set",
                        "action": function (obj) {
                            $("#data").jstree().select_node(node.id, function () { ; }, false);
                            var id = node.id;
                            isDemoSet = node.original.IsDemo;
                            if (isDemoSet || IsAdmin() || IsSuperAdmin()) {
                                $("#divEditSet").show();
                            } else {
                                $("#divEditSet").hide();
                            }

                            if (id !== undefined && id.indexOf("T") !== -1) {
                                var templateId = parseInt(id.replace("T", ""));

                                Excel.run(function (context) {
                                    var xmlpart = context.workbook.customXmlParts.getByNamespace("exceltoforms").getOnlyItem();
                                    xmlpart.load("id");
                                    return context.sync()
                                        .then(function () {
                                            if (xmlpart.id) {
                                                $.ajax({
                                                    url: "/api/Template/GetExcelVersion",
                                                    type: 'Get',
                                                    data: {
                                                        templateId: templateId
                                                    },
                                                    contentType: 'application/json;charset=utf-8'
                                                }).done(function (res) {
                                                    if (res.Error !== null && res.Error !== "") {
                                                        app.showNotification('Error', res.Error);
                                                    }
                                                    else if (res.ExcelVersion === null || res.ExcelVersion === "") {
                                                        $('#divUpdateInfo').html('<b>Modified By </b>' + res.UpdatedBy + " on " + res.UpdatedOn);
                                                        $('#divSendData').hide();
                                                        $('#divDownloadExcel').hide();
                                                        app.showNotification('Error', 'Custom xml part not found.');
                                                    }
                                                    else if (res === xmlpart.id) {
                                                        $('#divUpdateInfo').html('<b>Modified By </b>' + res.UpdatedBy + " on " + res.UpdatedOn);
                                                        isExcelVersioMatching = true;
                                                        $('#divEditSet').show();
                                                        $('#divSendData').show();
                                                        $('#divDownloadExcel').hide();
                                                        $('#divDownloadExcelNote').hide();
                                                        $('#templateSets').hide();
                                                        $('#selectedSet').show();
                                                        $("#accordion").accordion({ active: 2 });
                                                        //app.showNotification('Message', 'saved customxmlid=' + res + ', excel customxmlid=' + xmlpart.id);
                                                    }
                                                    else {
                                                        $('#divUpdateInfo').html('<b>Modified By </b>' + res.UpdatedBy + " on " + res.UpdatedOn);
                                                        isExcelVersioMatching = false;
                                                        $('#divEditSet').hide();
                                                        $('#divSendData').hide();
                                                        $('#divDownloadExcel').show();
                                                        $('#divDownloadExcelNote').show();
                                                        $('#templateSets').hide();
                                                        $('#selectedSet').show();
                                                        $("#accordion").accordion({ active: 2 });
                                                        //app.showNotification('Message', 'save customxmlid=' + res + ', excel customxmlid=' + xmlpart.id);
                                                    }
                                                }).fail(function (status) {
                                                    app.showNotification('Error', status.responseText);
                                                });
                                            }
                                            else {
                                                $.ajax({
                                                    url: "/api/Template/GetExcelVersion",
                                                    type: 'Get',
                                                    data: {
                                                        templateId: templateId
                                                    },
                                                    contentType: 'application/json;charset=utf-8'
                                                }).done(function (res) {
                                                    if (res.Error === null || res.Error === "") {
                                                        $('#divUpdateInfo').html('<b>Modified By </b>' + res.UpdatedBy + " on " + res.UpdatedOn);
                                                        $('#divSendData').hide();
                                                        $('#divEditSet').hide();
                                                        $('#divDownloadExcel').show();
                                                        $('#divDownloadExcelNote').hide();
                                                        $('#templateSets').hide();
                                                        $('#selectedSet').show();
                                                        $("#accordion").accordion({ active: 2 });
                                                        //app.showNotification('Message', 'excel customxmlid=' + xmlpart.id);
                                                    } else {
                                                        app.showNotification('Error', res.Error);
                                                    }
                                                }).fail(function (status) {
                                                    app.showNotification('Error', status.responseText);
                                                });
                                            }
                                        }).catch(function (error) {
                                            if (error.message === "This operation is not permitted for the current object.") {
                                                $.ajax({
                                                    url: "/api/Template/GetExcelVersion",
                                                    type: 'Get',
                                                    data: {
                                                        templateId: templateId
                                                    },
                                                    contentType: 'application/json;charset=utf-8'
                                                }).done(function (res) {
                                                    if (res.Error === null || res.Error === "") {
                                                        $('#divUpdateInfo').html('<b>Modified By </b>' + res.UpdatedBy + " on " + res.UpdatedOn);
                                                        $('#divSendData').hide();
                                                        $('#divEditSet').hide();
                                                        $('#divDownloadExcel').show();
                                                        $('#divDownloadExcelNote').hide();
                                                        $('#templateSets').hide();
                                                        $('#selectedSet').show();
                                                        $("#accordion").accordion({ active: 2 });
                                                        //app.showNotification('Message', 'excel customxmlid=undefined');
                                                    } else {
                                                        app.showNotification('Error', res.Error);
                                                    }
                                                }).fail(function (status) {
                                                    app.showNotification('Error', status.responseText);
                                                });
                                            } else {
                                                app.showNotification('Error', error.message);
                                            }
                                        });
                                });
                            }
                        }
                    },
                    "DuplicateSet": {
                        "separator_before": false,
                        "separator_after": false,
                        "label": "Duplicate Set",
                        "action": function (obj) {
                            var ref = $('#data').jstree(true),
                                sel = ref.get_selected();
                            if (!sel.length) { return false; }
                            sel = sel[0];
                            ref.copy(sel);
                        }
                    },
                    "DeleteSet": {
                        "separator_before": false,
                        "separator_after": false,
                        "label": "Delete Set",
                        "action": function (obj) {
                            confirmRemove();
                        }
                    },
                    "RenameSet": {
                        "separator_before": false,
                        "separator_after": false,
                        "label": "Rename Set",
                        "action": function (obj) {
                            var ref = $('#data').jstree(true),
                                sel = ref.get_selected();
                            if (!sel.length) { return false; }
                            sel = sel[0];
                            ref.edit(sel);
                        }
                    }
                };

                if (IsAdmin()) {
                    if (node.type === 'File') {
                        if (node.original.IsDemo) {
                            delete items.Expand;
                            delete items.Collapse;
                            delete items.DeleteFolder;
                            delete items.RenameFolder;
                            delete items.AddNewSet;
                            delete items.AddSubFolder;
                            delete items.DuplicateSet;
                            delete items.RenameSet;
                            delete items.DeleteSet;
                            delete items.AddFolder;
                        } else {
                            delete items.Expand;
                            delete items.Collapse;
                            delete items.DeleteFolder;
                            delete items.RenameFolder;
                            delete items.AddNewSet;
                            delete items.AddSubFolder;
                        }
                    }
                    else if (node.type === 'Folder') {
                        if (node.original.IsDemo) {
                            delete items.DeleteFolder;
                            delete items.RenameFolder;
                            delete items.AddNewSet;
                            delete items.AddSubFolder;
                            delete items.SelectSet;
                            delete items.DuplicateSet;
                            delete items.RenameSet;
                            delete items.DeleteSet;
                        } else {
                            delete items.SelectSet;
                            delete items.DuplicateSet;
                            delete items.RenameSet;
                            delete items.DeleteSet;
                        }
                    }
                }
                else if (IsSuperAdmin()) {
                    if (node.type === 'File') {
                        delete items.Expand;
                        delete items.Collapse;
                        delete items.DeleteFolder;
                        delete items.RenameFolder;
                        delete items.AddNewSet;
                        delete items.AddSubFolder;
                        delete items.DuplicateSet;
                        delete items.AddFolder;
                    }
                    else if (node.type === 'Folder') {
                        delete items.SelectSet;
                        delete items.DuplicateSet;
                        delete items.RenameSet;
                        delete items.DeleteSet;
                        if (node.original.IsDemo) {
                            delete items.DeleteFolder;
                            delete items.RenameFolder;
                            delete items.AddSubFolder;
                        }
                    }
                }
                else {
                    if (node.type === 'File') {
                        delete items.Expand;
                        delete items.Collapse;
                        delete items.DeleteFolder;
                        delete items.RenameFolder;
                        delete items.AddNewSet;
                        delete items.AddSubFolder;
                        delete items.DuplicateSet;
                        delete items.RenameSet;
                        delete items.DeleteSet;
                        delete items.AddFolder;
                        delete items.DuplicateSet;
                        delete items.RenameSet;
                        delete items.DeleteSet;
                    }
                    else if (node.type === 'Folder') {
                        delete items.SelectSet;
                        delete items.DuplicateSet;
                        delete items.RenameSet;
                        delete items.DeleteSet;
                        delete items.DeleteFolder;
                        delete items.RenameFolder;
                        delete items.AddNewSet;
                        delete items.AddFolder;
                        delete items.AddSubFolder;
                    }
                }

                return items;
            }

            function OpenInfoUrl() {
                window.open('http://exceltoforms.com', '_blank');
            }

            function OpenSetofFormsUrl() {
                window.open('http://exceltoforms.com', '_blank');
            }

            function confirmRemove() {
                $("#dialog-confirm").dialog("open");
            }

            function SetOfForms() {
                $('#info').hide();
                $('#templateSets').show();
                $("#accordion").accordion({ active: 1 });
            }

            function GetMappedPercentage(fileId) {
                jQuery.ajax({
                    url: '/api/Template/GetTemplateFileFieldsMappedPercentage/' + fileId,
                    success: function (result) {
                        return result;
                    },
                    async: false
                });
            }

            function BackToInfo() {
                $('#info').show();
                $('#templateSets').hide();
                $("#accordion").accordion({ active: 0 });
            }

            function BackToSet() {
                if (localStorage.getItem("isTemplateSetEdited") === "true") {
                    window.location.href = '../DashBoard/DashBoard.html';
                } else {
                    $('#selectedSet').hide();
                    $('#newSet').hide();
                    $('#templateSets').show();
                    $("#accordion").accordion({ active: 1 });
                }
            }

            function BackToSelectedSet() {
                deletedFiles = [];
                var id = localStorage.getItem("FolderId");
                var templateId = parseInt(id.replace("T", ""));

                Excel.run(function (context) {
                    //var settings = context.workbook.settings;
                    //var xmlPartIDSetting = settings.getItemOrNullObject("XmlPartId").load("value");

                    return context.sync()
                        .then(function () {
                            if (isExcelVersionUpdated) {
                                $.ajax({
                                    url: "/api/Template/GetExcelVersion",
                                    type: 'Get',
                                    data: {
                                        templateId: templateId
                                    },
                                    contentType: 'application/json;charset=utf-8'
                                }).done(function (res) {
                                    isExcelVersionUpdated = false;
                                    if (res.Error !== null) {
                                        app.showNotification('Error', res.Error);
                                    }
                                    else if (res.ExcelVersion === excelVersion) {
                                        isExcelVersioMatching = true;
                                        $('#divEditSet').show();
                                        $('#divSendData').show();
                                        $('#divDownloadExcel').hide();
                                        $('#editSet').hide();
                                        $('#selectedSet').show();
                                        $("#accordion").accordion({ active: 2 });
                                    }
                                    else {
                                        isExcelVersioMatching = false;
                                        $('#divEditSet').hide();
                                        $('#divSendData').hide();
                                        $('#divDownloadExcel').show();
                                        $('#editSet').hide();
                                        $('#selectedSet').show();
                                        $("#accordion").accordion({ active: 2 });
                                    }
                                }).fail(function (status) {
                                    app.showNotification('Error', 'Could not communicate with the server.');
                                });
                            } else {
                                $('#editSet').hide();
                                $('#selectedSet').show();
                                $("#accordion").accordion({ active: 2 });
                            }
                        });
                });
            }

            $("#dialog-deleteFile").dialog({
                autoOpen: false,
                resizable: false,
                height: "auto",
                width: 220,
                modal: true,
                draggable: false,
                buttons: {
                    "Delete": function () {
                        DeleteTemplateFile();
                        $(this).dialog("close");
                    },
                    Cancel: function () {
                        $(this).dialog("close");
                    }
                }
            });

            $("#dialog-removeAllMappings").dialog({
                autoOpen: false,
                resizable: false,
                height: "auto",
                width: 290,
                modal: true,
                draggable: false,
                buttons: {
                    "Remove": function () {
                        RemoveTemplateSetMapping();
                        $(this).dialog("close");
                    },
                    Cancel: function () {
                        $(this).dialog("close");
                    }
                }
            });

            $("#dialog-removeFileMappings").dialog({
                autoOpen: false,
                resizable: false,
                height: "auto",
                width: 290,
                modal: true,
                draggable: false,
                buttons: {
                    "Remove": function () {
                        RemoveFileMapping();
                        $(this).dialog("close");
                    },
                    Cancel: function () {
                        $(this).dialog("close");
                    }
                }
            });

            $("#dialog-autoMappings").dialog({
                autoOpen: false,
                resizable: false,
                height: "auto",
                width: 290,
                modal: true,
                draggable: false,
                buttons: {
                    "Auto Map": function () {
                        AutoMapFields();
                        $(this).dialog("close");
                    },
                    Cancel: function () {
                        $(this).dialog("close");
                    }
                }
            });

            $("#dialog-notification").dialog({
                autoOpen: false,
                resizable: false,
                height: "auto",
                width: 290,
                modal: true,
                draggable: true,
                buttons: {
                    Close: function () {
                        $(this).dialog("close");
                    }
                }
            });

            $("#dialog-editFieldsMappingBack").dialog({
                autoOpen: false,
                resizable: false,
                height: "auto",
                width: 300,
                modal: true,
                draggable: false,
                buttons: {
                    "Back": function () {
                        if (isFileMappingEdited) {
                            var id = parseInt(localStorage.getItem("FolderId").replace("T", ""));
                            var isdemo = localStorage.getItem("IsDemo");
                            $.ajax({
                                url: "/api/Template/GetTemplateById",
                                type: 'Get',
                                data: {
                                    id: id
                                },
                                contentType: 'application/json;charset=utf-8'
                            }).done(function (data) {
                                template = data;
                                BindTemplateFiles(data);
                                if (isdemo === "false") {
                                    $(document).on('click', '.fileDel', function () {
                                        DeleteFile(this);
                                    });
                                } else if (IsSuperAdmin()) {
                                    $(document).on('click', '.fileDel', function () {
                                        DeleteFile(this);
                                    });
                                }
                                $(document).on('click', '.fileEdit', function () {
                                    EditFieldsMapping(this);
                                });

                                isFileMappingEdited = false;
                                $("#dialog-editFieldsMappingBack").dialog("close");
                                $('#defaults').hide();
                                $('#editFieldsMapping').hide();
                                $('#editSet').show();
                                $("#accordion").accordion({ active: 3 });
                            }).fail(function (status) {
                                app.showNotification('Error', 'Could not communicate with the server.');
                            }).always(function () {
                                setTimeout(function () { $('#btnEditFieldsMappingBack').prop('disabled', false); }, 250);
                            });
                        } else {
                            $(this).dialog("close");
                            $('#defaults').hide();
                            $('#editFieldsMapping').hide();
                            $('#editSet').show();
                            $("#accordion").accordion({ active: 3 });
                        }
                    },
                    Cancel: function () {
                        $(this).dialog("close");
                    }
                }
            });

            $("#dialog-confirm").dialog({
                autoOpen: false,
                resizable: false,
                height: "auto",
                width: 220,
                modal: true,
                draggable: false,
                buttons: {
                    "Delete": function () {
                        var ref = $('#data').jstree(true),
                            sel = ref.get_selected();
                        if (!sel.length) { return false; }
                        ref.delete_node(sel);
                        $(this).dialog("close");
                    },
                    Cancel: function () {
                        $(this).dialog("close");
                    }
                }
            });

            $("#dialog-sendData").dialog({
                autoOpen: false,
                resizable: false,
                height: "auto",
                width: 340,
                modal: true,
                draggable: false,
                buttons: {
                    "OK": function () {
                        $(this).dialog("close");
                        SendDataToTemplateSet();
                    },
                    Cancel: function () {
                        $(this).dialog("close");
                    }
                }
            });

            $('#data').on('rename_node.jstree', function (e, data) {
                if (data.node.original.IsDemo && data.node.type === "Folder") {
                    e.preventDefault();
                    return;
                }
                if (data.node.original.IsDemo && !IsSuperAdmin()) {
                    e.preventDefault();
                    return;
                } else {
                    var ParentId = data.node.parent;
                    var upsertTemplateFolder = {
                        ParentFolderId: data.node.parent,
                        FolderName: data.text,
                        Old: data.old,
                        FolderId: data.node.id,
                        Type: data.node.type,
                        CompanyId: localStorage.getItem("CompanyID"),
                        TeamId: teamId
                    };
                    $.ajax({
                        url: "/api/TreeNode/UpsertTemplateFolder",
                        type: 'post',
                        data: JSON.stringify(upsertTemplateFolder),
                        contentType: 'application/json;charset=utf-8'
                    }).done(function (data) {
                        if (ParentId === "#") {
                            $("#data").jstree("refresh");
                        } else
                            $("#data").jstree().refresh_node(ParentId);
                    }).fail(function (status) {
                        app.showNotification('Error', 'Could not communicate with the server.');
                    }).always(function () {
                        $('.disable-while-sending').prop('disabled', false);
                    });
                }
            });

            $('#data').on('copy.jstree', function (e, data) {
                var param = {
                    Id: data.node[0],
                    CompanyId: localStorage.getItem("CompanyID"),
                    UserId: localStorage.getItem("UserID"),
                    TeamId: teamId
                };
                $.ajax({
                    url: "/api/TreeNode/CreateDuplicateTemplate",
                    type: 'post',
                    data: JSON.stringify(param),
                    contentType: 'application/json;charset=utf-8'
                }).done(function (res) {
                    if (res === 'false') {
                        app.showNotification('Error', 'Could not communicate with the server.');
                    } else {
                        $("#data").jstree().refresh_node($("#data").jstree().get_parent(res));
                    }
                }).fail(function (status) {
                    app.showNotification('Error', 'Could not communicate with the server.');
                }).always(function () {
                    $('.disable-while-sending').prop('disabled', false);
                });

            });

            $('#data').on('delete_node.jstree', function (e, data) {
                var param = {
                    Id: data.node.id,
                    Type: data.node.type
                };
                $.ajax({
                    url: "/api/TreeNode/Delete",
                    type: 'post',
                    data: JSON.stringify(param),
                    contentType: 'application/json;charset=utf-8'
                }).done(function (res) {
                    console.log(res);
                }).fail(function (status) {
                    app.showNotification('Error', 'Could not communicate with the server.');
                }).always(function () {
                    $('.disable-while-sending').prop('disabled', false);
                });

            });

            $('#plugins4_q').keyup(function () {
                if (to) { clearTimeout(to); }
                to = setTimeout(function () {
                    var v = $('#plugins4_q').val();
                    $('#data').jstree(true).search(v);
                }, 250);
            });

            $('#data').on('changed.jstree', function (e, data) {
                if (data.node !== undefined) {
                    if (data.node.id.indexOf('T') !== -1) {
                        templateId = data.node.id.replace('T', '');
                    }
                    localStorage.setItem("FolderId", data.node.id);
                    localStorage.setItem("IsDemo", data.node.original.IsDemo);
                    localStorage.setItem("FolderName", data.node.text);
                    var path = data.instance.get_path(data.node, ' > ');
                    $('#event_result').html('<b>Selected Set of Forms: </b>' + localStorage.getItem("CompanyName") + "'s Forms List > " + path);
                }

            }).jstree();

            $('#data').on('dblclick.jstree', function (e) {
                var id = localStorage.getItem("FolderId");
                var tree = $(this).jstree();
                var node = tree.get_node(event.target);
                isDemoSet = node.original.IsDemo;
                if (isDemoSet || IsAdmin() || IsSuperAdmin()) {
                    $("#divEditSet").show();
                } else {
                    $("#divEditSet").hide();
                }
                if (id !== undefined && id.indexOf("T") !== -1) {
                    var templateId = parseInt(id.replace("T", ""));

                    Excel.run(function (context) {
                        var xmlpart = context.workbook.customXmlParts.getByNamespace("exceltoforms").getOnlyItem();
                        xmlpart.load("id");
                        return context.sync()
                            .then(function () {
                                if (xmlpart.id) {
                                    $.ajax({
                                        url: "/api/Template/GetExcelVersion",
                                        type: 'Get',
                                        data: {
                                            templateId: templateId
                                        },
                                        contentType: 'application/json;charset=utf-8'
                                    }).done(function (res) {
                                        excelVersion = res.ExcelVersion;
                                        if (res.Error !== null && res.Error !== "") {
                                            app.showNotification('Error', res.Error);
                                        }
                                        else if (res.ExcelVersion === null || res.ExcelVersion === "") {
                                            $('#divUpdateInfo').html('<b>Modified By </b>' + res.UpdatedBy + " on " + res.UpdatedOn);
                                            $('#divSendData').hide();
                                            $('#divDownloadExcel').hide();
                                            app.showNotification('Error', 'Custom xml part not found.');
                                        }
                                        else if (res.ExcelVersion === xmlpart.id) {
                                            $('#divUpdateInfo').html('<b>Modified By </b>' + res.UpdatedBy + " on " + res.UpdatedOn);
                                            isExcelVersioMatching = true;
                                            $('#divEditSet').show();
                                            $('#divSendData').show();
                                            $('#divDownloadExcel').hide();
                                            $('#divDownloadExcelNote').hide();
                                            $('#templateSets').hide();
                                            $('#selectedSet').show();
                                            $("#accordion").accordion({ active: 2 });
                                            //app.showNotification('Message', 'saved customxmlid=' + res + ', excel customxmlid=' + xmlpart.id);
                                        }
                                        else {
                                            $('#divUpdateInfo').html('<b>Modified By </b>' + res.UpdatedBy + " on " + res.UpdatedOn);
                                            isExcelVersioMatching = false;
                                            $('#divEditSet').hide();
                                            $('#divSendData').hide();
                                            $('#divDownloadExcel').show();
                                            $('#divDownloadExcelNote').show();
                                            $('#templateSets').hide();
                                            $('#selectedSet').show();
                                            $("#accordion").accordion({ active: 2 });
                                            //app.showNotification('Message', 'save customxmlid=' + res + ', excel customxmlid=' + xmlpart.id);
                                        }
                                    }).fail(function (status) {
                                        app.showNotification('Error', status.responseText);
                                    });
                                }
                                else {
                                    $.ajax({
                                        url: "/api/Template/GetExcelVersion",
                                        type: 'Get',
                                        data: {
                                            templateId: templateId
                                        },
                                        contentType: 'application/json;charset=utf-8'
                                    }).done(function (res) {
                                        if (res.Error === null || res.Error === "") {
                                            $('#divUpdateInfo').html('<b>Modified By </b>' + res.UpdatedBy + " on " + res.UpdatedOn);
                                            $('#divSendData').hide();
                                            $('#divEditSet').hide();
                                            $('#divDownloadExcel').show();
                                            $('#divDownloadExcelNote').hide();
                                            $('#templateSets').hide();
                                            $('#selectedSet').show();
                                            $("#accordion").accordion({ active: 2 });
                                            //app.showNotification('Message', 'excel customxmlid=' + xmlpart.id);
                                        } else {
                                            app.showNotification('Error', res.Error);
                                        }
                                    }).fail(function (status) {
                                        app.showNotification('Error', status.responseText);
                                    });
                                }
                            }).catch(function (error) {
                                if (error.message === "This operation is not permitted for the current object.") {
                                    $.ajax({
                                        url: "/api/Template/GetExcelVersion",
                                        type: 'Get',
                                        data: {
                                            templateId: templateId
                                        },
                                        contentType: 'application/json;charset=utf-8'
                                    }).done(function (res) {
                                        if (res.Error === null || res.Error === "") {
                                            $('#divUpdateInfo').html('<b>Modified By </b>' + res.UpdatedBy + " on " + res.UpdatedOn);
                                            $('#divSendData').hide();
                                            $('#divEditSet').hide();
                                            $('#divDownloadExcel').show();
                                            $('#divDownloadExcelNote').hide();
                                            $('#templateSets').hide();
                                            $('#selectedSet').show();
                                            $("#accordion").accordion({ active: 2 });
                                            //app.showNotification('Message', 'excel customxmlid=undefined');
                                        } else {
                                            app.showNotification('Error', res.Error);
                                        }
                                    }).fail(function (status) {
                                        app.showNotification('Error', status.responseText);
                                    });

                                } else {
                                    app.showNotification('Error', error.message);
                                }
                            });
                    });
                }
            });

            var _readFileDataUrl = function (input, callback) {
                var len = input.files.length, _files = [], res = [];
                /*var readFile = function (filePos) {
                    if (!filePos) {
                        callback(false, res);
                    } else {
                        var reader = new FileReader();
                        reader.onload = function (e) {
                            res.push(e.target.result);
                            readFile(_files.shift());
                        };
                        reader.readAsDataURL(filePos);
                    }
                };*/
                BindTemplateFiles(template);
                var table_body = '';
                excludeFiles = [];
                for (var x = 0; x < len; x++) {
                    var isFileExist = false;
                    if (template.Files !== null) {
                        for (var i = 0; i < template.Files.length; i++) {
                            if (template.Files[i].FileName === input.files[x].name) {
                                isFileExist = true;
                                excludeFiles.push(input.files[x].name);
                            }
                        }
                    }
                    if (!isFileExist) {
                        _files.push(input.files[x]);
                        table_body = '<tr>';

                        //table_body += '<td>';
                        //table_body += "<a class='fileAutoMap' id='-1'><span style='color: red;cursor:pointer;'> Auto Map </span></a>";
                        //table_body += '</td>';
                        table_body += '<td>';
                        table_body += "<a class='fileRemoveMap' id='-1'><span style='color: red;cursor:pointer;'> Remove Map </span></a>";
                        table_body += '</td>';

                        table_body += '<td>';
                        table_body += " " + input.files[x].name;
                        table_body += '</td>';
                        table_body += '<td>';
                        table_body += "  0%";
                        table_body += '</td>';

                        table_body += '<td>';
                        table_body += "<a class='fileEdit' id='-1'><span style='color: red;cursor:pointer;'> Edit </span></a>";
                        table_body += '</td>';
                        table_body += '<td>';
                        table_body += "<a class='fileDel' id='-1'><span style='color: red;cursor:pointer;'> Delete </span></a>";
                        table_body += '</td>';
                        table_body += '<td style="display: none;">';
                        table_body += "false";
                        table_body += '</td>';
                        table_body += '</tr>';

                        $('#tableFiles').append(table_body);
                    }
                }
                /*input.value = '';
                readFile(_files.shift());*/
            };

            $('#newFiles').on('change', function () {
                _readFileDataUrl(this, function (err, files) {
                    if (err) { return; }
                });
            });

            var worksheet;
            Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, function () {
                Excel.run(function (context) {
                    worksheet = context.workbook.worksheets.getActiveWorksheet();
                    var eventResult = worksheet.tables.onChanged.add(handleSelectionChange);

                    return context.sync()
                        .then(function () {
                            console.log("Event handler successfully registered for onSelectionChanged event in the worksheet.");
                        });
                });
            }).catch(errorHandlerFunction);

            function handleSelectionChange(event) {
                return Excel.run(function (context) {
                    return context.sync()
                        .then(function () {
                            couterCheck = 1;
                            var odlValue = event.details.valueBefore;
                            var newValue = event.details.valueAfter;
                            var table = worksheet.tables.getItem(event.tableId);

                            Excel.run(function (context) {
                                var headerRange = table.getHeaderRowRange().load("values");
                                return context.sync()
                                    .then(function () {
                                        var headerValues = headerRange.values;
                                        if (headerValues.length > 0) {
                                            var value = headerValues[0].join(',');
                                            var headerArray = value.split(',');
                                            var column;

                                            $.each(headerArray, function (index, headerValue) {
                                                if (headerValue === newValue) {
                                                    if (!isNaN(templateId)) {
                                                        column = {
                                                            TemplateId: templateId,
                                                            OldValue: odlValue,
                                                            NewValue: newValue
                                                        };
                                                    }
                                                    return false;
                                                }
                                            });

                                            if (IsAdmin() || IsSuperAdmin()) {
                                                if (couterCheck === 1 && !isNaN(column.TemplateId) && column.OldValue !== null && column.NewValue !== null && column.OldValue !== '' && column.NewValue !== '' && column.OldValue !== column.NewValue) {
                                                    couterCheck = 0;
                                                    $.ajax({
                                                        url: "/api/Template/UpdateExcelColumnMapping",
                                                        type: 'post',
                                                        data: JSON.stringify(column),
                                                        contentType: 'application/json;charset=utf-8'
                                                    }).done(function (res) {
                                                        if (res === "success") {
                                                            app.showNotification('Messsage', 'Mapping updated.');
                                                        }
                                                        else {
                                                            app.showNotification('Error', 'Something went wrong. Please try again.');
                                                        }
                                                    }).fail(function (status) {
                                                        app.showNotification('Error', 'Could not communicate with the server.');
                                                    });
                                                }
                                            } else {
                                                event.details.valueAfter = event.details.valueBefore;
                                            }
                                        }
                                        return context.sync();
                                    });
                            });
                        });
                });
            }

        });
    };

    $("#dd_team").change(function () {
        teamId = $('option:selected', this).id;
        alert(teamId + " " + $('option:selected', this).text());
    });

    function getDocumentAsCompressed(formData) {
        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 /*64 KB*/ },
            function (result) {
                if (result.status === "succeeded") {
                    var myFile = result.value;
                    var sliceCount = myFile.sliceCount;
                    var slicesReceived = 0, gotAllSlices = true, docdataSlices = [];

                    getSliceAsync(myFile, 0, sliceCount, gotAllSlices, docdataSlices, slicesReceived, formData);
                }
                else {
                    $('#btnSave').show();
                    $('#btnSaving').hide();
                    app.showNotification("Error:", result.error.message);
                }
            });
    }

    function getExcelVersionDocumentAsCompressed(excelVersion) {
        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 /*64 KB*/ },
            function (result) {
                if (result.status === "succeeded") {
                    var myFile = result.value;
                    var sliceCount = myFile.sliceCount;
                    var slicesReceived = 0, gotAllSlices = true, docdataSlices = [];

                    getExcelVersionSliceAsync(myFile, 0, sliceCount, gotAllSlices, docdataSlices, slicesReceived, excelVersion);
                }
                else {
                    app.showNotification("Error:", result.error.message);
                }
            });
    }

    function EditDefaults() {
        $.ajax({
            url: "/api/Template/GetNamingOptions",
            type: 'Get',
            data: {
                templateId: parseInt(templateId)
            },
            contentType: 'application/json;charset=utf-8'
        }).done(function (res) {
            if (res.Error !== null) {
                app.showNotification('Error', res.Error);
            } else {
                if (res.Names.length > 0) {
                    $('#txtDefaultsSubFolderName').val(res.Names[0].SubFolderName);
                    $('#txtDefaultsFileName').val(res.Names[0].FileNamePart);
                }
                $('#editSet').hide();
                $('#defaults').show();
                $("#accordion").accordion({ active: 6 });
            }
        }).fail(function (status) {
            app.showNotification('Error', status.responseText);
        });
    }

    function BackFromNamingOptions() {
        $('#defaults').hide();
        $('#editFieldsMapping').hide();
        $('#editSet').show();
        $("#accordion").accordion({ active: 3 });
    }

    function SendDataToTemplateSet() {
        $("#btnSendData").hide();
        $("#btnSendingData").show();
        $.ajax({
            url: "/api/Template/GetSendDataParam",
            type: 'Get',
            data: {
                templateId: parseInt(templateId)
            },
            contentType: 'application/json;charset=utf-8'
        }).done(function (res) {
            if (res.Error === null) {
                if (res.Params.length > 0) {
                    var sheetName = res.Params[0].SheetName;
                    var parentTableName = res.Params[0].ParentTable;
                    Excel.run(function (context) {
                        var childTablesHeader = [];
                        var childTablesBody = [];
                        var childTableNames = [];

                        var sheet = context.workbook.worksheets.getItem(sheetName);
                        var parentTable = sheet.tables.getItem(parentTableName);

                        var headerRange = parentTable.getHeaderRowRange().load("values");
                        var bodyRange = parentTable.getDataBodyRange().load("values");

                        $.each(res.Params, function (index, obj) {
                            if (obj.ChildTable !== null) {
                                childTableNames.push(obj.ChildTable);
                                var childTable = sheet.tables.getItem(obj.ChildTable);
                                var childHeaderRange = childTable.getHeaderRowRange().load("values");
                                var childBodyRange = childTable.getDataBodyRange().load("values");
                                childTablesHeader.push(childHeaderRange);
                                childTablesBody.push(childBodyRange);
                            }
                        });

                        return context.sync()
                            .then(function () {
                                var parentTableData = [];
                                var headerValues = headerRange.values;
                                var bodyValues = bodyRange.values;

                                if (headerValues.length > 0) {
                                    $.each(headerValues, function (headerIndex, headerValue) {
                                        var headers = [];
                                        headerValue.forEach(function (column) {
                                            headers.push(column)
                                        })
                                        parentTableData.push(headers.join('*'));
                                    });
                                }
                                if (bodyValues.length > 0) {
                                    $.each(bodyValues, function (bodyIndex, bodyValue) {
                                        parentTableData.push(bodyValue.join('*'));
                                    });
                                }


                                //parentTableData.push(headerValues.join('*'));
                                //parentTableData.push(bodyValues.join('*'));

                                var childTables = [];
                                if (childTablesHeader.length > 0) {
                                    $.each(childTablesHeader, function (headerIndex, childTableHeaderRange) {
                                        var childTablesData = [];
                                        var childHeaderValues = childTableHeaderRange.values;
                                        $.each(childHeaderValues, function (childHeaderIndex, childHeaderValue) {
                                            childTablesData.push(childHeaderValue.join('*'));
                                        });
                                        //childTablesData.push(childHeaderValues.join('*'));
                                        if (childTablesBody.length > 0) {
                                            $.each(childTablesBody, function (bodyIndex, childTableBodyRange) {
                                                var childBodyValues = childTableBodyRange.values;
                                                if (headerIndex === bodyIndex) {
                                                    $.each(childBodyValues, function (childBodyIndex, childBodyValue) {
                                                        childTablesData.push(childBodyValue.join('*'));
                                                    });
                                                    //childTablesData.push(childBodyValues.join('*'));
                                                    childTables.push(childTablesData);
                                                }
                                            });
                                        }
                                    });
                                }

                                var sendData = {
                                    ParentTableData: parentTableData,
                                    ChildTableData: childTables,
                                    TemplateId: parseInt(templateId),
                                    ParentTableName: parentTableName,
                                    ChildTableNames: childTableNames,
                                    UserId: parseInt(localStorage.getItem("UserID"))
                                };

                                $.ajax({
                                    url: "/api/Template/SendDataToTemplateSet",
                                    type: 'post',
                                    data: JSON.stringify(sendData),
                                    contentType: 'application/json;charset=utf-8'
                                }).done(function (sres) {
                                    $("#btnSendData").show();
                                    $("#btnSendingData").hide();
                                    if (sres.Error === "" || sres.Error === null) {
                                        var messages = '';
                                        if (sres.Message.length > 0) {
                                            messages = sres.Message.join('|');
                                        }
                                        if (messages === '') {
                                            messages = 'Sending data to template set is completed.';
                                        } else {
                                            messages = 'Sending data to template set is completed.|' + messages;
                                        }
                                        app.showNotification('Message', messages);
                                        if (sres.ZipPath !== "" && sres.ZipPath !== null) {
                                            Office.context.ui.openBrowserWindow(location.origin + sres.ZipPath);

                                            setTimeout(function () {
                                                $.ajax({
                                                    url: "/api/Template/DeleteZipFolder",
                                                    type: 'post',
                                                    data: JSON.stringify(sres.ZipPath),
                                                    contentType: 'application/json;charset=utf-8'
                                                }).done(function (dres) {
                                                    console.log(dres);
                                                });
                                            }, 5000);
                                        }
                                    } else {
                                        app.showNotification('Error', sres.Error);
                                    }
                                }).fail(function (status) {
                                    $("#btnSendData").show();
                                    $("#btnSendingData").hide();
                                    app.showNotification('Error', status.responseText);
                                });

                                return context.sync();
                            });
                    }).catch(errorHandlerFunction);
                } else {
                    $("#btnSendData").show();
                    $("#btnSendingData").hide();
                    app.showNotification('Error', "No fields found for the template set.");
                }

            } else {
                $("#btnSendData").show();
                $("#btnSendingData").hide();
                app.showNotification('Error', res.Error);
            }
        }).fail(function (status) {
            $("#btnSendData").show();
            $("#btnSendingData").hide();
            app.showNotification('Error', status.responseText);
        });
    }

    function BackToEditSet() {
        $.ajax({
            url: "/api/Template/BackFromEditFieldsMapping",
            type: 'Get',
            data: {
                templateFileId: editFileId
            },
            contentType: 'application/json;charset=utf-8'
        }).done(function (res) {
            if (!isXfa) {
                if (res.IsAnyFiedMapped) {
                    $('#dialog-editFieldsMappingBack').html('<p><span class="ui-icon ui-icon-alert" style="float:left; margin:12px 12px 20px 0;"></span>Are you sure you want to go back?</p>');
                    confirmEditFieldsMappingBack();
                }
                else {
                    $('#dialog-editFieldsMappingBack').html('<p><span class="ui-icon ui-icon-alert" style="float:left; margin:12px 12px 20px 0;"></span>There is no field mapped for the file. Are you sure you want to go back?</p>');
                    confirmEditFieldsMappingBack();
                }
            } else {
                if (!res.IsAnyFiedMapped) {
                    $('#dialog-editFieldsMappingBack').html('<p><span class="ui-icon ui-icon-alert" style="float:left; margin:12px 12px 20px 0;"></span>There is no field mapped for the file. Are you sure you want to go back?</p>');
                    confirmEditFieldsMappingBack();
                } else {
                    if (res.DynamicFieldsCount === 0) {
                        $('#dialog-editFieldsMappingBack').html('<p><span class="ui-icon ui-icon-alert" style="float:left; margin:12px 12px 20px 0;"></span>Are you sure you want to go back?</p>');
                        confirmEditFieldsMappingBack();
                    }
                    else if (res.ParentTable === null || res.ParentTable === "") {
                        $('#dialog-editFieldsMappingBack').html('<p><span class="ui-icon ui-icon-alert" style="float:left; margin:12px 12px 20px 0;"></span>There is no parent table found for dynamic fields. Are you sure you want to go back?</p>');
                        confirmEditFieldsMappingBack();
                    } else {
                        if (res.ChildTablesCount === 0) {
                            $('#dialog-editFieldsMappingBack').html('<p><span class="ui-icon ui-icon-alert" style="float:left; margin:12px 12px 20px 0;"></span>There is no child table found for dynamic fields. Are you sure you want to go back?</p>');
                            confirmEditFieldsMappingBack();
                        } else {
                            if (res.ParentChildRelationshipCount === 0) {
                                $('#dialog-editFieldsMappingBack').html('<p><span class="ui-icon ui-icon-alert" style="float:left; margin:12px 12px 20px 0;"></span>There is no relationship found between parent and child tables for dynamic fields. Are you sure you want to go back?</p>');
                                confirmEditFieldsMappingBack();
                            } else if (res.ParentChildRelationshipCount !== res.DynamicFieldsCount) {
                                $('#dialog-editFieldsMappingBack').html('<p><span class="ui-icon ui-icon-alert" style="float:left; margin:12px 12px 20px 0;"></span>Not all child tables are mapped with parent table for dynamic fields. Are you sure you want to go back?</p>');
                                confirmEditFieldsMappingBack();
                            } else {
                                $('#dialog-editFieldsMappingBack').html('<p><span class="ui-icon ui-icon-alert" style="float:left; margin:12px 12px 20px 0;"></span>Are you sure you want to go back?</p>');
                                confirmEditFieldsMappingBack();
                            }
                        }
                    }
                }
            }
        }).fail(function (status) {
            app.showNotification('Error', 'Could not communicate with the server.');
        });
    }

    function SaveNamingOptions() {
        var fields = '';
        var separator = '';
        if ($("#separator").val() === '-') {
            separator = '_';
        } else if ($("#separator").val() === '{space}') {
            separator = ' ';
        }
        if ($("#fields").val().length > 0) {
            var fieldArray = $("#fields").val().join(',').split(',');
            fields = fieldArray.join(separator);
        } else {
            app.showNotification('Message', 'Please select one or more available field(s).');
            return;
        }
        if (isFileNameOption) {
            var fileNamingOption = {
                TemplateId: templateId,
                Fields: fields,
                UpdatedBy: parseInt(localStorage.getItem("UserID"))
            };
            $.ajax({
                url: "/api/Template/UpdateFileNamingOption",
                type: 'post',
                data: JSON.stringify(fileNamingOption),
                contentType: 'application/json;charset=utf-8'
            }).done(function (res) {
                if (res !== null) {
                    isNamingOptionsModified = true;
                    $.ajax({
                        url: "/api/Template/GetNamingOptions",
                        type: 'Get',
                        data: {
                            templateId: parseInt(templateId)
                        },
                        contentType: 'application/json;charset=utf-8'
                    }).done(function (res) {
                        if (res.Error !== null) {
                            app.showNotification('Error', res.Error);
                        } else {
                            if (res.Names.length > 0) {
                                $('#txtDefaultsSubFolderName').val(res.Names[0].SubFolderName);
                                $('#txtDefaultsFileName').val(res.Names[0].FileNamePart);
                            }
                            $("#txtSetOuptputFileName").val(fields);
                            $("#divDefaults").show();
                            $("#divNamingOptions").hide();
                        }
                    }).fail(function (status) {
                        app.showNotification('Error', status.responseText);
                    });

                } else {
                    app.showNotification('Error', 'Something went wrong. Please try again.');
                }
            }).fail(function (status) {
                app.showNotification('Error', 'Could not communicate with the server.');
            });
        } else if (isSubFolderNameOption) {
            var folderNamingOption = {
                TemplateId: templateId,
                Fields: fields,
                UpdatedBy: parseInt(localStorage.getItem("UserID"))
            };
            $.ajax({
                url: "/api/Template/UpdateFolderNamingOption",
                type: 'post',
                data: JSON.stringify(folderNamingOption),
                contentType: 'application/json;charset=utf-8'
            }).done(function (res) {
                if (res !== null) {
                    isNamingOptionsModified = true;
                    $.ajax({
                        url: "/api/Template/GetNamingOptions",
                        type: 'Get',
                        data: {
                            templateId: parseInt(templateId)
                        },
                        contentType: 'application/json;charset=utf-8'
                    }).done(function (res) {
                        if (res.Error !== null) {
                            app.showNotification('Error', res.Error);
                        } else {
                            if (res.Names.length > 0) {
                                $('#txtDefaultsSubFolderName').val(res.Names[0].SubFolderName);
                                $('#txtDefaultsFileName').val(res.Names[0].FileNamePart);
                            }
                            $("#txtSetSubfolderName").val(fields);
                            $("#divDefaults").show();
                            $("#divNamingOptions").hide();
                        }
                    }).fail(function (status) {
                        app.showNotification('Error', status.responseText);
                    });

                } else {
                    app.showNotification('Error', 'Something went wrong. Please try again.');
                }
            }).fail(function (status) {
                app.showNotification('Error', 'Could not communicate with the server.');
            });
        }
    }

    function GoToFileNamingOptions() {
        $.ajax({
            url: '/api/Template/GetParentTableFields',
            type: 'Get',
            data: {
                templateId: templateId
            },
            contentType: 'application/json;charset=utf-8'
        }).done(function (res) {
            $('#fields').empty();
            if (res !== null) {
                isSubFolderNameOption = false;
                isFileNameOption = true;
                $('#fields').append('<option value="{OriginalFileName}">{OriginalFileName}</option>');
                $('#fields').append('<option value="{ID}">{ID}</option>');
                if (res.length > 0) {
                    $.each(res, function (index, field) {
                        $('#fields').append('<option value="{' + field + '}">{' + field + '}</option>');
                    });
                }
                if (isDemoSet) {
                    if (IsSuperAdmin()) {
                        $('#btnSaveNameOptions').prop('disabled', false);
                    } else {
                        $('#btnSaveNameOptions').prop('disabled', true);
                    }
                } else {
                    $('#btnSaveNameOptions').prop('disabled', false);
                }
                $("#divDefaults").hide();
                $("#divNamingOptions").show();

            } else {
                app.showNotification('Error', 'Something went wrong. Please try again.');
            }
        }).fail(function (status) {
            app.showNotification('Error', 'Could not communicate with the server.');
        });
    }

    function GoToFolderNamingOptions() {
        $.ajax({
            url: '/api/Template/GetParentTableFields',
            type: 'Get',
            data: {
                templateId: templateId
            },
            contentType: 'application/json;charset=utf-8'
        }).done(function (res) {
            $('#fields').empty();
            if (res !== null) {
                isSubFolderNameOption = true;
                isFileNameOption = false;
                $('#fields').append('<option value="{ID}">{ID}</option>');

                if (res.length > 0) {
                    $.each(res, function (index, field) {
                        $('#fields').append('<option value="{' + field + '}">{' + field + '}</option>');
                    });
                }
                if (isDemoSet) {
                    if (IsSuperAdmin()) {
                        $('#btnSaveNameOptions').prop('disabled', false);
                    } else {
                        $('#btnSaveNameOptions').prop('disabled', true);
                    }
                } else {
                    $('#btnSaveNameOptions').prop('disabled', false);
                }
                $("#divDefaults").hide();
                $("#divNamingOptions").show();

            } else {
                app.showNotification('Error', 'Something went wrong. Please try again.');
            }
        }).fail(function (status) {
            app.showNotification('Error', 'Could not communicate with the server.');
        });
    }

    function GoToDefaults() {
        $("#divDefaults").show();
        $("#divNamingOptions").hide();
    }

    function GoToAddCustomText() {
        if (isDemoSet) {
            if (IsSuperAdmin())
                $('#btnSaveCustomText').prop('disabled', false);
            else {
                $('#btnSaveCustomText').prop('disabled', true);
            }
        } else {
            $('#btnRemoveFileMapping').prop('disabled', false);
        }
        $("#divNamingOptions").hide();
        $("#divCustomText").show();
    }

    function GoToNamingOption() {
        $("#divCustomText").hide();
        $("#divNamingOptions").show();
    }

    function BackToEditFieldsMapping() {
        $('#ParentChildTable').hide();
        $('#editFieldsMapping').show();
        $("#accordion").accordion({ active: 4 });
    }

    function BackToTableRelationship() {
        if (isParentChildRelationshipSaved) {
            $.ajax({
                url: "/api/Template/GetTableParentChildTables",
                type: 'Get',
                data: {
                    fileId: editFileId
                },
                contentType: 'application/json;charset=utf-8'
            }).done(function (res) {
                isParentChildRelationshipSaved = false;
                if (res === null) {
                    app.showNotification('Error', 'Something went wrong. Please try again.');
                }
                else if (res.length > 0) {
                    BindParentChildTable(res);
                    $(document).on('click', '.mapFields', function () {
                        MapParentChildField(this);
                    });
                    $('#ParentChildTableRelationship').hide();
                    $('#ParentChildTable').show();
                    $("#accordion").accordion({ active: 7 });
                } else {
                    app.showNotification('Message', 'No parent and child table found for the file.');
                }
            }).fail(function (status) {
                app.showNotification('Error', 'Could not communicate with the server.');
            }).always(function () {
                setTimeout(function () { $('#btnParentChildTableRelationshipBack').prop('disabled', false); }, 250);
            });
        } else {
            $('#ParentChildTableRelationship').hide();
            $('#ParentChildTable').show();
            $("#accordion").accordion({ active: 7 });
        }
    }

    function confirmFileDelete() {
        $("#dialog-deleteFile").dialog("open");
    }

    $(document).on('click', '.fileRemoveMap', function () {
        var fileObj = this;
        if (fileObj !== undefined && fileObj !== null && $.trim(fileObj.innerText) === 'Remove Map') {
            if (fileObj.id === "-1")
                return;

            currentRemoveFileId = parseInt(fileObj.id);
        }

        $("#dialog-removeFileMappings").dialog("open");
    });
    function confirmRemoveAllMappings() {
        currentRemoveFileId = -1;
        $("#dialog-removeAllMappings").dialog("open");
    }

    function confirmRemoveFileMappings() {
        $("#dialog-removeFileMappings").dialog("open");
    }

    //$(document).on('click', '.fileAutoMap', function () {
    //    var fileObj = this;
    //    if (fileObj !== undefined && fileObj !== null && $.trim(fileObj.innerText) === 'Auto Map') {
    //        if (fileObj.id === "-1")
    //            return;

    //        currentTemplateFileId = parseInt(fileObj.id);
    //    }

    //    $("#dialog-autoMappings").dialog("open");
    //});
    function confirmAutoFieldsMappings() {
        arrFileMap = [];
        currentTemplateFileId = -1;
        $("#dialog-autoMappings").dialog("open");
    }

    function confirmEditFieldsMappingBack() {
        $("#dialog-editFieldsMappingBack").dialog("open");
    }
    function confirmSendData() {
        $("#dialog-sendData").dialog("open");
    }
    function SaveCustomText() {
        var customText = $.trim($('#txtCustomText').val());
        if (customText === '') {
            return;
        }
        $('#fields').append('<option value="' + customText + '">' + customText + '</option>');
        $("#divCustomText").hide();
        $("#divNamingOptions").show();
    }

    function DownloadExcelTemplate() {
        $('#btnDownloadExcel').hide();
        $('#btnDownloadingExcel').show();
        $.ajax({
            type: 'get',
            cache: false,
            url: '/api/Template/DownloadExcelTemplate',
            data: {
                templateId: templateId,
                userId: parseInt(localStorage.getItem("UserID"))
            },
            contentType: 'application/json;charset=utf-8'
        }).done(function (res) {
            if (res !== null && res !== "") {
                Office.context.ui.openBrowserWindow(location.origin + res);
                setTimeout(function () {
                    $.ajax({
                        url: "/api/Template/DeleteExcelTemplateFile",
                        type: 'post',
                        data: JSON.stringify(res),
                        contentType: 'application/json;charset=utf-8'
                    }).done(function (dres) {
                        console.log(dres);
                    });
                }, 5000);
            } else {
                app.showNotification('Error', 'File not found.');
            }
            $('#btnDownloadExcel').show();
            $('#btnDownloadingExcel').hide();

        }).fail(function (status) {
            app.showNotification('Error', 'File not found.');
        });
    }

    function MoveUp() {
        if ($("#fields").val().length > 1)
            return;
        var selected = $("#fields").find(":selected");
        var before = selected.prev();
        if (before.length > 0)
            selected.detach().insertBefore(before);
    }

    function MoveDown() {
        if ($("#fields").val().length > 1)
            return;
        var selected = $("#fields").find(":selected");
        var next = selected.next();
        if (next.length > 0)
            selected.detach().insertAfter(next);
    }
    function addNumbersToDuplicates(a) {
        let counts = {}
        var newArray = [];
        for (let i = 0; i < a.length; i++) {
            if (counts[a[i]]) {
                counts[a[i]] += 1;
                newArray.push(a[i] + "--" + counts[a[i]]);
            } else {
                counts[a[i]] = 1;
                newArray.push(a[i]);
            }
        }
        return newArray;
    }

    function SyncMappedFields() {          

        if (templateId > 0) {

            $('#btnAutomapFields').hide();
            $('#btnAutomappingFields').show();
            $.ajax({
                url: "/api/Template/SyncMappedFields",
                type: 'get',
                data: { templateId: templateId },
                contentType: 'application/json;charset=utf-8'
            }).done(function (res) {

                if (res) {

                    var previousParentId = '';
                    var dynamicFieldIds = [];
                    var dtParentFields = [];
                    var parentTableColumnNames = [];
                    parentTableColumnNames.push("ID");                   

                    $.each(res, function (index, obj) {

                        var templateFileMappingId = obj.TemplateFileMappingId;
                        var pdfFieldName = obj.PDFFieldName;
                        var isDynamic = obj.IsDynamic;
                        var fieldId = obj.FieldId;
                        var parentFieldId = obj.ParentFieldId;
                        var hasChildFields = obj.HasChildFields;

                        if (isDynamic) {
                            previousParentId = fieldId;
                            dynamicFieldIds.push(fieldId);
                            return true;
                        }
                        else if ((previousParentId !== null || previousParentId !== '') && previousParentId === parentFieldId)
                            return true;
                        else if (parentFieldId === null || parentFieldId === '') {
                            var parentField = {
                                TemplateFileMappingId: templateFileMappingId,
                                ExcelFieldName: null,
                                SheetName: "",
                                ExcelTableName: "",
                                IsMapped: false
                            };

                            dtParentFields.push(parentField);
                        }
                        else if (!hasChildFields) {
                            var parentField1 = {
                                TemplateFileMappingId: templateFileMappingId,
                                ExcelFieldName: pdfFieldName,
                                SheetName: "",
                                ExcelTableName: "",
                                IsMapped: true
                            };
                            parentTableColumnNames.push(parentField1.ExcelFieldName);
                            dtParentFields.push(parentField1);
                        }

                    });
                    console.log(parentTableColumnNames);
                    Excel.run(function (context) {
                        var columnCount = parentTableColumnNames.length;
                        var colname = GetColumnName(columnCount);
                        var range = "A1:" + colname + "1";

                        //var sheet = context.workbook.worksheets.getActiveWorksheet();

                        var sheet = context.workbook.worksheets.add();
                        sheet.load("name");
                        var parentTable = sheet.tables.add(range, true /*hasHeaders*/);
                        parentTable.load("name");
                        parentTableColumnNames = addNumbersToDuplicates(parentTableColumnNames);
                        parentTable.getHeaderRowRange().values = [parentTableColumnNames];
                        var childTables = [];
                        var dtChildFields = [];
                        var childHeaders = [];

                        for (var i = 0; i < dynamicFieldIds.length; i++) {
                            var newArray = res.filter(function (item) {
                                return item.ParentFieldId === dynamicFieldIds[i];
                            });


                            var start = columnCount + 2;
                            var startColName = GetColumnName(start);
                            columnCount = start + newArray.length;
                            var endColName = GetColumnName(columnCount);
                            range = startColName + "1:" + endColName + "1";
                            var childTableColNames = [];
                            childTableColNames.push("ID");

                            $.each(newArray, function (j, arrobj) {
                                var dynamicField = {
                                    TemplateFileMappingId: arrobj.TemplateFileMappingId,
                                    ExcelFieldName: arrobj.PDFFieldName,
                                    SheetName: "",
                                    ExcelTableName: i,
                                    IsMapped: true,
                                    ParentFieldId: arrobj.ParentFieldId
                                };

                                dtChildFields.push(dynamicField);
                                childTableColNames.push(arrobj.PDFFieldName);
                            });

                            var childTable = sheet.tables.add(range, true /*hasHeaders*/);


                            childTables.push(childTable);
                            childTable.load("name");
                            childTable.getHeaderRowRange().values = [childTableColNames];
                            var childHeaderRange = childTable.getHeaderRowRange().load("values");
                            childHeaders.push(childHeaderRange);
                        }

                        if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
                            sheet.getUsedRange().format.autofitColumns();
                            sheet.getUsedRange().format.autofitRows();
                        }

                        sheet.activate();

                        return context.sync()
                            .then(function () {
                                $.each(dtParentFields, function (k, parentField) {
                                    parentField.SheetName = sheet.name;
                                    parentField.ExcelTableName = parentTable.name;
                                });

                                $.each(dtChildFields, function (l, childField) {
                                    childField.SheetName = sheet.name;
                                    /*$.each(childHeaders, function (m, childHeader) {
                                        var headerValues = childHeader.values;
                                    });*/
                                    $.each(childTables, function (n, childTable) {
                                        if (childField.ExcelTableName === n)
                                            childField.ExcelTableName = childTable.name;
                                    });
                                });
                                var fields = {
                                    ParentFields: dtParentFields,
                                    ChildFields: dtChildFields,
                                    DynamicFieldIds: dynamicFieldIds,
                                    TemplateId: templateId
                                };                                

                            });
                    }).catch(errorHandlerFunction);
                }
                else if (res === null) {
                    app.showNotification('Error', 'Something went wrong. Please try again.');
                }
            }).fail(function (status) {
                app.showNotification('Error', 'Could not communicate with the server.');
            }).always(function () {
                setTimeout(function () {
                    $('#btnAutomapFields').show();
                    $('#btnAutomappingFields').hide();
                }, 2500);
            });
        }
    }

    function AutoMapFields() {
        //debugger;

        var fileId = 0;
        if (currentTemplateFileId != -1) {

            if (arrFileMap.indexOf(currentTemplateFileId) < 0) {
                arrFileMap.push(currentTemplateFileId);
            }
            fileId = currentTemplateFileId;
        }        

        if (templateId > 0) {

            $('#btnAutomapFields').hide();
            $('#btnAutomappingFields').show();
            $.ajax({
                url: "/api/Template/AutoMapFields",
                type: 'get',
                data: { templateId: templateId, templateFileId: fileId },
                contentType: 'application/json;charset=utf-8'
            }).done(function (res) {

                if (res) {

                    currentTemplateFileId = -1;

                    var previousParentId = '';
                    var dynamicFieldIds = [];
                    var dtParentFields = [];
                    var parentTableColumnNames = [];
                    parentTableColumnNames.push("ID");

                    if (arrFileMap.length > 0) {                        
                        res = res.filter(function (item) {
                            return arrFileMap.indexOf(item.TemplateFileId) > -1;
                        });
                    }
                 
                    $.each(res, function (index, obj) {

                        var templateFileMappingId = obj.TemplateFileMappingId;
                        var pdfFieldName = obj.PDFFieldName;
                        var isDynamic = obj.IsDynamic;
                        var fieldId = obj.FieldId;
                        var parentFieldId = obj.ParentFieldId;
                        var hasChildFields = obj.HasChildFields;

                        if (isDynamic) {
                            previousParentId = fieldId;
                            dynamicFieldIds.push(fieldId);
                            return true;
                        }
                        else if ((previousParentId !== null || previousParentId !== '') && previousParentId === parentFieldId)
                            return true;
                        else if (parentFieldId === null || parentFieldId === '') {
                            var parentField = {
                                TemplateFileMappingId: templateFileMappingId,
                                ExcelFieldName: null,
                                SheetName: "",
                                ExcelTableName: "",
                                IsMapped: false
                            };

                            dtParentFields.push(parentField);
                        }
                        else if (!hasChildFields) {
                            var parentField1 = {
                                TemplateFileMappingId: templateFileMappingId,
                                ExcelFieldName: pdfFieldName,
                                SheetName: "",
                                ExcelTableName: "",
                                IsMapped: true
                            };
                            parentTableColumnNames.push(parentField1.ExcelFieldName);
                            dtParentFields.push(parentField1);
                        }

                    });
                    
                    Excel.run(function (context) {
                        var columnCount = parentTableColumnNames.length;
                        var colname = GetColumnName(columnCount);
                        var range = "A1:" + colname + "1";

                        var sheet = context.workbook.worksheets.add();
                        sheet.load("name");
                        var parentTable = sheet.tables.add(range, true /*hasHeaders*/);
                        parentTable.load("name");

                        //Try start
                        
                        //var colname = GetColumnName(parentTableColumnNames.length);
                        //var sheet = context.workbook.worksheets.getItem("sheet1");
                        //var parentTableExisting = sheet.tables.getItem("Table2");

                        //var headerRange = parentTableExisting.getHeaderRowRange().load("values");
                        //var bodyRange = parentTableExisting.getDataBodyRange().load("values");
                        ////var arrExistingColCount = headerRange.load("values");
                                               
                        //var columnCountStartTest = headerRange.values.length + 1;
                        //var colnameStartTest = GetColumnName(columnCountStartTest);

                        //var columnCountEndTest = columnCountStartTest + parentTableColumnNames.length;
                        //var colnameEndTest = GetColumnName(columnCountEndTest);

                        //var range = colnameStartTest + "1:" + colnameEndTest + "1";

                        //var parentTable = sheet.tables.add(range, true /*hasHeaders*/);
                        //parentTable.load("name");
                      
                        //console.log(range);
                        //var columnCount = columnCountEndTest;
                        //console.log(columnCount);
                        //Try end
                    
                        parentTableColumnNames = addNumbersToDuplicates(parentTableColumnNames);
                        parentTable.getHeaderRowRange().values = [parentTableColumnNames];
                        var childTables = [];
                        var dtChildFields = [];
                        var childHeaders = [];

                        for (var i = 0; i < dynamicFieldIds.length; i++) {
                            var newArray = res.filter(function (item) {
                                return item.ParentFieldId === dynamicFieldIds[i];
                            });


                            var start = columnCount + 2;
                            var startColName = GetColumnName(start);
                            columnCount = start + newArray.length;
                            var endColName = GetColumnName(columnCount);
                            range = startColName + "1:" + endColName + "1";
                            var childTableColNames = [];
                            childTableColNames.push("ID");

                            $.each(newArray, function (j, arrobj) {
                                var dynamicField = {
                                    TemplateFileMappingId: arrobj.TemplateFileMappingId,
                                    ExcelFieldName: arrobj.PDFFieldName,
                                    SheetName: "",
                                    ExcelTableName: i,
                                    IsMapped: true,
                                    ParentFieldId: arrobj.ParentFieldId
                                };

                                dtChildFields.push(dynamicField);
                                childTableColNames.push(arrobj.PDFFieldName);
                            });

                            var childTable = sheet.tables.add(range, true /*hasHeaders*/);


                            childTables.push(childTable);
                            childTable.load("name");
                            childTable.getHeaderRowRange().values = [childTableColNames];
                            var childHeaderRange = childTable.getHeaderRowRange().load("values");
                            childHeaders.push(childHeaderRange);
                        }

                        if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
                            sheet.getUsedRange().format.autofitColumns();
                            sheet.getUsedRange().format.autofitRows();
                        }

                        sheet.activate();

                        return context.sync()
                            .then(function () {
                                $.each(dtParentFields, function (k, parentField) {
                                    parentField.SheetName = sheet.name;
                                    parentField.ExcelTableName = parentTable.name;
                                });

                                $.each(dtChildFields, function (l, childField) {
                                    childField.SheetName = sheet.name;
                                    /*$.each(childHeaders, function (m, childHeader) {
                                        var headerValues = childHeader.values;
                                    });*/
                                    $.each(childTables, function (n, childTable) {
                                        if (childField.ExcelTableName === n)
                                            childField.ExcelTableName = childTable.name;
                                    });
                                });
                                var fields = {
                                    ParentFields: dtParentFields,
                                    ChildFields: dtChildFields,
                                    DynamicFieldIds: dynamicFieldIds,
                                    TemplateId: templateId
                                };

                                $.ajax({
                                    url: "/api/Template/SaveAutomapFields",
                                    type: 'post',
                                    data: JSON.stringify(fields),
                                    contentType: 'application/json;charset=utf-8'
                                }).done(function (res) {
                                    if (res === "success") {
                                        EditSet();
                                    }
                                    else {
                                        app.showNotification('Error', res);
                                    }
                                }).fail(function (status) {
                                    app.showNotification('Error', 'Could not communicate with the server.');
                                }).always(function () {
                                    setTimeout(function () {
                                        $('#btnAutomapFields').show();
                                        $('#btnAutomappingFields').hide();
                                    }, 250);
                                });


                            });
                    }).catch(errorHandlerFunction);
                }
                else if (res === null) {
                    app.showNotification('Error', 'Something went wrong. Please try again.');
                }
            }).fail(function (status) {
                app.showNotification('Error', 'Could not communicate with the server.');
            }).always(function () {
                setTimeout(function () {
                    $('#btnAutomapFields').show();
                    $('#btnAutomappingFields').hide();
                }, 2500);
            });
        }
    }

    function errorHandlerFunction(error) {
        $('#btnAutomapFields').prop('disabled', false);
        $('#btnSave').prop('disabled', false);
        //$('#btnUpdateVersion').prop('disabled', false);
        $("#btnSendData").show();
        $("#btnSendingData").hide();
        $('#btnSave').show();
        $('#btnSaving').hide();
        //$('#btnUpdateVersion').show();
        //$('#btnUpdatingVersion').hide();
        $('#btnAutomapFields').show();
        $('#btnAutomappingFields').hide();
        app.showNotification('Error', error.message);
    }

    function GetColumnName(num) {
        for (var ret = '', a = 1, b = 26; (num -= a) >= 0; a = b, b *= 26) {
            ret = String.fromCharCode(parseInt((num % b) / a) + 65) + ret;
        }
        return ret;
    }

    function RemoveFileMapping() {

        if (currentRemoveFileId != -1) {
            editFileId = currentRemoveFileId;

            const index = arrFileMap.indexOf(currentRemoveFileId);
            arrFileMap.splice(index, 1);
            currentRemoveFileId = -1;
            currentTemplateFileId = -1;
        }

        if (editFileId > 0) {
            $('#btnRemoveFileMapping').hide();
            $('#btnRemovingFileMapping').show();
            $.ajax({
                url: "/api/Template/RemoveFileMappings",
                type: 'post',
                data: JSON.stringify(editFileId),
                contentType: 'application/json;charset=utf-8'
            }).done(function (res) {
                if (res.IsAnyFieldMapped) {
                    isFileMappingEdited = true;
                    //SyncMappedFields();
                    app.showNotification('Message', 'All Mappings are removed for the file.');
                } else if (!res.IsAnyFieldMapped) {
                    app.showNotification('Message', 'There is no field mapped for the file.');
                }
                else if (res.Error !== null) {
                    app.showNotification('Error', res.Error);
                }
            }).fail(function (status) {
                app.showNotification('Error', 'Could not communicate with the server.');
            }).always(function () {
                setTimeout(function () {
                    $('#btnRemoveFileMapping').show();
                    $('#btnRemovingFileMapping').hide();
                }, 250);
            });
        }
    }

    function DeleteFile(obj) {
        if (obj !== undefined && obj !== null && $.trim(obj.innerText) === 'Delete') {
            if (obj.id === "-1")
                return;
            var fileId = parseInt(obj.id);
            deletedFiles.push(fileId);
            confirmFileDelete();
        }
    }

    function RemoveAllMappings() {
        confirmRemoveAllMappings();
    }

    function GetTableParentChildTables() {
        $.ajax({
            url: "/api/Template/GetTableParentChildTables",
            type: 'Get',
            data: {
                fileId: editFileId
            },
            contentType: 'application/json;charset=utf-8'
        }).done(function (res) {
            if (res === null) {
                app.showNotification('Error', 'Something went wrong. Please try again.');
            }
            else if (res.length > 0) {
                BindParentChildTable(res);
                $(document).on('click', '.mapFields', function () {
                    MapParentChildField(this);
                });
                $('#editFieldsMapping').hide();
                $('#ParentChildTable').show();
                $("#accordion").accordion({ active: 7 });
            } else {
                app.showNotification('Message', 'No parent and child table found for the file.');
            }
        }).fail(function (status) {
            app.showNotification('Error', 'Could not communicate with the server.');
        }).always(function () {
            setTimeout(function () { $('#btnTableRelationship').prop('disabled', false); }, 250);
        });
    }

    function SaveParentChildMapping() {
        $('#btnSaveParentChildTableRelationship').prop('disabled', false);
        var parentTableField = $('#parent').val();
        var childTableField = $('#child').val();
        if (parentTableField === 'Select') {
            app.showNotification('Message', 'Please select parent field.');
            return;
        } else if (childTableField === 'Select') {
            app.showNotification('Message', 'Please select child field.');
            return;
        }

        var parentChildMapping = {
            TemplateFileId: editFileId,
            ParentTable: parentTable,
            ParentTableField: parentTableField,
            ChildTable: childTable,
            ChildTableField: childTableField
        };
        $.ajax({
            url: "/api/Template/SaveParentChildMapping",
            type: 'post',
            data: JSON.stringify(parentChildMapping),
            contentType: 'application/json;charset=utf-8'
        }).done(function (res) {
            if (res === "success") {
                isParentChildRelationshipSaved = true;
                BackToTableRelationship();
            }
            else {
                app.showNotification('Error', 'Something went wrong. Please try again.');
            }
        }).fail(function (status) {
            app.showNotification('Error', 'Could not communicate with the server.');
        }).always(function () {
            setTimeout(function () { $('#btnSaveParentChildTableRelationship').prop('disabled', false); }, 250);
        });
    }

    function MapParentChildField(obj) {
        if (obj !== undefined && obj !== null && $.trim(obj.innerText) === 'Set Relationship') {
            parentTable = $(obj).closest("tr").find("td:nth-child(1)")[0].innerText;
            childTable = $(obj).closest("tr").find("td:nth-child(2)")[0].innerText;
            if (isDemoSet) {
                if (IsSuperAdmin()) {
                    $('#btnSaveParentChildTableRelationship').prop('disabled', false);
                } else {
                    $('#btnSaveParentChildTableRelationship').prop('disabled', true);
                }
            } else {
                $('#btnSaveParentChildTableRelationship').prop('disabled', false);
            }
            $.ajax({
                url: '/api/Template/GetTableFields',
                type: 'Get',
                data: {
                    fileId: editFileId,
                    parentTable: parentTable,
                    childTable: childTable
                },
                contentType: 'application/json;charset=utf-8'
            }).done(function (res) {
                if (res !== null) {
                    $('#child').empty().append('<option selected="selected" value="Select">Select</option>');
                    $('#child').append('<option value="ID">ID</option>');
                    $('#parent').empty().append('<option selected="selected" value="Select">Select</option>');
                    $('#parent').append('<option value="ID">ID</option>');

                    $.each(res.ChildFields, function (index, child) {
                        $('#child').append('<option value="' + child + '">' + child + '</option>');
                    });
                    $.each(res.ParentFields, function (index, parent) {
                        $('#parent').append('<option value="' + parent + '">' + parent + '</option>');
                    });
                    if (res.ParentField !== '' && res.ParentField !== null) {
                        $('#parent').val(res.ParentField);
                    } else {
                        $('#parent').val('ID');
                    }
                    if (res.ChildField !== '' && res.ChildField !== null) {
                        $('#child').val(res.ChildField);
                    } else {
                        $('#child').val('ID');
                    }
                    $('#ParentChildTable').hide();
                    $('#ParentChildTableRelationship').show();
                    $("#accordion").accordion({ active: 8 });
                } else {
                    app.showNotification('Error', 'Something went wrong. Please try again.');
                }
            }).fail(function (status) {
                app.showNotification('Error', 'Could not communicate with the server.');
            });
        }
    }

    function BindParentChildTable(data) {
        var table_body = '<table id="tableParentChild" width="100%" class="table table-striped">'
            + '<thead><tr><th>Parent Table</th><th>Child Table</th><th>Mapped</th><th>Action</th></tr></thead >';
        for (var i = 0; i < data.length; i++) {
            table_body += '<tr>';
            table_body += '<td>';
            table_body += data[i].ParentTable;
            table_body += '</td>';
            table_body += '<td>';
            table_body += " " + data[i].ChildTable;
            table_body += '</td>';
            table_body += '<td>';
            table_body += " " + data[i].IsMapped;
            table_body += '</td>';
            table_body += '<td>';
            table_body += "<a class='mapFields'><span style='color: red;cursor:pointer;'> Set Relationship </span></a>";
            table_body += '</td>';

            table_body += '</tr>';
        }
        table_body += '</table>';
        $('#divParentChildTable').html(table_body);
    }

    function RemoveTemplateSetMapping() {

        if (templateId > 0) {
            $('#btnRemoveMappings').hide();
            $('#btnRemovingMappings').show();
            $.ajax({
                url: "/api/Template/RemoveAllMappings",
                type: 'post',
                data: JSON.stringify(templateId),
                contentType: 'application/json;charset=utf-8'
            }).done(function (res) {
                if (res.IsAnyFieldMapped) {
                    EditSet();
                    app.showNotification('Message', 'All Mappings are removed for the template set.');
                } else if (!res.IsAnyFieldMapped) {
                    app.showNotification('Message', 'There is no field mapped for the template set.');
                }
                else if (res.Error !== null) {
                    app.showNotification('Error', res.Error);
                }
            }).fail(function (status) {
                app.showNotification('Error', 'Could not communicate with the server.');
            }).always(function () {
                setTimeout(function () {
                    $('#btnRemoveMappings').show();
                    $('#btnRemovingMappings').hide();
                }, 250);
            });
        }
    }

    function EditFieldsMapping(obj) {
        //debugger;
        if (!isExcelVersioMatching) {
            app.showNotification('Message', 'The open excel file is not the valid mapping file.');
            return;
        }
        if (obj !== undefined && obj !== null && $.trim(obj.innerText) === 'Edit') {
            if (obj.id === "-1")
                return;

            var fileId = parseInt(obj.id);
            editFileId = fileId;
            console.log(editFileId);
            if ($(obj).parent().next()[0].innerText === "true") {
                isXfa = true;
                if ($(obj).parent().next().next()[0].innerText === "0") {
                    $("#btnTableRelationship").hide();
                } else {
                    $("#btnTableRelationship").show();
                }
            }
            else {
                isXfa = false;
                $("#btnTableRelationship").hide();
            }
            if (isDemoSet) {
                if (IsSuperAdmin())
                    $('#btnRemoveFileMapping').prop('disabled', false);
                else {
                    $('#btnRemoveFileMapping').prop('disabled', true);
                }
            } else {
                $('#btnRemoveFileMapping').prop('disabled', false);
            }
            $('#editdata').jstree('destroy');
            $('#editdata').jstree({
                'core': {
                    "multiple": false,
                    'data': {
                        "url": "/api/Template/BindFieldsMappingTreeView",
                        "data": function (node) {
                            /*if (node.parent === '#') {
                                rootNodeText = node.text;
                            }*/
                            return { "id": node.id, "fileId": fileId };
                        },
                        "dataType": "json",
                        "type": "get",
                        "error": function (jqXHR, textStatus, errorThrown) { $('#editdata').html("<h3>There was an error while loading data for this tree</h3><p>" + jqXHR.responseText + "</p>"); }
                    },
                    "check_callback": true
                },
                'plugins': [
                    "contextmenu", "types"
                ],
                'contextmenu': {
                    'items': EditDataContextMenu
                },
                'types': {
                    '#': { /* options */ },
                    'Folder': { /* options */ },
                    'SF': { /* Static Field */ },
                    'DF': { /* Dynamic Field */ },
                    'DE': { /* Dynamic Element */ }

                }
            }).on('hover_node.jstree', function (e, data) {
                if (data.node.original.title !== null) {
                    $("#" + data.node.id).prop('title', data.node.original.title);
                }
            }).on('select_node.jstree', function (e, data) {
                if (data.node !== undefined) {
                    if (data.node.parent === '#') {
                        $.ajax({
                            url: '/api/Template/GetRangeParameter',
                            type: 'Get',
                            data: {
                                fieldId: data.node.id
                            },
                            contentType: 'application/json;charset=utf-8'
                        }).done(function (res) {
                            if (res !== null && res.TableName !== '' && res.SheetName !== '') {
                                Excel.run(function (context) {
                                    var sheet = context.workbook.worksheets.getItem(res.SheetName);
                                    var table = sheet.tables.getItem(res.TableName);
                                    var range = table.getRange();
                                    range.select();

                                    return context.sync();
                                }).catch(errorHandlerFunction);
                            }
                        }).fail(function (status) {
                            app.showNotification('Error', 'Could not communicate with the server.');
                        });
                    } else if (data.node.original.isMapped && data.node.original.type === 'DF') {
                        $.ajax({
                            url: '/api/Template/GetRangeParameter',
                            type: 'Get',
                            data: {
                                fieldId: data.node.id
                            },
                            contentType: 'application/json;charset=utf-8'
                        }).done(function (res) {
                            if (res !== null && res.TableName !== '' && res.SheetName !== '') {
                                Excel.run(function (context) {
                                    var sheet = context.workbook.worksheets.getItem(res.SheetName);
                                    var table = sheet.tables.getItem(res.TableName);
                                    var range = table.getRange();
                                    range.select();

                                    return context.sync();
                                }).catch(errorHandlerFunction);
                            } else {
                                app.showNotification('Error', 'Something went wrong. Please try again.');
                            }
                        }).fail(function (status) {
                            app.showNotification('Error', 'Could not communicate with the server.');
                        });
                    } else if (data.node.original.isMapped && (data.node.original.type === 'DE' || data.node.original.type === 'SF')) {
                        $.ajax({
                            url: '/api/Template/GetRangeParameter',
                            type: 'Get',
                            data: {
                                fieldId: data.node.id
                            },
                            contentType: 'application/json;charset=utf-8'
                        }).done(function (res) {
                            if (res !== null && res.ColumnName !== '' && res.TableName !== '' && res.SheetName !== '') {
                                Excel.run(function (context) {
                                    var sheet = context.workbook.worksheets.getItem(res.SheetName);
                                    var table = sheet.tables.getItem(res.TableName);
                                    var column = table.columns.getItem(res.ColumnName);
                                    var range = column.getRange();
                                    range.select();

                                    return context.sync();
                                }).catch(errorHandlerFunction);
                            }
                        }).fail(function (status) {
                            app.showNotification('Error', 'Could not communicate with the server.');
                        });
                    }

                }

            }).on('loaded.jstree', function () {
                $("#editdata").jstree("open_all");
            });

            $('#editSet').hide();
            $('#editFieldsMapping').show();
            $("#accordion").accordion({ active: 4 });
        }
    }

    function EditDataContextMenu(node) {
        var tree = $("#editdata").jstree(true);
        var items = {
            "AddStaticField": {
                "separator_before": false,
                "separator_after": false,
                "label": "Add as Static Field",
                "action": function (obj) {
                    var asf = true;
                    if (isDemoSet) {
                        if (IsSuperAdmin()) {
                            asf = true;
                        } else {
                            asf = false;
                        }
                    }
                    if (asf) {
                        Excel.run(function (context) {
                            var sheet = context.workbook.worksheets.getActiveWorksheet();
                            var selection = context.workbook.getSelectedRange();
                            var tables = sheet.tables.load("name");
                            sheet.load("name");
                            selection.load("address");
                            return context.sync()
                                .then(function () {
                                    var intersections = [];
                                    for (var i = 0; i < tables.items.length; i++) {
                                        var table = tables.items[i];
                                        intersections[table.name] = table.getRange().
                                            getIntersectionOrNullObject(selection).load("address");
                                    }
                                    return context.sync()
                                        .then(function () {
                                            var found = false;
                                            var activeTable = '';
                                            for (var tableName in intersections) {
                                                var rangeOrNull = intersections[tableName];
                                                if (!rangeOrNull.isNullObject) {
                                                    found = true;
                                                    activeTable = tableName;
                                                    break;
                                                }
                                            }
                                            if (!found) {
                                                app.showNotification('Error', 'Please select cell inside a table.');
                                                return;
                                            }

                                            var tableColumns = context.workbook.tables.getItem(activeTable).columns;
                                            tableColumns.load('items');
                                            return context.sync()
                                                .then(function () {
                                                    var columnIntersections = [];
                                                    for (var j = 0; j < tableColumns.items.length; j++) {
                                                        var column = tableColumns.items[j];
                                                        columnIntersections[column.name] = column.getRange().
                                                            getIntersectionOrNullObject(selection).load("address");
                                                    }
                                                    return context.sync()
                                                        .then(function () {
                                                            var columnFound = false;
                                                            var activeColumnName = '';
                                                            for (var columnName in columnIntersections) {
                                                                var crangeOrNull = columnIntersections[columnName];
                                                                if (!crangeOrNull.isNullObject) {
                                                                    columnFound = true;
                                                                    activeColumnName = columnName;
                                                                    break;
                                                                }
                                                            }
                                                            if (!columnFound) {
                                                                app.showNotification('Error', 'Please select cell inside a table.');
                                                                return;
                                                            }

                                                            var mapFieldParam = {
                                                                TemplateFileMappingId: node.original.templateFileMappingId,
                                                                SheetName: sheet.name,
                                                                TableName: activeTable,
                                                                ColumnName: activeColumnName,
                                                                IsDynamicElement: node.original.type === "DE"
                                                            };
                                                            $.ajax({
                                                                url: "/api/Template/MapAcroField",
                                                                type: 'post',
                                                                data: JSON.stringify(mapFieldParam),
                                                                contentType: 'application/json;charset=utf-8'
                                                            }).done(function (res) {
                                                                if (res === "success") {
                                                                    isFileMappingEdited = true;
                                                                    node.original.icon = node.icon = "glyphicon glyphicon-file text-success";
                                                                    node.original.isMapped = true;
                                                                    $("#editdata").jstree("refresh");
                                                                } else {
                                                                    app.showNotification('Error', res);
                                                                }
                                                            }).fail(function (status) {
                                                                app.showNotification('Error', 'Could not communicate with the server.');
                                                            });
                                                        });
                                                });
                                        });
                                });
                        }).catch(errorHandlerFunction);
                    }
                }
            },
            "RemoveStaticField": {
                "separator_before": false,
                "separator_after": false,
                "label": "Remove Static Field",
                "action": function (obj) {
                    var rsf = true;
                    if (isDemoSet) {
                        if (IsSuperAdmin()) {
                            rsf = true;
                        } else {
                            rsf = false;
                        }
                    }
                    if (rsf) {
                        $.ajax({
                            url: "/api/Template/RemoveStaticFieldMapping",
                            type: 'post',
                            data: JSON.stringify(node.original.templateFileMappingId),
                            contentType: 'application/json;charset=utf-8'
                        }).done(function (res) {
                            if (res === "success") {
                                isFileMappingEdited = true;
                                node.original.icon = node.icon = "glyphicon glyphicon-file text-danger";
                                node.original.isMapped = false;
                                $("#editdata").jstree("refresh");
                            } else {
                                app.showNotification('Error', res);
                            }
                        }).fail(function (status) {
                            app.showNotification('Error', 'Could not communicate with the server.');
                        });
                    }
                }
            },
            "AddDynamicElement": {
                "separator_before": false,
                "separator_after": false,
                "label": "Add as Dynamic Element",
                "action": function (obj) {
                    var ade = true;
                    if (isDemoSet) {
                        if (IsSuperAdmin()) {
                            ade = true;
                        } else {
                            ade = false;
                        }
                    }
                    if (ade) {
                        Excel.run(function (context) {
                            var sheet = context.workbook.worksheets.getActiveWorksheet();
                            var selection = context.workbook.getSelectedRange();
                            var tables = sheet.tables.load("name");
                            sheet.load("name");
                            selection.load("address");
                            return context.sync()
                                .then(function () {
                                    var intersections = [];
                                    for (var i = 0; i < tables.items.length; i++) {
                                        var table = tables.items[i];
                                        intersections[table.name] = table.getRange().
                                            getIntersectionOrNullObject(selection).load("address");
                                    }
                                    return context.sync()
                                        .then(function () {
                                            var found = false;
                                            var activeTable = '';
                                            for (var tableName in intersections) {
                                                var rangeOrNull = intersections[tableName];
                                                if (!rangeOrNull.isNullObject) {
                                                    found = true;
                                                    activeTable = tableName;
                                                    break;
                                                }
                                            }
                                            if (!found) {
                                                app.showNotification('Error', 'Please select cell inside a table.');
                                                return;
                                            }

                                            var tableColumns = context.workbook.tables.getItem(activeTable).columns;
                                            tableColumns.load('items');
                                            return context.sync()
                                                .then(function () {
                                                    var columnIntersections = [];
                                                    for (var j = 0; j < tableColumns.items.length; j++) {
                                                        var column = tableColumns.items[j];
                                                        columnIntersections[column.name] = column.getRange().
                                                            getIntersectionOrNullObject(selection).load("address");
                                                    }
                                                    return context.sync()
                                                        .then(function () {
                                                            var columnFound = false;
                                                            var activeColumnName = '';
                                                            for (var columnName in columnIntersections) {
                                                                var crangeOrNull = columnIntersections[columnName];
                                                                if (!crangeOrNull.isNullObject) {
                                                                    columnFound = true;
                                                                    activeColumnName = columnName;
                                                                    break;
                                                                }
                                                            }
                                                            if (!columnFound) {
                                                                app.showNotification('Error', 'Please select cell inside a table.');
                                                                return;
                                                            }

                                                            var mapFieldParam = {
                                                                TemplateFileMappingId: node.original.templateFileMappingId,
                                                                SheetName: sheet.name,
                                                                TableName: activeTable,
                                                                ColumnName: activeColumnName,
                                                                IsDynamicElement: node.original.type === "DE"
                                                            };
                                                            $.ajax({
                                                                url: "/api/Template/MapXfaField",
                                                                type: 'post',
                                                                data: JSON.stringify(mapFieldParam),
                                                                contentType: 'application/json;charset=utf-8'
                                                            }).done(function (res) {
                                                                if (res === "success") {
                                                                    isFileMappingEdited = true;
                                                                    node.original.icon = node.icon = "glyphicon glyphicon-file text-success";
                                                                    node.original.isMapped = true;
                                                                    $("#editdata").jstree("refresh");
                                                                } else {
                                                                    app.showNotification('Error', res);
                                                                }
                                                            }).fail(function (status) {
                                                                app.showNotification('Error', 'Could not communicate with the server.');
                                                            });
                                                        });
                                                });
                                        });
                                });
                        }).catch(errorHandlerFunction);
                    }
                }
            },
            "RemoveDynamicField": {
                "separator_before": false,
                "separator_after": false,
                "label": "Remove Dynamic Field",
                "action": function (obj) {
                    var rdf = true;
                    if (isDemoSet) {
                        if (IsSuperAdmin()) {
                            rdf = true;
                        } else {
                            rdf = false;
                        }
                    }
                    if (rdf) {

                        var dynamicParam = {
                            TemplateFileMappingId: node.original.templateFileMappingId,
                            IsDynamicField: true
                        };
                        $.ajax({
                            url: "/api/Template/RemoveDynamicFieldMapping",
                            type: 'post',
                            data: JSON.stringify(dynamicParam),
                            contentType: 'application/json;charset=utf-8'
                        }).done(function (res) {
                            if (res === "success") {
                                isFileMappingEdited = true;
                                node.original.icon = node.icon = "glyphicon glyphicon-duplicate text-danger";
                                node.original.isMapped = false;
                                $("#editdata").jstree("refresh");
                            } else {
                                app.showNotification('Error', res);
                            }
                        }).fail(function (status) {
                            app.showNotification('Error', 'Could not communicate with the server.');
                        });
                    }
                }
            },
            "RemoveDynamicElement": {
                "separator_before": false,
                "separator_after": false,
                "label": "Remove Dynamic Element",
                "action": function (obj) {
                    var rde = true;
                    if (isDemoSet) {
                        if (IsSuperAdmin()) {
                            rde = true;
                        } else {
                            rde = false;
                        }
                    }
                    if (rde) {
                        var dynamicParam = {
                            TemplateFileMappingId: node.original.templateFileMappingId,
                            IsDynamicField: false
                        };
                        $.ajax({
                            url: "/api/Template/RemoveDynamicFieldMapping",
                            type: 'post',
                            data: JSON.stringify(dynamicParam),
                            contentType: 'application/json;charset=utf-8'
                        }).done(function (res) {
                            if (res === "success") {
                                isFileMappingEdited = true;
                                node.original.icon = node.icon = "glyphicon glyphicon-file text-danger";
                                node.original.isMapped = false;
                                $("#editdata").jstree("refresh");
                            } else {
                                app.showNotification('Error', res);
                            }
                        }).fail(function (status) {
                            app.showNotification('Error', 'Could not communicate with the server.');
                        });
                    }
                }
            }
        };

        if (node.original.isMapped) {
            var mappedNode = node.original.title;
            switch (mappedNode) {
                case "Static Field":
                    delete items.RemoveDynamicElement;
                    delete items.RemoveDynamicField;
                    delete items.AddDynamicElement;
                    delete items.AddStaticField;
                    break;
                case "Dynamic Section":
                    delete items.RemoveDynamicElement;
                    delete items.RemoveStaticField;
                    delete items.AddDynamicElement;
                    delete items.AddStaticField;
                    break;
                case "Dynamic Field":
                    delete items.AddStaticField;
                    delete items.RemoveStaticField;
                    delete items.AddDynamicElement;
                    delete items.RemoveDynamicField;
                    break;
                default:
                    delete items.AddStaticField;
                    delete items.RemoveStaticField;
                    delete items.AddDynamicElement;
                    delete items.RemoveDynamicElement;
                    delete items.RemoveDynamicField;
            }
        } else if (!node.original.isMapped) {
            var unmappedNode = node.original.title;
            switch (unmappedNode) {
                case "Static Field":
                    delete items.RemoveDynamicElement;
                    delete items.RemoveDynamicField;
                    delete items.AddDynamicElement;
                    delete items.RemoveStaticField;
                    break;
                case "Dynamic Section":
                    delete items.RemoveDynamicElement;
                    delete items.RemoveStaticField;
                    delete items.AddDynamicElement;
                    delete items.RemoveDynamicField;
                    delete items.AddStaticField;
                    break;
                case "Dynamic Field":
                    delete items.AddStaticField;
                    delete items.RemoveStaticField;
                    delete items.RemoveDynamicElement;
                    delete items.RemoveDynamicField;
                    break;
                default:
                    delete items.AddStaticField;
                    delete items.RemoveStaticField;
                    delete items.AddDynamicElement;
                    delete items.RemoveDynamicElement;
                    delete items.RemoveDynamicField;
            }
        } else if (node.original.title === null) {
            delete items.AddStaticField;
            delete items.RemoveStaticField;
            delete items.AddDynamicElement;
            delete items.RemoveDynamicElement;
            delete items.RemoveDynamicField;
        }

        return items;
    }

    function EditSet() {
        var id = parseInt(templateId);
        $('#btnEditSet').hide();
        $('#btnEditingSet').show();
        $.ajax({
            url: "/api/Template/GetTemplateById",
            type: 'Get',
            data: {
                id: id
            },
            contentType: 'application/json;charset=utf-8'
        }).done(function (res) {
            if (res !== null) {
                $('#txtExistingSetName').val(res.TemplateName);
                $('#txtSetDescription').val(res.Description);
                $('#txtSetSubfolderName').val(res.SubFolderName);
                $('#txtSetOuptputFileName').val(res.FileNamePart);
                $('#txtSetSubfolderName').attr('readonly', 'true');
                $('#txtSetOuptputFileName').attr('readonly', 'true');

                if (isDemoSet) {
                    if (IsSuperAdmin()) {
                        $('#btnSaveExistingSet').prop('disabled', false);
                        $('#btnAutomapFields').prop('disabled', false);
                        $('#btnRemoveMappings').prop('disabled', false);
                        //$('#btnUpdateVersion').prop('disabled', false);
                        $('#txtExistingSetName').prop('disabled', false);
                        $('#txtSetDescription').prop('disabled', false);
                    } else {
                        $('#btnSaveExistingSet').prop('disabled', true);
                        $('#btnAutomapFields').prop('disabled', true);
                        $('#btnRemoveMappings').prop('disabled', true);
                        //$('#btnUpdateVersion').prop('disabled', true);
                        $('#txtExistingSetName').prop('disabled', true);
                        $('#txtSetDescription').prop('disabled', true);
                    }
                }
                else {
                    $('#btnSaveExistingSet').prop('disabled', false);
                    $('#btnAutomapFields').prop('disabled', false);
                    $('#btnRemoveMappings').prop('disabled', false);
                    //$('#btnUpdateVersion').prop('disabled', false);
                    $('#txtExistingSetName').prop('disabled', false);
                    $('#txtSetDescription').prop('disabled', false);
                }
                template = res;
                BindTemplateFiles(res);
                if (isDemoSet) {
                    if (IsSuperAdmin()) {
                        $(document).on('click', '.fileDel', function () {
                            DeleteFile(this);
                        });
                    }
                } else {
                    $(document).on('click', '.fileDel', function () {
                        DeleteFile(this);
                    });
                }
                $(document).on('click', '.fileEdit', function () {
                    EditFieldsMapping(this);
                });

                $('#selectedSet').hide();
                $('#editSet').show();
                $("#accordion").accordion({ active: 3 });
            } else {
                app.showNotification('Error', 'Could not communicate with the server.');
            }
        }).fail(function (status) {
            app.showNotification('Error', 'Could not communicate with the server.');
        }).always(function () {
            setTimeout(function () {
                $('#btnEditSet').show();
                $('#btnEditingSet').hide();
            }, 2500);
        });
    }

    function SaveExistingSet() {
        console.log('Save');
        var delFiles = deletedFiles.join(',');
        var templateName = $.trim($('#txtExistingSetName').val());
        var description = $.trim($('#txtSetDescription').val());

        if (templateName !== "" && description !== "") {
            var formData = new FormData();
            var pdfFiles = $("#newFiles").get(0).files;
            if (pdfFiles.length > 0) {
                for (var i = 0; i < pdfFiles.length; i++) {
                    var f = pdfFiles[i];
                    formData.append("UploadedFiles", f);
                }
            }

            $('#btnSaveExistingSet').hide();
            $('#btnSavingExistingSet').show();

            var jeditPdfTemplate = {
                TemplateId: parseInt(localStorage.getItem("FolderId").replace("T", "")),
                DeletedTemplateFileIds: delFiles,
                TemplateName: templateName,
                Description: description,
                TemplateFileZip: null,
                UpdatedBy: parseInt(localStorage.getItem("UserID")),
                ExcelZip: null,
                IsFileDelete: false,
                SubFolderName: $('#txtSubfolderName').val(),
                FileNamePart: $('#txtOuptputFileName').val()
            };

            formData.append('TemplateId', jeditPdfTemplate.TemplateId);
            formData.append('TemplateName', jeditPdfTemplate.TemplateName);
            formData.append('Description', jeditPdfTemplate.Description);
            formData.append('SubFolderName', jeditPdfTemplate.SubFolderName);
            formData.append('FileNamePart', jeditPdfTemplate.FileNamePart);
            formData.append('DeletedTemplateFileIds', jeditPdfTemplate.DeletedTemplateFileIds);
            formData.append('UpdatedBy', jeditPdfTemplate.UpdatedBy);
            formData.append('ExcludeFiles', excludeFiles);

            $.ajax({
                url: '/api/Template/UpdateTemplate',
                type: 'post',
                contentType: false,
                processData: false,
                data: formData
            }).done(function (res) {
                console.log(res);
                if (res === "success" || res === "success with warning") {
                    localStorage.setItem("isTemplateSetEdited", true);
                    deletedFiles = [];
                    $("#newFiles").val('');
                    EditSet(); 
                    if (res === "success with warning") { app.showNotification('Error', 'Dynamic templates are not allowed'); }
                }
                else {
                    app.showNotification('Error', res);
                }
            }).fail(function (status) {
                app.showNotification('Error', 'Could not communicate with the server.');
            }).always(function () {
                setTimeout(function () {
                    $('#btnSaveExistingSet').show();
                    $('#btnSavingExistingSet').hide();
                }, 2500);
            });
        }
        else
            app.showNotification('Error', "Fields cannot be empty.");

        UpdateExcelVersion();
    }

    function DeleteTemplateFile() {
        $('#btnSaveExistingSet').prop('disabled', true);

        var delFile = deletedFiles.join(',');

        $.ajax({
            url: '/api/Template/DeleteTemplateFile',
            type: 'post',
            data: JSON.stringify(parseInt(delFile)),
            contentType: 'application/json;charset=utf-8'
        }).done(function (res) {
            if (res === 'success') {
                localStorage.setItem("isTemplateSetEdited", true);
                deletedFiles = [];
                EditSet();
            } else {
                app.showNotification('Error', 'Could not communicate with the server.');
            }
        }).fail(function (status) {
            app.showNotification('Error', 'Could not communicate with the server.');
        }).always(function () {
            setTimeout(function () { $('#btnSaveExistingSet').prop('disabled', false); }, 250);
        });
    }

    function BindTemplateFiles(data) {
        //debugger;
        if (data.Files === null) {
            var table = '<table id="tableFiles" width="100%" class="table table-striped">'
                + '<thead><tr><th></th>Action<th>File Name</th><th>Mapped</th><th>Action</th></tr></thead >';
            table += '</table>';
            $('#divSetFiles').html(table);
            return;
        }
        var files = [];
        var table_body = '<table id="tableFiles" width="100%" class="table table-striped">'
            + '<thead><tr><th>Action</th><th>File Name</th><th>Mapped</th><th>Action</th></tr></thead >';
        for (var i = 0; i < data.Files.length; i++) {
            table_body += '<tr>';

            //table_body += '<td>';
            //table_body += "<a class='fileAutoMap' id=" + data.Files[i].TemplateFileId + "><span style='color: red;cursor:pointer;'> Auto Map </span></a>";
            //table_body += '</td>';
            table_body += '<td>';
            table_body += "<a class='fileRemoveMap' id=" + data.Files[i].TemplateFileId + "><span style='color: red;cursor:pointer;'> Remove Map </span></a>";
            table_body += '</td>';
            table_body += '<td>';
            table_body += " " + data.Files[i].FileName;
            table_body += '</td>';
            table_body += '<td>';
            table_body += " " + data.Files[i].MappedPercentage;
            table_body += '</td>';
            table_body += '<td>';
            table_body += "<a class='fileEdit' id=" + data.Files[i].TemplateFileId + "><span style='color: red;cursor:pointer;'> Edit </span></a>";
            table_body += '</td>';
            table_body += '<td>';
            table_body += "<a class='fileDel' id=" + data.Files[i].TemplateFileId + "><span style='color: red;cursor:pointer;'> Delete </span></a>";
            table_body += '</td>';
            table_body += '<td style="display: none;">';
            table_body += data.Files[i].IsXFA;
            table_body += '</td>';
            table_body += '<td style="display: none;">';
            table_body += data.Files[i].DynamicFieldsCount;
            table_body += '</td>';

            table_body += '</tr>';
            files.push([data.Files[i].FileName]);
        }
        table_body += '</table>';
        $('#divSetFiles').html(table_body);
    }

    function IsAdmin() {
        var UType = localStorage.getItem("UserType");
        if (UType.toUpperCase() === "A") {
            return true;
        }
        return false;
    }

    function IsSuperAdmin() {
        var UType = localStorage.getItem("UserType");
        if (UType.toUpperCase() === "S") {
            return true;
        }
        return false;
    }

    function CreateTemplate() {
        var templateName = $.trim($('#txtSetName').val());
        var description = $.trim($('#txtDescription').val());

        pdffiles = [];
        var formData = new FormData();
        var files = $("#files").get(0).files;

        var filesPath = $('#files').val();
        var paths = filesPath.split(",");

        if (files.length > 0) {
            for (var i = 0; i < files.length; i++) {
                var f = files[i];
                formData.append("UploadedFiles", f);
                for (var j = 0; j < paths.length; j++) {
                    if (paths[j].indexOf(f.name) !== -1) {
                        pdffiles.push(f.name + ";" + paths[j]);
                    }
                }
            }
        }

        if (pdffiles.length === 0) {
            app.showNotification('Error', "Please select atleast one pdf file.");
            return;
        }

        if (templateName !== "" && description !== "") {
            $('#btnSave').hide();
            $('#btnSaving').show();
            var pdfTemplate = {
                TemplateName: templateName,
                Description: description,
                CompanyId: parseInt(localStorage.getItem("CompanyID")),
                TeamId: teamId,
                TemplateFileZip: null,
                IsActive: true,
                CreatedOn: new Date(),
                CreatedBy: parseInt(localStorage.getItem("UserID")),
                ExcelZip: null,
                TemplateFile: pdffiles.join(','),
                TemplateFileFieldMapping: null,
                TemplateFolderId: parseInt(localStorage.getItem("FolderId").replace("F", "")),
                FolderName: localStorage.getItem("FolderName"),
                SubFolderName: $('#txtSubfolderName').val(),
                FileNamePart: $('#txtOuptputFileName').val(),
                ExcelVersion: "",
                IsDemo: localStorage.getItem("IsDemo")
            };
            formData.append('TemplateName', pdfTemplate.TemplateName);
            formData.append('Description', pdfTemplate.Description);
            formData.append('CompanyId', pdfTemplate.CompanyId);
            formData.append('TeamId', pdfTemplate.TeamId);
            formData.append('TemplateFileZip', pdfTemplate.TemplateFileZip);
            formData.append('IsActive', pdfTemplate.IsActive);
            formData.append('CreatedOn', pdfTemplate.CreatedOn);

            formData.append('ExcelZip', pdfTemplate.ExcelZip);
            formData.append('TemplateFile', pdfTemplate.TemplateFile);
            formData.append('TemplateFileFieldMapping', pdfTemplate.TemplateFileFieldMapping);
            formData.append('CreatedBy', pdfTemplate.CreatedBy);
            formData.append('TemplateFolderId', pdfTemplate.TemplateFolderId);
            formData.append('FolderName', pdfTemplate.FolderName);
            formData.append('SubFolderName', pdfTemplate.SubFolderName);
            formData.append('FileNamePart', pdfTemplate.FileNamePart);
            formData.append('IsDemo', pdfTemplate.IsDemo);

            Excel.run(function (context) {
                var xmlpart = context.workbook.customXmlParts.getByNamespace("exceltoforms").getOnlyItem();
                xmlpart.load("id");
                return context.sync()
                    .then(function () {
                        var partCount = context.workbook.customXmlParts.getCount();
                        return context.sync().then(function () {
                            if (partCount.value <= 0) {
                                var parts = context.workbook.customXmlParts;
                                parts.load();

                                return context.sync().then(function () {
                                    var xmlData = "<templates xmlns='exceltoforms'><guid>5a14ac11-731e-408a-aa2d-4c8335706e6a</guid></templates>";
                                    var part = parts.add(xmlData);
                                    part.load("id");
                                    return context.sync().then(function () {
                                        pdfTemplate.ExcelVersion = part.id;
                                        formData.append('ExcelVersion', pdfTemplate.ExcelVersion);

                                        Office.context.document.getFilePropertiesAsync(function (asyncResult) {
                                            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                                                var fileUrl = asyncResult.value.url;
                                                formData.append('FileExtension', fileUrl.substring(fileUrl.lastIndexOf('.')));
                                                getDocumentAsCompressed(formData);
                                                return context.sync();
                                            } else if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                                                app.showNotification('Error', "The excel file hasn't been saved yet. Save the file and try again");
                                            }
                                        });

                                    });
                                });
                            } else {
                                return context.sync().then(function () {
                                    $.ajax({
                                        url: "/api/Template/IsExcelVersionExist",
                                        type: 'Get',
                                        data: {
                                            excelVersion: xmlpart.id
                                        },
                                        contentType: 'application/json;charset=utf-8'
                                    }).done(function (res) {
                                        if (res.Error === null || res.Error === "") {
                                            if (res.IsExcelVersionExist) {
                                                app.showNotification('Message', 'The excel file belongs to an existing set, please use a new excel file for a new set. <a style="color:#2191f1" href="http://exceltoforms.com/pricing/" target="_blank">Learn More.</a>');
                                                return;
                                            }
                                            pdfTemplate.ExcelVersion = xmlpart.id;
                                            formData.append('ExcelVersion', pdfTemplate.ExcelVersion);

                                            Office.context.document.getFilePropertiesAsync(function (asyncResult) {
                                                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                                                    var fileUrl = asyncResult.value.url;
                                                    formData.append('FileExtension', fileUrl.substring(fileUrl.lastIndexOf('.')));
                                                    getDocumentAsCompressed(formData);
                                                    return context.sync();
                                                } else if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                                                    app.showNotification('Error', "The excel file hasn't been saved yet. Save the file and try again");
                                                }
                                            });
                                        } else {
                                            app.showNotification('Error', res.Error);
                                        }
                                    }).fail(function (status) {
                                        app.showNotification('Error', status.responseText);
                                    });
                                });
                            }
                        });
                    }).catch(function (error) {
                        if (error.message === "This operation is not permitted for the current object.") {
                            var parts = context.workbook.customXmlParts;
                            parts.load();

                            return context.sync().then(function () {
                                var xmlData = "<templates xmlns='exceltoforms'><guid>5a14ac11-731e-408a-aa2d-4c8335706e6a</guid></templates>";
                                var part = parts.add(xmlData);
                                part.load("id");
                                return context.sync().then(function () {
                                    pdfTemplate.ExcelVersion = part.id;
                                    formData.append('ExcelVersion', pdfTemplate.ExcelVersion);

                                    Office.context.document.getFilePropertiesAsync(function (asyncResult) {
                                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                                            var fileUrl = asyncResult.value.url;
                                            formData.append('FileExtension', fileUrl.substring(fileUrl.lastIndexOf('.')));
                                            getDocumentAsCompressed(formData);
                                            return context.sync();
                                        } else if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                                            app.showNotification('Error', "The excel file hasn't been saved yet. Save the file and try again");
                                        }
                                    });
                                });
                            });
                        } else {
                            app.showNotification('Error', error.message);
                        }
                    });
            });
        }
        else
            app.showNotification('Error', "Fields cannot be empty.");
    }

    function getSliceAsync(file, nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived, formData) {
        file.getSliceAsync(nextSlice, function (sliceResult) {
            if (sliceResult.status === "succeeded") {
                if (!gotAllSlices) {
                    return;
                }

                docdataSlices[sliceResult.value.index] = sliceResult.value.data;
                if (++slicesReceived === sliceCount) {
                    file.closeAsync();
                    onGotAllSlices(docdataSlices, formData);
                }
                else {
                    getSliceAsync(file, ++nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived, formData);
                }
            }
            else {
                $('#btnSave').show();
                $('#btnSaving').hide();
                gotAllSlices = false;
                file.closeAsync();
                app.showNotification("Error", "getSliceAsync Error:", sliceResult.error.message);
            }
        });
    }

    function getExcelVersionSliceAsync(file, nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived, excelVersion) {
        file.getSliceAsync(nextSlice, function (sliceResult) {
            if (sliceResult.status === "succeeded") {
                if (!gotAllSlices) {
                    return;
                }

                docdataSlices[sliceResult.value.index] = sliceResult.value.data;
                if (++slicesReceived === sliceCount) {
                    file.closeAsync();
                    onGotExcelVersionAllSlices(docdataSlices, excelVersion);
                }
                else {
                    getExcelVersionSliceAsync(file, ++nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived, excelVersion);
                }
            }
            else {
                gotAllSlices = false;
                file.closeAsync();
                app.showNotification("Error", "getSliceAsync Error:", sliceResult.error.message);
            }
        });
    }

    function onGotExcelVersionAllSlices(docdataSlices, excelVersion) {
        docdata = [];
        for (var i = 0; i < docdataSlices.length; i++) {
            docdata = docdata.concat(docdataSlices[i]);
        }

        if (docdata.length > 0) {
            excelVersion.fileBytes = docdata;
            $.ajax({
                type: "post",
                url: "/api/Template/UpdateExcelVersion",
                data: JSON.stringify(excelVersion),
                contentType: 'application/json;charset=utf-8'
            }).done(function (res) {
                if (res === "success") {
                    isExcelVersionUpdated = true;
                    isExcelVersioMatching = true;
                } else {
                    isExcelVersionUpdated = false;
                    isExcelVersioMatching = false;
                    app.showNotification('Error', 'Could not communicate with the server.');
                }
            }).fail(function (status) {
                app.showNotification('Error', 'Could not communicate with the server.');
            }).always(function () {
                //setTimeout(function () {
                //    $('#btnUpdateVersion').show();
                //    $('#btnUpdatingVersion').hide();
                //}, 250);
            });
        }
    }

    function onGotAllSlices(docdataSlices, formData) {
        docdata = [];
        for (var i = 0; i < docdataSlices.length; i++) {
            docdata = docdata.concat(docdataSlices[i]);
        }

        if (docdata.length > 0) {
            $.ajax({
                type: "post",
                url: "/api/Template/CreateTemplateSet",
                contentType: false,
                processData: false,
                data: formData
            }).done(function (res) {
                if (res.TemplateId > 0) {
                    var excelFile = {
                        TemplateId: res.TemplateId,
                        fileBytes: docdata
                    };
                    $.ajax({
                        url: '/api/Template/UploadExcelFile',
                        type: 'post',
                        data: JSON.stringify(excelFile),
                        contentType: 'application/json;charset=utf-8'
                    }).done(function (res) {
                        if (res === 'success') {
                            docdata = [];
                            $('#files').val('');
                            pdffiles = [];
                            localStorage.setItem("isSaved", true);
                            window.location.href = '../DashBoard/DashBoard.html';
                        } else {
                            app.showNotification('Error', res);
                        }
                    }).fail(function (status) {
                        app.showNotification('Error', status.responseText);
                    });
                } else {
                    app.showNotification('Error', res.Error);
                }
            }).fail(function (status) {
                app.showNotification('Error', status.responseText);
            }).always(function () {
                setTimeout(function () {
                    $('#btnSave').show();
                    $('#btnSaving').hide();
                }, 5000);
            });
        }
    }

    function UpdateExcelVersion() {
        console.log('Update');
        //$('#btnUpdateVersion').hide();
        //$('#btnUpdatingVersion').show();
        var userId = parseInt(localStorage.getItem("UserID"));
        var templateId = parseInt(localStorage.getItem("FolderId").replace("T", ""));

        Excel.run(function (context) {
            /*var originalXml = "<templates xmlns='http://schemas.contoso.com/review/1.0'><guid>5a14ac11-731e-408a-aa2d-4c8335706e6a</guid></templates>";
            var customXmlPart = context.workbook.customXmlParts.add(originalXml);
            customXmlPart.load("id");*/
            return context.sync()
                .then(function () {
                    /*var settings = context.workbook.settings;
                    settings.add("XmlPartId", customXmlPart.id);*/
                    Office.context.document.getFilePropertiesAsync(function (asyncResult) {
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            var fileUrl = asyncResult.value.url;

                            var excelVersion = {
                                TemplateId: templateId,
                                UserId: userId,
                                ExcelVersionId: '',
                                FileExtension: fileUrl.substring(fileUrl.lastIndexOf('.'))
                            };
                            getExcelVersionDocumentAsCompressed(excelVersion);
                        } else if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                            app.showNotification('Error', "The excel file hasn't been saved yet. Save the file and try again");
                        }
                    });
                }).catch(errorHandlerFunction);
        });
    }

    function LogoutBtn() {
        $("#logout").hide();
        $("#loggingout").show();
        localStorage.clear();
        setTimeout(function () {
            app.Signout();
        }, 250);
    }

    function PanelOpen(PageNameDisplay) {
        app.InfoDisplay(PageNameDisplay);
    }
})();
