/* Common app functionality */
/// <reference path="Scripts/FabricUI/Panel.js" />

var app = (function () {
    'use strict';

    var app = {};

    app.bindingID = 'myBinding';

    // Common initialization function (to be called from each page).
    app.initialize = function () {
        $('body').append(
            '<div id="notification-message">' +
            '<div class="padding">' +
            '<div id="notification-message-close"></div>' +
            '<div id="notification-message-header"></div>' +
            '<div id="notification-message-body"></div>' +
            '</div>' +
            '</div>' +
            '<div class="ms-Panel">' +
            '<div class="ms-Overlay ms-PanelAction-close"></div>' +
            '    <div class="ms-Panel-main">' +
            '        <div class="ms-Panel-commands">' +
            '            <button class="ms-Panel-closeButton ms-PanelAction-close">' +
            '                <i class="ms-Panel-closeIcon ms-Icon ms-Icon--x"></i>' +
            '            </button>                                                         ' +
            '        </div>                                                                ' +
            '        <div class="ms-Panel-contentInner">                                   ' +
            '           <p class="ms-Panel-headerText" id="InfoPanelHeader"></p>          ' +
            '          <div class="ms-Panel-content">                                    ' +
            '             <span class="ms-font-m" id="InfoPanelBody">                   ' +
            '            </span>' +
            '       </div>' +
            '  </div>' +
            '</div>' +
            '</div>');



        $('#notification-message-close').click(function () {
            $('#notification-message').hide();
        });

        $('.ms-Panel').Panel();

        // After initialization, expose a common notification function.
        app.showNotification = function (header, text) {
            $('#Login-button').prop('disabled', false);
            $('#Save').prop('disabled', false);
            $('#btnSaveExistingSet').prop('disabled', false);
            $('#btnUpdateVersion').prop('disabled', false);
            $('#btnSave').show();
            $('#btnSaving').hide();
            $('#btnUpdateVersion').show();
            $('#btnUpdatingVersion').hide();
            $('#btnSaveExistingSet').show();
            $('#btnSavingExistingSet').hide();
            $('#btnAutomapFields').show();
            $('#btnAutomappingFields').hide();
            $('#btnRemoveMappings').show();
            $('#btnRemovingMappings').hide();
            $('#btnRemoveFileMapping').show();
            $('#btnRemovingFileMapping').hide();
            $('#btnDownloadExcel').show();
            $('#btnDownloadingExcel').hide();
            $('#btnEditSet').show();
            $('#btnEditingSet').hide();

            $('#notification-message-header').text(header);
            $('#notification-message-body').html(text);
            $('#notification-message').slideDown('slow');
            setTimeout(function () { $('#notification-message').hide(); }, 10000);

            //$('#dialog-notification').dialog('option', 'title', header);
            //$('#dialog-notification').html('<p><span class="ui-icon ui-icon-alert" style="float:left; margin:5px 12px 20px 0;"></span>'+text+'</p>');
            //$("#dialog-notification").dialog("open");
            //setTimeout(function () { $('#dialog-notification').dialog("close"); }, 10000);
        };
    };

    app.Signout = function () {
        Office.context.document.settings.remove('UserName');
        window.location.href = '../Login.html';
    };

    app.InfoDisplay = function (PageName) {
        $('.ms-Panel').Panel();
        $('#InfoPanelHeader').text(PageName);
        $('#InfoPanelBody').text("Retested This add-in allows you to pull data from web api" +
            "and update ExForm.You can display the data and show some advanced info." + PageName);
    };

    return app;
})();