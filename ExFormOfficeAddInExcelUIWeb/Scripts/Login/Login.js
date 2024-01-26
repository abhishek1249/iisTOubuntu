/// <reference path="/Scripts/FabricUI/message.banner.js" />

(function () {
    "use strict";
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function () {
        $(document).ready(function () {
            app.initialize();
            // Initialize the FabricUI notification mechanism and hide it
            //var element = document.querySelector('.ms-MessageBanner');
            //messageBanner = new fabric.MessageBanner(element);
            //messageBanner.hideBanner();
            $('#txtusername').val("catalin@exceltoforms.com");
            $('#txtpassword').val("ExcelDev");
            //$('#txtAccount').val("exforms");
            /*$('#txtusername').focus();*/
            $('#Login-button').click(LoginProcess);
            $("#dialog-notification").dialog({
                autoOpen: false,
                resizable: false,
                height: "auto",
                width: 290,
                modal: true,
                draggable: false,
                buttons: {
                    Close: function () {
                        $(this).dialog("close");
                    }
                }
            });
        });
        $('#btnBack').click(RedirectToLandingPage);
    };

    function LoginProcess() {
        Excel.run(function (ctx) {
            return ctx.sync()
                .then(function () {
                    var dataToPassToService = {
                        UserName: $.trim($('#txtusername').val()),
                        Password: $.trim($('#txtpassword').val()),
                        UserAccount: "" //$.trim($('#txtAccount').val())
                    };

                    if (validateUser(dataToPassToService.UserName, dataToPassToService.Password, dataToPassToService.UserAccount)) {
                        Office.context.document.settings.set('UserName', dataToPassToService.UserName);
                        $('#Login-button').hide();
                        $('#loging').show();
                        $.ajax({
                            url: 'api/login/LoginProcess',
                            type: 'Get',
                            data: {
                                UName: dataToPassToService.UserName,
                                Upassword: dataToPassToService.Password,
                                Uaccount: dataToPassToService.UserAccount
                            },
                            contentType: 'application/json;charset=utf-8'
                        }).done(function (data) {
                            if (data.Status === "Success!") {
                                console.log(data);
                                localStorage.setItem("CompanyID", data.CompanyID);
                                localStorage.setItem("UserID", data.UserID);
                                localStorage.setItem("UserName", data.UserName);
                                localStorage.setItem("FullName", data.FullName);
                                localStorage.setItem("CompanyName", data.CompanyName);
                                localStorage.setItem("UserType", data.UserType);
                                window.location.href = '../DashBoard/DashBoard.html';
                            }
                            else {
                                app.showNotification('Error', data.Message);
                            }
                        }).fail(function (status) {
                            app.showNotification('Error', 'Could not communicate with the server.');
                        }).always(function () {
                            setTimeout(function () { $('#loging').hide(); $('#Login-button').show(); }, 2500);
                        });
                    }
                })
                .then(ctx.sync);
        }).catch(errorHandler);

    }
    function validateUser(name, password, account) {
        if (name === '') {
            app.showNotification('Error', "Please enter user name.");
            $('#txtusername').focus();
            return false;
        }
        else if (name.length > 150) {
            app.showNotification('Error', "User name cannot exceed 150 characters.");
            $('#txtusername').focus();
            return false;
        }
        else if (password === '') {
            app.showNotification('Error', "Please enter password.");
            $('#txtpassword').focus();
            return false;
        }
        else if (password.length > 50) {
            app.showNotification('Error', "Password cannot exceed 50 characters.");
            $('#txtpassword').focus();
            return false;
        }
        //else if (account === '') {
        //    app.showNotification('Error', "Please enter account name.");
        //    $('#txtAccount').focus();
        //    return false;
        //}
        //else if (account.length > 150) {
        //    app.showNotification('Error', "Account name cannot exceed 150 characters.");
        //    $('#txtAccount').focus();
        //    return false;
        //}
        return true;
    }

    function errorHandler(error) {
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        $('#ms-MessageBanner').slideDown('slow');
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }

    function RedirectToLandingPage() {
        window.location.href = '../LandingPage/LandingPage.html';
    }
})();