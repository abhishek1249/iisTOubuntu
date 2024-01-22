using ExFormOfficeAddInBAL;
using ExFormOfficeAddInEntities;
using System;
using System.Collections.Generic;
using System.Web;
using System.Web.Http;

namespace ExFormOfficeAddInExcelUIWeb.Controllers
{
    public class LoginController : ApiController
    {
        public class CLSResponse
        {
            public string Status { get; set; }
            public string Message { get; set; }
            public string UserID { get; set; }
            public string CompanyID { get; set; }
            public string CompanyName { get; set; }
            public char UserType { get; set; }
        }

        [HttpGet()]
        public CLSResponse sam()
        {
            return new CLSResponse()
            {
                Status = string.Empty,
                Message = string.Empty
            };
        }

        [HttpGet()]
        public CLSResponse LoginProcess(string UName, string Upassword, string Uaccount)
        {
            try
            {
                var result = ValidateUser(UName, Upassword, Uaccount);
                if (result == "")
                {
                    var obju = Helper.LogIn(UName, Upassword, Uaccount);
                    if (obju != null)
                    {
                        return new CLSResponse()
                        {
                            Status = "Success!",
                            Message = "Success!",
                            CompanyID = Convert.ToString(obju.CompanyId),
                            UserID = Convert.ToString(obju.UserId),
                            UserType = obju.UserType,
                            CompanyName = obju.CompanyName
                        };
                    }
                    else
                    {
                        return new CLSResponse()
                        {
                            Status = "Authentication failed.",
                            Message = "Please Enter Valid Login Credentials."
                        };
                    }
                }
                else
                {
                    return new CLSResponse()
                    {
                        Status = "Authentication failed.",
                        Message = result
                    };
                }
            }
            catch (Exception ex)
            {
                return new CLSResponse()
                {
                    Status = ex.Message,
                    Message = ex.Message
                };
            }
        }

        [HttpGet()]
        [Route("api/GetAccountByUserId")]
        public List<KeyValuePair> GetAccountByUserId(int UserId)
        {
            try
            {
                var result = Helper.GetAccountByUserId(UserId);
                return result;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        [HttpGet()]
        [Route("api/GetTeamByAccount")]
        public List<KeyValuePair> GetTeamByAccount(int UserId, int CompanyId)
        {
            try
            {
                var result = Helper.GetTeamByAccount(UserId, CompanyId);
                return result;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        private string ValidateUser(string UName, string Upassword, string Uaccount)
        {
            var userName = UName.Trim();
            var password = Upassword.Trim();
            var account = Uaccount.Trim();
            var smslabel = "";

            if (string.IsNullOrWhiteSpace(userName))
            {
                smslabel = "Please enter user name.";
            }
            else if (userName.Length > 150)
            {
                smslabel = "User name should not exceed 150 characters.";
            }
            else if (string.IsNullOrWhiteSpace(password))
            {
                smslabel = "Please enter password.";
            }
            else if (password.Length > 50)
            {
                smslabel = "Password should not exceed 50 characters.";
            }
            else if (string.IsNullOrWhiteSpace(account))
            {
                smslabel = "Please enter user account.";
            }
            else if (account.Length > 50)
            {
                smslabel = "Account name should not exceed 150 characters.";
            }
            return smslabel;
        }
    }
}
