using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;

namespace ExFormOfficeAddInExcelUIWeb.Models
{
    public static class SessionModel
    {
        public class CurrentUser
        {
            public static int UserId
            {
                get
                {
                    return Convert.ToInt32(HttpContext.Current.Session["UserId"]);
                }

                set
                {
                    HttpContext.Current.Session["UserId"] = value;
                }
            }

            public static string UserName
            {
                get
                {
                    return Convert.ToString(HttpContext.Current.Session["UserName"]);
                }

                set
                {
                    HttpContext.Current.Session["UserName"] = value;
                }
            }

            public static string UserType
            {
                get
                {
                    return Convert.ToString(HttpContext.Current.Session["UserType"]);
                }

                set
                {
                    HttpContext.Current.Session["UserType"] = value;
                }
            }

            public static string CompanyID
            {
                get
                {
                    return Convert.ToString(HttpContext.Current.Session["CompanyID"]);
                }

                set
                {
                    HttpContext.Current.Session["CompanyID"] = value;
                }
            }

            public static string UserEmail
            {
                get
                {
                    return Convert.ToString(HttpContext.Current.Session["UserEmail"]);
                }

                set
                {
                    HttpContext.Current.Session["UserEmail"] = value;
                }
            }

            public static string FirstName
            {
                get
                {
                    return Convert.ToString(HttpContext.Current.Session["FirstName"]);
                }

                set
                {
                    HttpContext.Current.Session["FirstName"] = value;
                }
            }

            public static string LastName
            {
                get
                {
                    return Convert.ToString(HttpContext.Current.Session["LastName"]);
                }

                set
                {
                    HttpContext.Current.Session["LastName"] = value;
                }
            }

            public static string FullName
            {
                get
                {
                    return Convert.ToString(HttpContext.Current.Session["FullName"]);
                }

                set
                {
                    HttpContext.Current.Session["FullName"] = value;
                }
            }

            public static int? RoleId
            {
                get
                {
                    return Convert.ToInt32(HttpContext.Current.Session["RoleId"]);
                }

                set
                {
                    HttpContext.Current.Session["RoleId"] = value;
                }
            }

            public static string RoleCode
            {
                get
                {
                    return Convert.ToString(HttpContext.Current.Session["RoleCode"]);
                }

                set
                {
                    HttpContext.Current.Session["RoleCode"] = value;
                }
            }
            public static string RoleCodePermissionName
            {
                get
                {
                    return Convert.ToString(HttpContext.Current.Session["RoleCodePermissionName"]);
                }

                set
                {
                    HttpContext.Current.Session["RoleCodePermissionName"] = value;
                }
            }

            public static string ManagerExist
            {
                get
                {
                    return Convert.ToString(HttpContext.Current.Session["ManagerExist"]);
                }

                set
                {
                    HttpContext.Current.Session["ManagerExist"] = value;
                }
            }
            public static string LoginToken
            {
                get
                {
                    return Convert.ToString(HttpContext.Current.Session["LoginToken"]);
                }

                set
                {
                    HttpContext.Current.Session["LoginToken"] = value;
                }
            }
            public static int SessionTimeout
            {
                get
                {
                    return Convert.ToInt32(HttpContext.Current.Session["SessionTimeout"]);
                }

                set
                {
                    HttpContext.Current.Session["SessionTimeout"] = value;
                }
            }
            public static DateTime SessionWillExpireOn
            {
                get
                {
                    return Convert.ToDateTime(HttpContext.Current.Session["SessionWillExpireOn"]);
                }

                set
                {
                    HttpContext.Current.Session["SessionWillExpireOn"] = value;
                }
            }
            //[System.ComponentModel.DefaultValue()]
            //public static string ADServerExist
            //{
            //    get
            //    {
            //        return Convert.ToString(HttpContext.Current.Session["ADServerExist"]);
            //    }

            //    set
            //    {
            //        HttpContext.Current.Session["ADServerExist"] = value;
            //    }
            //}

            public static string ADServerExist { get; set; } = ConfigurationManager.AppSettings["ADServer"].ToString();

        }

    }
}