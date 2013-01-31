using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointFacebookAPI
{
    public partial class ThisAddIn
    {
        public const string _appkey = "444619572275296";
        public const string _appsecret = "6254b75eed00ad4edb403a3fd132fc75";
        private static string _userid = "";
        private static string _access_token = "";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public static void SetAccessToken(string a)
        {
            _access_token = a;
        }

        public static string GetAccessToken()
        {
            return _access_token;
        }

        public static void SetUserID(string id)
        {
            _userid = id;
        }

        public static string GetUserID()
        {
            return _userid;
        }

        #region VSTO 產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器
        /// 修改這個方法的內容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
