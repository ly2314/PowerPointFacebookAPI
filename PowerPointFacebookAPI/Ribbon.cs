using Facebook;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointFacebookAPI
{
    public partial class Ribbon
    {
        public readonly String[] PERMISSIONS = new String[] { "email", "publish_stream", "user_about_me", "publish_actions" };

        private FacebookClient _facebook = new FacebookClient();
        
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            if (ThisAddIn.GetAccessToken() == "")
            {
                try
                {
                    FacebookClient fc = new FacebookClient();
                    LoginForm l = new LoginForm(fc.GetLoginUrl(new
                    {
                        client_id = ThisAddIn._appkey,
                        response_type = "token",
                        display = "touch",
                        redirect_uri = "https://www.facebook.com/connect/login_success.html",
                        scope = String.Join(",", PERMISSIONS)
                    }).ToString(), this, true);
                    l.Show();
                }
                catch (Exception ex)
                {
                }
            }
            else
            {
                var fb = new FacebookClient(ThisAddIn.GetAccessToken());
                dynamic me = fb.Get("me");
                var url = me.link;
                Process.Start(url);
            }
        }

        private string HttpGetResponse(string url)
        {
            WebClient wc = new WebClient();
            Stream st = wc.OpenRead(url);
            StreamReader sr = new StreamReader(st);
            string res = sr.ReadToEnd();
            return res;
        }

        public void LoginDone()
        {
            button1.Enabled = false;
            button2.Enabled = true;
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            if (ThisAddIn.GetAccessToken() == "")
            {
                MessageBox.Show("請先登入！", "Facebook", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
                return;
            }
            string select = "";
            try
            {
                PowerPoint._Application myPPT = Globals.ThisAddIn.Application;
                var Sld = myPPT.ActiveWindow.View.Slide;
                select = myPPT.ActiveWindow.Selection.TextRange.Text;
            }
            catch (Exception ex)
            {
                MessageBox.Show("請選取文字！", "Facebook", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
                return;
            }
            if (select != "")
            {
                FacebookClient facebook = new FacebookClient(ThisAddIn.GetAccessToken()); // 使用 Token 建立一個 FacebookClient
                var parameters = new Dictionary<String, Object>();
                facebook.PostCompleted += OnFacebookPostCompleted;
                DateTime dt = DateTime.Now;
                parameters["message"] = select;
                facebook.PostAsync("me/feed", parameters);
            }
            else
            {
                MessageBox.Show("請選取文字！", "Facebook", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
        }

        private void OnFacebookPostCompleted(object sender, FacebookApiEventArgs e)
        {
            MessageBox.Show("成功張貼訊息！", "Facebook", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            FacebookClient fc = new FacebookClient();
            LoginForm l = new LoginForm(fc.GetLogoutUrl(new
            {
                next = "https://www.facebook.com/connect/login_success.html",
                access_token = ThisAddIn.GetAccessToken()
            }).ToString(), this, false);
            l.Show();
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            if (ThisAddIn.GetAccessToken() == "")
            {
                MessageBox.Show("請先登入！", "Facebook", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
                return;
            }
            var tempClipboard = Clipboard.GetDataObject();
            try
            {
                PowerPoint._Application myPPT = Globals.ThisAddIn.Application;
                PowerPoint.Slide Sld = myPPT.ActiveWindow.View.Slide;
                myPPT.ActiveWindow.Selection.Copy();
            }
            catch
            {
                MessageBox.Show("請選取圖片！", "Facebook", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
                return;
            }
            if (Clipboard.ContainsImage())
            {
                UploadImage up = new UploadImage(Clipboard.GetImage());
                Clipboard.Clear();
                Clipboard.SetDataObject(tempClipboard);
                up.Show();
            }
            else
            {
                MessageBox.Show("請選取圖片！", "Facebook", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            if (ThisAddIn.GetAccessToken() == "")
            {
                MessageBox.Show("請先登入！", "Facebook", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
                return;
            }
            PowerPoint._Application myPPT = Globals.ThisAddIn.Application;
            PowerPoint.Slide Sld = myPPT.ActiveWindow.View.Slide;
            DateTime dt = DateTime.Now;
            string filename = String.Format("{0:DD.HH.MM.ss}", dt);
            Sld.Export("D:\\" + filename + ".jpg", "JPG");
            FileStream stream = new FileStream("D:\\" + filename + ".jpg", FileMode.Open, FileAccess.Read);
            Image img = Image.FromStream(stream);
            stream.Close();
            UploadImage up = new UploadImage(img);
            File.Delete("D:\\" + filename + ".jpg");
            up.Show();
        }
    }
}
