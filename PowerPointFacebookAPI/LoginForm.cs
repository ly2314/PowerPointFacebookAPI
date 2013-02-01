using Facebook;
using System;
using System.Drawing;
using System.Net;
using System.Windows.Forms;

namespace PowerPointFacebookAPI
{
    public partial class LoginForm : Form
    {
        private Ribbon parent;
        private bool _login; 

        public LoginForm(string url, Ribbon _parent, bool login)
        {
            InitializeComponent();
            parent = _parent;
            _login = login;
            webBrowser1.Navigated += webBrowser1_Navigated;
            webBrowser1.Url = new Uri(url, UriKind.Absolute);
        }

        private void webBrowser1_Navigated(object sender, WebBrowserNavigatedEventArgs e)
        {
            if (webBrowser1.Url.AbsolutePath == "/connect/login_success.html" && _login)
            {
                string url = webBrowser1.Url.ToString();
                string access_token = url.Remove(0, "https://www.facebook.com/connect/login_success.html#".Length);
                access_token = access_token.Remove(access_token.IndexOf("&"));
                access_token = access_token.Remove(0, "access_token=".Length);
                ThisAddIn.SetAccessToken(access_token);
                var fb = new FacebookClient(access_token);
                dynamic me = fb.Get("me", new { fields = new[] { "id", "name", "picture" } });
                var id = me.id;
                var name = me.name;
                var icon = me.picture;
                icon = icon.data.url;
                ThisAddIn.SetUserID(id);
                parent.button1.Label = name;
                WebRequest wr = WebRequest.Create(icon);
　　            WebResponse res = wr.GetResponse();
　　            Bitmap bmp = new Bitmap(res.GetResponseStream());
                parent.button1.Image = new Bitmap(bmp);
                parent.button2.Enabled = true;
                parent.button1.SuperTip = "在facebook上的 " + name;
                this.Hide();
            }
            else if (webBrowser1.Url.AbsolutePath == "/connect/login_success.html" && !_login)
            {
                System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon));
                parent.button1.Image = global::PowerPointFacebookAPI.Properties.Resources.facebook;
                parent.button1.Label = "登入";
                parent.button1.SuperTip = "使用你的Facebook帳號登入";
                ThisAddIn.SetUserID("");
                ThisAddIn.SetAccessToken("");
                parent.button2.Enabled = false;
                this.Hide();           
            }
        }
    }
}
