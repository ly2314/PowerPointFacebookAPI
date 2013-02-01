using Facebook;
using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace PowerPointFacebookAPI
{
    public partial class UploadImage : Form
    {
        Image image;

        public UploadImage(Image i)
        {
            InitializeComponent();
            image = i;
            pictureBox1.Image = image;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                DateTime dt = DateTime.Now;
                string filename = String.Format("{0:HH.mm.ss}", dt) + ".jpg";
                var fb = new FacebookClient(ThisAddIn.GetAccessToken());
                fb.PostCompleted += fb_PostCompleted;
                fb.PostTaskAsync("me/photos",new
                    {
                        message = richTextBox1.Text,
                        file = new FacebookMediaObject
                        {
                            ContentType = "image/jpeg",
                            FileName = filename
                        }.SetValue(ImageToBuffer(image, System.Drawing.Imaging.ImageFormat.Jpeg))
                    });
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        void fb_PostCompleted(object sender, FacebookApiEventArgs e)
        {
            MessageBox.Show("圖片已經上傳。", "Facebook", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
        }

        private byte[] ImageToBuffer(Image Image, System.Drawing.Imaging.ImageFormat imageFormat)
        {
            if (Image == null) { return null; }
            byte[] data = null;
            using (MemoryStream oMemoryStream = new MemoryStream())
            {
                using (Bitmap oBitmap = new Bitmap(Image))
                {
                    oBitmap.Save(oMemoryStream, imageFormat);
                    oMemoryStream.Position = 0;
                    data = new byte[oMemoryStream.Length];
                    oMemoryStream.Read(data, 0, Convert.ToInt32(oMemoryStream.Length));
                    oMemoryStream.Flush();
                }
            }
            return data;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
