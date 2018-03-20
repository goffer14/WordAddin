using System;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using eDocs_Editor.Properties;
using System.Diagnostics;
using Auth0.OidcClient;
using IdentityModel.OidcClient;
using System.Threading;
using System.Threading.Tasks;
using System.Text;
using System.Collections.Generic;
using System.Net;
using System.IO;
using System.Web;
using System.Web.Script.Serialization;
using Newtonsoft.Json.Linq;

namespace eDocs_Editor
{
    public partial class AuthenticateForm : Form
    {
        public bool isOnline=false;
        public string company_name_Text;
        public string contact_name_Text;
        public string email_Text;
        public string addin_license_Text;
        public AuthenticateForm()
        {
            InitializeComponent();
        }

        private void AuthenticateForm_Load(object sender, EventArgs e)
        {
            isOnline = settings.CheckForInternetConnection();
            if(!isOnline)
            {
                out_put_text.Text = "Need Internet Connection";
            }
            
        }
    public bool ValidData()
        {

        company_name_Text = company_name.Text;
        contact_name_Text = contact_name.Text;
        email_Text = email.Text;
        addin_license_Text = addin_license.Text;
            if (company_name_Text.Length < 1)
            {
                out_put_text.Text = "Text Error in Company Name";
                return false;
            }
            if (contact_name_Text.Length < 1)
            {
                out_put_text.Text = "Text Error in Contact Name";
                return false;
            }
            if (email_Text.Length < 1)
            {
                out_put_text.Text = "Text Error in Email";
                return false;
            }
            if (addin_license_Text.Length<1)
            {
                out_put_text.Text = "License Error";
                return false;
            }
            return true;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (button1.Text == "Close")
            {
                this.Close();
                this.Dispose();
            }
            if (!settings.CheckForInternetConnection())
            {
                out_put_text.Text = "Need Internet Connection";
                return;
            }
            if (ValidData())
                authUser();
        }
        public void authUser()
        {
            try
            {
                var request = (HttpWebRequest)HttpWebRequest.Create("https://global-edocs-auth.herokuapp.com/users/auth");
                request.Method = "POST";
                var stream = request.GetRequestStream();
                var bytes = Encoding.UTF8.GetBytes(getUserData());
                request.ContentType = "application/x-www-form-urlencoded;charset=UTF-8";
                stream.Write(bytes, 0, bytes.Length);
                stream.Close();
                var response = (HttpWebResponse)request.GetResponse();
                if (response.StatusCode == System.Net.HttpStatusCode.OK)
                {
                    var rawJson = new StreamReader(response.GetResponseStream()).ReadToEnd();
                    var json = JObject.Parse(rawJson);  //Turns your raw string into a key value lookup
                    string userId = json["_id"].ToObject<string>();
                    System.Diagnostics.Debug.WriteLine("userId - " + userId);
                    Settings.Default.userId = userId;
                    Settings.Default.is_active_auth = "true";
                    Settings.Default.Save();
                    MessageBox.Show("Success, you can now use the product");
                    this.Close();
                    this.Dispose();

                }
                else
                {
                    out_put_text.Text = "Error trying to Auth the user\r\nPlease contect support";
                    button1.Text = "Close";
                }
            }
            catch
            {
                out_put_text.Text = "Error trying to Auth the user\r\nPlease contect support";
                button1.Text = "Close";
            }
        }
        public string getUserData()
        {
            int i = 0;
            var sb = new StringBuilder();
            sb.AppendFormat("{0}={1}&", "email", HttpUtility.UrlEncode(email_Text));
            sb.AppendFormat("{0}={1}&", "password", HttpUtility.UrlEncode(addin_license_Text));
            sb.AppendFormat("{0}={1}&", "companyName", HttpUtility.UrlEncode(company_name_Text));
            sb.AppendFormat("{0}={1}&", "contactName", HttpUtility.UrlEncode(contact_name_Text));
            sb.Remove(sb.Length - 1, 1);
            return sb.ToString();
        }
    }
}
