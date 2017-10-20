using System;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using eDocs_Editor.Properties;
using System.Diagnostics;
using IdentityModel.OidcClient;

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

            DateTime CurrentDate = DateTime.Now.Date;
            MyThreadStartMethod();
        }
        private  void MyThreadStartMethod()
        {
            if (button1.Text == "Close")
                this.Close();
            if (ValidData())
            {
                if (Settings.Default.is_active == "true")
                {
                    if (isOnline)
                        check_license();
                    else
                    {
                        isOnline = settings.CheckForInternetConnection();
                        if (isOnline)
                            check_license();
                        else
                        {
                            out_put_text.Text = "Need Internet Connection";
                            check_license();
                        }
                    }
                }
                else
                    out_put_text.Text = "Alreadגy Authenticate";
                button1.Text = "Close";
            }
        }
        private void DisplayResult(LoginResult loginResult)
        {
            DateTime CurrentDate = DateTime.Now.Date;
            // Display error
            if (loginResult.IsError)
            {
                Debug.WriteLine($"An error occurred during login: {loginResult.Error}");
                Settings.Default.is_active = "true";

            }
            else if (!loginResult.IsError)
            {

                Debug.WriteLine($"name: {loginResult.User.FindFirst(c => c.Type == "name")?.Value}");
                Debug.WriteLine($"email: {loginResult.User.FindFirst(c => c.Type == "email")?.Value}");
                var days_for_use = 180;

                Settings.Default.days_for_use = days_for_use;
                Settings.Default.addin_license = loginResult.User.FindFirst(c => c.Type == "name")?.Value;
                Settings.Default.StartTime = CurrentDate;
                Settings.Default.Last_connction = CurrentDate;
                Settings.Default.FirstUse = "false";
                Settings.Default.is_active = "true";
                Settings.Default.Save();
            }
        }
        public void check_license()
        {
            DateTime CurrentDate = DateTime.Now.Date;
            out_put_text.Text = "Conncting to eDocs servers...";
            string s;
            MySqlConnection mcon = new MySqlConnection(Settings.Default.serverString);
            MySqlCommand mcd;
            MySqlDataReader mdr;
            try
            {
                mcon.Open();
                s = "select * from license_table where license_key = '" + addin_license_Text + "'";
                mcd = new MySqlCommand(s, mcon);
                mdr = mcd.ExecuteReader();
                if (mdr.Read())
                {
                    int days_for_use = Int32.Parse(mdr.GetString("days_for_use"));
                    int is_active = mdr.GetInt32("is_active");
                    int license_id = mdr.GetInt32("license_id");
                    mcon.Close();
                    if (is_active != 0)
                    {
                        out_put_text.Text = "License Already Used";
                        return;
                    }
                    string str1 = "update license_table set company_name='" + company_name_Text + "', contact_name='" + contact_name_Text + "', email='" + email_Text + "', is_active='" + 1 + "', activation_date='" + CurrentDate.ToShortDateString() + "' where license_id='" + license_id + "'";
                    mcon.Open();
                    MySqlCommand update = new MySqlCommand(str1, mcon);
                    update.ExecuteNonQuery();
                    mcon.Close();

                    Settings.Default.days_for_use = days_for_use;
                    Settings.Default.addin_license = addin_license.Text;
                    Settings.Default.StartTime = CurrentDate;
                    Settings.Default.Last_connction = CurrentDate;
                    Settings.Default.FirstUse = "false";
                    Settings.Default.is_active = "true";
                    Settings.Default.Save();
                    out_put_text.Text = "Authenticate Complete";
                    button1.Text = "Close";
                    return;
                }
                else
                {
                    out_put_text.Text = "No License in Database";
                    return;
                }
            }
            catch
            {
                out_put_text.Text = "Error Connecting to Server";
                if (Settings.Default.FirstUse == "true" && addin_license_Text == "blDAWY52R0cNBoq")
                {
                    Settings.Default.days_for_use = 14;
                    Settings.Default.addin_license = addin_license_Text;
                    Settings.Default.StartTime = CurrentDate;
                    Settings.Default.Last_connction = CurrentDate;
                    Settings.Default.FirstUse = "false";
                    Settings.Default.is_active = "true";
                    Settings.Default.Save();
                    out_put_text.Text = "Authenticate Complete";
                    button1.Text = "Close";
                }

                return;
            }
        }
    }
}
