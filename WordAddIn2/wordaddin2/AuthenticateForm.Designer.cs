namespace eDocs_Editor
{
    partial class AuthenticateForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.company_name = new System.Windows.Forms.TextBox();
            this.contact_name = new System.Windows.Forms.TextBox();
            this.email = new System.Windows.Forms.TextBox();
            this.addin_license = new System.Windows.Forms.TextBox();
            this.out_put_text = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(74, 165);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(115, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Authenticate Code";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(85, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Company Name:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(15, 48);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(78, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Contact Name:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(15, 75);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(35, 13);
            this.label3.TabIndex = 3;
            this.label3.Text = "Email:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(15, 101);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(47, 13);
            this.label4.TabIndex = 4;
            this.label4.Text = "License:";
            // 
            // company_name
            // 
            this.company_name.Location = new System.Drawing.Point(107, 24);
            this.company_name.Name = "company_name";
            this.company_name.Size = new System.Drawing.Size(145, 20);
            this.company_name.TabIndex = 5;
            // 
            // contact_name
            // 
            this.contact_name.Location = new System.Drawing.Point(107, 50);
            this.contact_name.Name = "contact_name";
            this.contact_name.Size = new System.Drawing.Size(145, 20);
            this.contact_name.TabIndex = 6;
            // 
            // email
            // 
            this.email.Location = new System.Drawing.Point(107, 76);
            this.email.Name = "email";
            this.email.Size = new System.Drawing.Size(145, 20);
            this.email.TabIndex = 7;
            // 
            // addin_license
            // 
            this.addin_license.Location = new System.Drawing.Point(107, 102);
            this.addin_license.Name = "addin_license";
            this.addin_license.Size = new System.Drawing.Size(145, 20);
            this.addin_license.TabIndex = 8;
            // 
            // out_put_text
            // 
            this.out_put_text.AutoSize = true;
            this.out_put_text.Location = new System.Drawing.Point(104, 128);
            this.out_put_text.Name = "out_put_text";
            this.out_put_text.Size = new System.Drawing.Size(108, 13);
            this.out_put_text.TabIndex = 9;
            this.out_put_text.Text = "Need to Authenticate";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(15, 128);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(40, 13);
            this.label5.TabIndex = 10;
            this.label5.Text = "Status:";
            // 
            // AuthenticateForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(264, 200);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.out_put_text);
            this.Controls.Add(this.addin_license);
            this.Controls.Add(this.email);
            this.Controls.Add(this.contact_name);
            this.Controls.Add(this.company_name);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Name = "AuthenticateForm";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Authenticate addIn";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.AuthenticateForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox company_name;
        private System.Windows.Forms.TextBox contact_name;
        private System.Windows.Forms.TextBox email;
        private System.Windows.Forms.TextBox addin_license;
        private System.Windows.Forms.Label out_put_text;
        private System.Windows.Forms.Label label5;
    }
}