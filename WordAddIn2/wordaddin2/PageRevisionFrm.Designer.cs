namespace WordAddIn2
{
    partial class PageRevisionFrm
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
            this.page_rev = new System.Windows.Forms.TextBox();
            this.btn_ok = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.rev_date = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            this.radioButton2 = new System.Windows.Forms.RadioButton();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.radioButton3 = new System.Windows.Forms.RadioButton();
            this.FromPagebox = new System.Windows.Forms.ComboBox();
            this.ToPagebox = new System.Windows.Forms.ComboBox();
            this.ToChapterbox = new System.Windows.Forms.ComboBox();
            this.FromChapterbox = new System.Windows.Forms.ComboBox();
            this.radioButton4 = new System.Windows.Forms.RadioButton();
            this.SuspendLayout();
            // 
            // page_rev
            // 
            this.page_rev.Location = new System.Drawing.Point(7, 50);
            this.page_rev.Name = "page_rev";
            this.page_rev.Size = new System.Drawing.Size(143, 22);
            this.page_rev.TabIndex = 1;
            // 
            // btn_ok
            // 
            this.btn_ok.Location = new System.Drawing.Point(163, 49);
            this.btn_ok.Name = "btn_ok";
            this.btn_ok.Size = new System.Drawing.Size(59, 73);
            this.btn_ok.TabIndex = 3;
            this.btn_ok.Text = "OK";
            this.btn_ok.UseVisualStyleBackColor = true;
            this.btn_ok.Click += new System.EventHandler(this.btn_ok_Click);
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.label1.Location = new System.Drawing.Point(-3, -6);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(177, 33);
            this.label1.TabIndex = 0;
            this.label1.Text = "Please Enter Page";
            this.label1.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // rev_date
            // 
            this.rev_date.Location = new System.Drawing.Point(7, 100);
            this.rev_date.Name = "rev_date";
            this.rev_date.Size = new System.Drawing.Size(143, 22);
            this.rev_date.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(4, 27);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(66, 17);
            this.label2.TabIndex = 0;
            this.label2.Text = "Revision:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(4, 78);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(42, 17);
            this.label3.TabIndex = 0;
            this.label3.Text = "Date:";
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.Location = new System.Drawing.Point(7, 130);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(123, 21);
            this.radioButton1.TabIndex = 5;
            this.radioButton1.Text = "Update All Doc";
            this.radioButton1.UseVisualStyleBackColor = true;
            this.radioButton1.CheckedChanged += new System.EventHandler(this.radioButton1_CheckedChanged);
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.Location = new System.Drawing.Point(6, 153);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(121, 21);
            this.radioButton2.TabIndex = 6;
            this.radioButton2.Text = "Specific pages";
            this.radioButton2.UseVisualStyleBackColor = true;
            this.radioButton2.CheckedChanged += new System.EventHandler(this.radioButton2_CheckedChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(8, 234);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(98, 17);
            this.label4.TabIndex = 7;
            this.label4.Text = "From Chapter:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(8, 266);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(83, 17);
            this.label5.TabIndex = 9;
            this.label5.Text = "To Chapter:";
            // 
            // radioButton3
            // 
            this.radioButton3.AutoSize = true;
            this.radioButton3.Location = new System.Drawing.Point(6, 176);
            this.radioButton3.Name = "radioButton3";
            this.radioButton3.Size = new System.Drawing.Size(139, 21);
            this.radioButton3.TabIndex = 19;
            this.radioButton3.Text = "Specific Chapters";
            this.radioButton3.UseVisualStyleBackColor = true;
            this.radioButton3.CheckedChanged += new System.EventHandler(this.radioButton3_CheckedChanged);
            // 
            // FromPagebox
            // 
            this.FromPagebox.FormattingEnabled = true;
            this.FromPagebox.Location = new System.Drawing.Point(113, 231);
            this.FromPagebox.Name = "FromPagebox";
            this.FromPagebox.Size = new System.Drawing.Size(86, 24);
            this.FromPagebox.TabIndex = 20;
            // 
            // ToPagebox
            // 
            this.ToPagebox.FormattingEnabled = true;
            this.ToPagebox.Location = new System.Drawing.Point(114, 263);
            this.ToPagebox.Name = "ToPagebox";
            this.ToPagebox.Size = new System.Drawing.Size(86, 24);
            this.ToPagebox.TabIndex = 21;
            // 
            // ToChapterbox
            // 
            this.ToChapterbox.FormattingEnabled = true;
            this.ToChapterbox.Location = new System.Drawing.Point(114, 263);
            this.ToChapterbox.Name = "ToChapterbox";
            this.ToChapterbox.Size = new System.Drawing.Size(86, 24);
            this.ToChapterbox.TabIndex = 23;
            // 
            // FromChapterbox
            // 
            this.FromChapterbox.FormattingEnabled = true;
            this.FromChapterbox.Location = new System.Drawing.Point(113, 231);
            this.FromChapterbox.Name = "FromChapterbox";
            this.FromChapterbox.Size = new System.Drawing.Size(86, 24);
            this.FromChapterbox.TabIndex = 22;
            // 
            // radioButton4
            // 
            this.radioButton4.AutoSize = true;
            this.radioButton4.Location = new System.Drawing.Point(6, 200);
            this.radioButton4.Name = "radioButton4";
            this.radioButton4.Size = new System.Drawing.Size(100, 21);
            this.radioButton4.TabIndex = 24;
            this.radioButton4.Text = "New Pages";
            this.radioButton4.UseVisualStyleBackColor = true;
            this.radioButton4.CheckedChanged += new System.EventHandler(this.radioButton4_CheckedChanged);
            // 
            // PageRevisionFrm
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(229, 290);
            this.Controls.Add(this.radioButton4);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.ToChapterbox);
            this.Controls.Add(this.FromChapterbox);
            this.Controls.Add(this.ToPagebox);
            this.Controls.Add(this.FromPagebox);
            this.Controls.Add(this.radioButton3);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.radioButton2);
            this.Controls.Add(this.radioButton1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.rev_date);
            this.Controls.Add(this.btn_ok);
            this.Controls.Add(this.page_rev);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "PageRevisionFrm";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "Page Revision Options";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.PageRevisionFrm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_ok;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        public System.Windows.Forms.TextBox page_rev;
        public System.Windows.Forms.TextBox rev_date;
        private System.Windows.Forms.RadioButton radioButton1;
        private System.Windows.Forms.RadioButton radioButton2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.RadioButton radioButton3;
        private System.Windows.Forms.ComboBox FromPagebox;
        private System.Windows.Forms.ComboBox ToPagebox;
        private System.Windows.Forms.ComboBox ToChapterbox;
        private System.Windows.Forms.ComboBox FromChapterbox;
        private System.Windows.Forms.RadioButton radioButton4;
    }
}