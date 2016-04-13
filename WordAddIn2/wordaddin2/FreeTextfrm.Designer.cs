namespace WordAddIn2
{
    partial class FreeTextfrm
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
            this.text1 = new System.Windows.Forms.TextBox();
            this.btn_ok = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.text2 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
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
            this.label3 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.text3 = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.text4 = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // text1
            // 
            this.text1.Location = new System.Drawing.Point(59, 38);
            this.text1.Name = "text1";
            this.text1.Size = new System.Drawing.Size(278, 20);
            this.text1.TabIndex = 1;
            // 
            // btn_ok
            // 
            this.btn_ok.Location = new System.Drawing.Point(278, 171);
            this.btn_ok.Name = "btn_ok";
            this.btn_ok.Size = new System.Drawing.Size(59, 73);
            this.btn_ok.TabIndex = 9;
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
            this.label1.Text = "Please Enter Text";
            this.label1.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // text2
            // 
            this.text2.Location = new System.Drawing.Point(59, 66);
            this.text2.Name = "text2";
            this.text2.Size = new System.Drawing.Size(278, 20);
            this.text2.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(4, 39);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(40, 13);
            this.label2.TabIndex = 0;
            this.label2.Text = "Text 1:";
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.Location = new System.Drawing.Point(7, 153);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(97, 17);
            this.radioButton1.TabIndex = 5;
            this.radioButton1.Text = "Update All Doc";
            this.radioButton1.UseVisualStyleBackColor = true;
            this.radioButton1.CheckedChanged += new System.EventHandler(this.radioButton1_CheckedChanged);
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.Location = new System.Drawing.Point(6, 176);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(95, 17);
            this.radioButton2.TabIndex = 6;
            this.radioButton2.Text = "Specific pages";
            this.radioButton2.UseVisualStyleBackColor = true;
            this.radioButton2.CheckedChanged += new System.EventHandler(this.radioButton2_CheckedChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(8, 257);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(73, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "From Chapter:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(8, 289);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(63, 13);
            this.label5.TabIndex = 9;
            this.label5.Text = "To Chapter:";
            // 
            // radioButton3
            // 
            this.radioButton3.AutoSize = true;
            this.radioButton3.Location = new System.Drawing.Point(6, 199);
            this.radioButton3.Name = "radioButton3";
            this.radioButton3.Size = new System.Drawing.Size(108, 17);
            this.radioButton3.TabIndex = 7;
            this.radioButton3.Text = "Specific Chapters";
            this.radioButton3.UseVisualStyleBackColor = true;
            this.radioButton3.CheckedChanged += new System.EventHandler(this.radioButton3_CheckedChanged);
            // 
            // FromPagebox
            // 
            this.FromPagebox.FormattingEnabled = true;
            this.FromPagebox.Location = new System.Drawing.Point(113, 254);
            this.FromPagebox.Name = "FromPagebox";
            this.FromPagebox.Size = new System.Drawing.Size(86, 21);
            this.FromPagebox.TabIndex = 20;
            // 
            // ToPagebox
            // 
            this.ToPagebox.FormattingEnabled = true;
            this.ToPagebox.Location = new System.Drawing.Point(114, 286);
            this.ToPagebox.Name = "ToPagebox";
            this.ToPagebox.Size = new System.Drawing.Size(86, 21);
            this.ToPagebox.TabIndex = 21;
            // 
            // ToChapterbox
            // 
            this.ToChapterbox.FormattingEnabled = true;
            this.ToChapterbox.Location = new System.Drawing.Point(114, 286);
            this.ToChapterbox.Name = "ToChapterbox";
            this.ToChapterbox.Size = new System.Drawing.Size(86, 21);
            this.ToChapterbox.TabIndex = 11;
            // 
            // FromChapterbox
            // 
            this.FromChapterbox.FormattingEnabled = true;
            this.FromChapterbox.Location = new System.Drawing.Point(113, 254);
            this.FromChapterbox.Name = "FromChapterbox";
            this.FromChapterbox.Size = new System.Drawing.Size(86, 21);
            this.FromChapterbox.TabIndex = 10;
            // 
            // radioButton4
            // 
            this.radioButton4.AutoSize = true;
            this.radioButton4.Location = new System.Drawing.Point(6, 223);
            this.radioButton4.Name = "radioButton4";
            this.radioButton4.Size = new System.Drawing.Size(80, 17);
            this.radioButton4.TabIndex = 8;
            this.radioButton4.Text = "New Pages";
            this.radioButton4.UseVisualStyleBackColor = true;
            this.radioButton4.CheckedChanged += new System.EventHandler(this.radioButton4_CheckedChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(4, 66);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(40, 13);
            this.label3.TabIndex = 25;
            this.label3.Text = "Text 2:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(4, 94);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(40, 13);
            this.label6.TabIndex = 27;
            this.label6.Text = "Text 3:";
            // 
            // text3
            // 
            this.text3.Location = new System.Drawing.Point(59, 94);
            this.text3.Name = "text3";
            this.text3.Size = new System.Drawing.Size(278, 20);
            this.text3.TabIndex = 3;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(4, 122);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(40, 13);
            this.label7.TabIndex = 29;
            this.label7.Text = "Text 4:";
            // 
            // text4
            // 
            this.text4.Location = new System.Drawing.Point(59, 122);
            this.text4.Name = "text4";
            this.text4.Size = new System.Drawing.Size(278, 20);
            this.text4.TabIndex = 4;
            // 
            // FreeTextfrm
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(349, 319);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.text4);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.text3);
            this.Controls.Add(this.label3);
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
            this.Controls.Add(this.label2);
            this.Controls.Add(this.text2);
            this.Controls.Add(this.btn_ok);
            this.Controls.Add(this.text1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FreeTextfrm";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "Free Text Options";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.PageRevisionFrm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_ok;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        public System.Windows.Forms.TextBox text1;
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
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label6;
        public System.Windows.Forms.TextBox text3;
        private System.Windows.Forms.Label label7;
        public System.Windows.Forms.TextBox text4;
        public System.Windows.Forms.TextBox text2;
    }
}