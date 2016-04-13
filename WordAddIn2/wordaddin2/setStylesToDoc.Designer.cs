namespace WordAddIn2
{
    partial class setStylesToDoc
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
            this.btn_ok = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.H1 = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.H2 = new System.Windows.Forms.ComboBox();
            this.I1 = new System.Windows.Forms.ComboBox();
            this.I2 = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // btn_ok
            // 
            this.btn_ok.Location = new System.Drawing.Point(191, 148);
            this.btn_ok.Name = "btn_ok";
            this.btn_ok.Size = new System.Drawing.Size(59, 25);
            this.btn_ok.TabIndex = 9;
            this.btn_ok.Text = "Done";
            this.btn_ok.UseVisualStyleBackColor = true;
            this.btn_ok.Click += new System.EventHandler(this.btn_ok_Click);
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.label1.Location = new System.Drawing.Point(38, -10);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(177, 33);
            this.label1.TabIndex = 0;
            this.label1.Text = "Select Main Headings";
            this.label1.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(4, 39);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(59, 13);
            this.label2.TabIndex = 0;
            this.label2.Text = "Heading 1:";
            // 
            // H1
            // 
            this.H1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.H1.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.H1.FormattingEnabled = true;
            this.H1.Location = new System.Drawing.Point(88, 36);
            this.H1.Name = "H1";
            this.H1.Size = new System.Drawing.Size(162, 21);
            this.H1.TabIndex = 10;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(4, 66);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(59, 13);
            this.label3.TabIndex = 25;
            this.label3.Text = "Heading 2:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(4, 94);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(75, 13);
            this.label6.TabIndex = 27;
            this.label6.Text = "Introduction 1‎:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(4, 122);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(75, 13);
            this.label7.TabIndex = 29;
            this.label7.Text = "Introduction‎ 2:";
            // 
            // H2
            // 
            this.H2.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.H2.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.H2.FormattingEnabled = true;
            this.H2.Location = new System.Drawing.Point(88, 63);
            this.H2.Name = "H2";
            this.H2.Size = new System.Drawing.Size(162, 21);
            this.H2.TabIndex = 30;
            // 
            // I1
            // 
            this.I1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.I1.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.I1.FormattingEnabled = true;
            this.I1.Location = new System.Drawing.Point(88, 90);
            this.I1.Name = "I1";
            this.I1.Size = new System.Drawing.Size(162, 21);
            this.I1.TabIndex = 31;
            // 
            // I2
            // 
            this.I2.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.I2.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.I2.FormattingEnabled = true;
            this.I2.Location = new System.Drawing.Point(88, 118);
            this.I2.Name = "I2";
            this.I2.Size = new System.Drawing.Size(162, 21);
            this.I2.TabIndex = 32;
            // 
            // setStylesToDoc
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(261, 179);
            this.Controls.Add(this.I2);
            this.Controls.Add(this.I1);
            this.Controls.Add(this.H2);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.H1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btn_ok);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "setStylesToDoc";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Doc Main Headings";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.setStylesToDoc_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_ok;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox H1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox H2;
        private System.Windows.Forms.ComboBox I1;
        private System.Windows.Forms.ComboBox I2;
    }
}