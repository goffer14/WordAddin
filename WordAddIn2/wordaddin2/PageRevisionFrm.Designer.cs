namespace eDocs_Editor
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
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.button1 = new System.Windows.Forms.Button();
            this.page_rev = new System.Windows.Forms.TextBox();
            this.date_rev = new System.Windows.Forms.TextBox();
            this.pageView = new System.Windows.Forms.ListView();
            this.Real_Page = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.HeadingPage = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Rev = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Date = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Issue = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Effective_Date = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.text1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.text2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.text3 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.text4 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.text2_text = new System.Windows.Forms.TextBox();
            this.text1_text = new System.Windows.Forms.TextBox();
            this.text4_text = new System.Windows.Forms.TextBox();
            this.text3_text = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.dateCheckBox = new System.Windows.Forms.CheckBox();
            this.revCheckBox = new System.Windows.Forms.CheckBox();
            this.text1CheckBox = new System.Windows.Forms.CheckBox();
            this.text4CheckBox = new System.Windows.Forms.CheckBox();
            this.text3CheckBox = new System.Windows.Forms.CheckBox();
            this.text2CheckBox = new System.Windows.Forms.CheckBox();
            this.button2 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.issueCheckBox = new System.Windows.Forms.CheckBox();
            this.effectiveCheckBox = new System.Windows.Forms.CheckBox();
            this.effectiveText = new System.Windows.Forms.TextBox();
            this.issueText = new System.Windows.Forms.TextBox();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.checkBox1.Location = new System.Drawing.Point(235, 271);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(70, 17);
            this.checkBox1.TabIndex = 0;
            this.checkBox1.Text = "Select All";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.Click += new System.EventHandler(this.checkBox1_Click);
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.button1.Location = new System.Drawing.Point(19, 401);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 1;
            this.button1.Text = "Save";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // page_rev
            // 
            this.page_rev.Dock = System.Windows.Forms.DockStyle.Top;
            this.page_rev.Enabled = false;
            this.page_rev.Location = new System.Drawing.Point(127, 3);
            this.page_rev.Name = "page_rev";
            this.page_rev.Size = new System.Drawing.Size(215, 20);
            this.page_rev.TabIndex = 4;
            // 
            // date_rev
            // 
            this.date_rev.Dock = System.Windows.Forms.DockStyle.Top;
            this.date_rev.Enabled = false;
            this.date_rev.Location = new System.Drawing.Point(127, 30);
            this.date_rev.Name = "date_rev";
            this.date_rev.Size = new System.Drawing.Size(215, 20);
            this.date_rev.TabIndex = 5;
            // 
            // pageView
            // 
            this.pageView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pageView.AutoArrange = false;
            this.pageView.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pageView.CheckBoxes = true;
            this.pageView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.Real_Page,
            this.HeadingPage,
            this.Rev,
            this.Date,
            this.Issue,
            this.Effective_Date,
            this.text1,
            this.text2,
            this.text3,
            this.text4});
            this.pageView.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.pageView.FullRowSelect = true;
            this.pageView.GridLines = true;
            this.pageView.Location = new System.Drawing.Point(378, 42);
            this.pageView.Name = "pageView";
            this.pageView.Size = new System.Drawing.Size(964, 383);
            this.pageView.TabIndex = 6;
            this.pageView.UseCompatibleStateImageBehavior = false;
            this.pageView.View = System.Windows.Forms.View.Details;
            this.pageView.SelectedIndexChanged += new System.EventHandler(this.pageView_SelectedIndexChanged);
            this.pageView.MouseClick += new System.Windows.Forms.MouseEventHandler(this.pageView_MouseClick);
            // 
            // Real_Page
            // 
            this.Real_Page.Text = "Page";
            this.Real_Page.Width = 74;
            // 
            // HeadingPage
            // 
            this.HeadingPage.Text = "eDoc Page";
            this.HeadingPage.Width = 116;
            // 
            // Rev
            // 
            this.Rev.Text = "Revision";
            this.Rev.Width = 79;
            // 
            // Date
            // 
            this.Date.Text = "Date";
            this.Date.Width = 110;
            // 
            // Issue
            // 
            this.Issue.Text = "Issue";
            this.Issue.Width = 75;
            // 
            // Effective_Date
            // 
            this.Effective_Date.Text = "Effective Date";
            this.Effective_Date.Width = 126;
            // 
            // text1
            // 
            this.text1.Text = "Text 1";
            this.text1.Width = 102;
            // 
            // text2
            // 
            this.text2.Text = "Text 2";
            this.text2.Width = 84;
            // 
            // text3
            // 
            this.text3.Text = "Text 3";
            this.text3.Width = 97;
            // 
            // text4
            // 
            this.text4.Text = "Text 4";
            this.text4.Width = 101;
            // 
            // text2_text
            // 
            this.text2_text.Dock = System.Windows.Forms.DockStyle.Top;
            this.text2_text.Enabled = false;
            this.text2_text.Location = new System.Drawing.Point(127, 138);
            this.text2_text.Name = "text2_text";
            this.text2_text.Size = new System.Drawing.Size(215, 20);
            this.text2_text.TabIndex = 10;
            // 
            // text1_text
            // 
            this.text1_text.Dock = System.Windows.Forms.DockStyle.Top;
            this.text1_text.Enabled = false;
            this.text1_text.Location = new System.Drawing.Point(127, 111);
            this.text1_text.Name = "text1_text";
            this.text1_text.Size = new System.Drawing.Size(215, 20);
            this.text1_text.TabIndex = 9;
            // 
            // text4_text
            // 
            this.text4_text.Dock = System.Windows.Forms.DockStyle.Top;
            this.text4_text.Enabled = false;
            this.text4_text.Location = new System.Drawing.Point(127, 192);
            this.text4_text.Name = "text4_text";
            this.text4_text.Size = new System.Drawing.Size(215, 20);
            this.text4_text.TabIndex = 14;
            // 
            // text3_text
            // 
            this.text3_text.Dock = System.Windows.Forms.DockStyle.Top;
            this.text3_text.Enabled = false;
            this.text3_text.Location = new System.Drawing.Point(127, 165);
            this.text3_text.Name = "text3_text";
            this.text3_text.Size = new System.Drawing.Size(215, 20);
            this.text3_text.TabIndex = 13;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.label7.Location = new System.Drawing.Point(13, 10);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(40, 16);
            this.label7.TabIndex = 15;
            this.label7.Text = "Use:";
            // 
            // dateCheckBox
            // 
            this.dateCheckBox.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.dateCheckBox.AutoSize = true;
            this.dateCheckBox.Location = new System.Drawing.Point(3, 32);
            this.dateCheckBox.Name = "dateCheckBox";
            this.dateCheckBox.Size = new System.Drawing.Size(52, 17);
            this.dateCheckBox.TabIndex = 16;
            this.dateCheckBox.Text = "Date:";
            this.dateCheckBox.UseVisualStyleBackColor = true;
            this.dateCheckBox.CheckedChanged += new System.EventHandler(this.dateCheckBox_CheckedChanged);
            // 
            // revCheckBox
            // 
            this.revCheckBox.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.revCheckBox.AutoSize = true;
            this.revCheckBox.Location = new System.Drawing.Point(3, 5);
            this.revCheckBox.Name = "revCheckBox";
            this.revCheckBox.Size = new System.Drawing.Size(70, 17);
            this.revCheckBox.TabIndex = 17;
            this.revCheckBox.Text = "Revision:";
            this.revCheckBox.UseVisualStyleBackColor = true;
            this.revCheckBox.CheckedChanged += new System.EventHandler(this.revCheckBox_CheckedChanged);
            // 
            // text1CheckBox
            // 
            this.text1CheckBox.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.text1CheckBox.AutoSize = true;
            this.text1CheckBox.Location = new System.Drawing.Point(3, 113);
            this.text1CheckBox.Name = "text1CheckBox";
            this.text1CheckBox.Size = new System.Drawing.Size(59, 17);
            this.text1CheckBox.TabIndex = 18;
            this.text1CheckBox.Text = "Text 1:";
            this.text1CheckBox.UseVisualStyleBackColor = true;
            this.text1CheckBox.CheckedChanged += new System.EventHandler(this.text1CheckBox_CheckedChanged);
            // 
            // text4CheckBox
            // 
            this.text4CheckBox.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.text4CheckBox.AutoSize = true;
            this.text4CheckBox.Location = new System.Drawing.Point(3, 194);
            this.text4CheckBox.Name = "text4CheckBox";
            this.text4CheckBox.Size = new System.Drawing.Size(59, 17);
            this.text4CheckBox.TabIndex = 19;
            this.text4CheckBox.Text = "Text 4:";
            this.text4CheckBox.UseVisualStyleBackColor = true;
            this.text4CheckBox.CheckedChanged += new System.EventHandler(this.text4CheckBox_CheckedChanged);
            // 
            // text3CheckBox
            // 
            this.text3CheckBox.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.text3CheckBox.AutoSize = true;
            this.text3CheckBox.Location = new System.Drawing.Point(3, 167);
            this.text3CheckBox.Name = "text3CheckBox";
            this.text3CheckBox.Size = new System.Drawing.Size(59, 17);
            this.text3CheckBox.TabIndex = 20;
            this.text3CheckBox.Text = "Text 3:";
            this.text3CheckBox.UseVisualStyleBackColor = true;
            this.text3CheckBox.CheckedChanged += new System.EventHandler(this.text3CheckBox_CheckedChanged);
            // 
            // text2CheckBox
            // 
            this.text2CheckBox.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.text2CheckBox.AutoSize = true;
            this.text2CheckBox.Location = new System.Drawing.Point(3, 140);
            this.text2CheckBox.Name = "text2CheckBox";
            this.text2CheckBox.Size = new System.Drawing.Size(59, 17);
            this.text2CheckBox.TabIndex = 21;
            this.text2CheckBox.Text = "Text 2:";
            this.text2CheckBox.UseVisualStyleBackColor = true;
            this.text2CheckBox.CheckedChanged += new System.EventHandler(this.text2CheckBox_CheckedChanged);
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.button2.Location = new System.Drawing.Point(100, 401);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 22;
            this.button2.Text = "Close";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(316, 272);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(45, 13);
            this.label1.TabIndex = 23;
            this.label1.Text = "Clear All";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // issueCheckBox
            // 
            this.issueCheckBox.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.issueCheckBox.AutoSize = true;
            this.issueCheckBox.Location = new System.Drawing.Point(3, 59);
            this.issueCheckBox.Name = "issueCheckBox";
            this.issueCheckBox.Size = new System.Drawing.Size(54, 17);
            this.issueCheckBox.TabIndex = 27;
            this.issueCheckBox.Text = "Issue:";
            this.issueCheckBox.UseVisualStyleBackColor = true;
            this.issueCheckBox.CheckedChanged += new System.EventHandler(this.issueCheckBox_CheckedChanged);
            // 
            // effectiveCheckBox
            // 
            this.effectiveCheckBox.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.effectiveCheckBox.AutoSize = true;
            this.effectiveCheckBox.Location = new System.Drawing.Point(3, 86);
            this.effectiveCheckBox.Name = "effectiveCheckBox";
            this.effectiveCheckBox.Size = new System.Drawing.Size(97, 17);
            this.effectiveCheckBox.TabIndex = 26;
            this.effectiveCheckBox.Text = "Effective Date:";
            this.effectiveCheckBox.UseVisualStyleBackColor = true;
            this.effectiveCheckBox.CheckedChanged += new System.EventHandler(this.effectiveCheckBox_CheckedChanged);
            // 
            // effectiveText
            // 
            this.effectiveText.Dock = System.Windows.Forms.DockStyle.Top;
            this.effectiveText.Enabled = false;
            this.effectiveText.Location = new System.Drawing.Point(127, 84);
            this.effectiveText.Name = "effectiveText";
            this.effectiveText.Size = new System.Drawing.Size(215, 20);
            this.effectiveText.TabIndex = 25;
            // 
            // issueText
            // 
            this.issueText.Dock = System.Windows.Forms.DockStyle.Top;
            this.issueText.Enabled = false;
            this.issueText.Location = new System.Drawing.Point(127, 57);
            this.issueText.Name = "issueText";
            this.issueText.Size = new System.Drawing.Size(215, 20);
            this.issueText.TabIndex = 24;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 35.95614F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 64.04386F));
            this.tableLayoutPanel1.Controls.Add(this.revCheckBox, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.effectiveText, 1, 3);
            this.tableLayoutPanel1.Controls.Add(this.issueCheckBox, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.issueText, 1, 2);
            this.tableLayoutPanel1.Controls.Add(this.text4_text, 1, 7);
            this.tableLayoutPanel1.Controls.Add(this.dateCheckBox, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.text3_text, 1, 6);
            this.tableLayoutPanel1.Controls.Add(this.effectiveCheckBox, 0, 3);
            this.tableLayoutPanel1.Controls.Add(this.text2_text, 1, 5);
            this.tableLayoutPanel1.Controls.Add(this.text1CheckBox, 0, 4);
            this.tableLayoutPanel1.Controls.Add(this.text1_text, 1, 4);
            this.tableLayoutPanel1.Controls.Add(this.text4CheckBox, 0, 7);
            this.tableLayoutPanel1.Controls.Add(this.text3CheckBox, 0, 6);
            this.tableLayoutPanel1.Controls.Add(this.text2CheckBox, 0, 5);
            this.tableLayoutPanel1.Controls.Add(this.page_rev, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.date_rev, 1, 1);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(16, 42);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 8;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(345, 216);
            this.tableLayoutPanel1.TabIndex = 28;
            // 
            // PageRevisionFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(1354, 442);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.pageView);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.checkBox1);
            this.Name = "PageRevisionFrm";
            this.ShowIcon = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show;
            this.Text = "Control Panel - ";
            this.Load += new System.EventHandler(this.PageRevisionFrm_Load);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.CheckBox checkBox1;
        public System.Windows.Forms.TextBox page_rev;
        public System.Windows.Forms.TextBox date_rev;
        private System.Windows.Forms.ListView pageView;
        private System.Windows.Forms.ColumnHeader Rev;
        private System.Windows.Forms.ColumnHeader Date;
        private System.Windows.Forms.ColumnHeader Real_Page;
        private System.Windows.Forms.ColumnHeader text1;
        private System.Windows.Forms.ColumnHeader text2;
        private System.Windows.Forms.ColumnHeader text3;
        private System.Windows.Forms.ColumnHeader text4;
        public System.Windows.Forms.TextBox text2_text;
        public System.Windows.Forms.TextBox text1_text;
        public System.Windows.Forms.TextBox text4_text;
        public System.Windows.Forms.TextBox text3_text;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.CheckBox dateCheckBox;
        private System.Windows.Forms.CheckBox revCheckBox;
        private System.Windows.Forms.CheckBox text1CheckBox;
        private System.Windows.Forms.CheckBox text4CheckBox;
        private System.Windows.Forms.CheckBox text3CheckBox;
        private System.Windows.Forms.CheckBox text2CheckBox;
        private System.Windows.Forms.ColumnHeader HeadingPage;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ColumnHeader Issue;
        private System.Windows.Forms.ColumnHeader Effective_Date;
        private System.Windows.Forms.CheckBox issueCheckBox;
        private System.Windows.Forms.CheckBox effectiveCheckBox;
        public System.Windows.Forms.TextBox effectiveText;
        public System.Windows.Forms.TextBox issueText;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
    }
}