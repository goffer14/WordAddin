namespace eDocs_Editor
{
    partial class LoepCreator
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
            this.pageView = new System.Windows.Forms.ListView();
            this.Doc_Name = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.location = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.number = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.button2 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(397, 188);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(97, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Add Document";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // pageView
            // 
            this.pageView.AutoArrange = false;
            this.pageView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.number,
            this.Doc_Name,
            this.location});
            this.pageView.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.pageView.FullRowSelect = true;
            this.pageView.GridLines = true;
            this.pageView.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.pageView.Location = new System.Drawing.Point(22, 12);
            this.pageView.MultiSelect = false;
            this.pageView.Name = "pageView";
            this.pageView.Size = new System.Drawing.Size(472, 170);
            this.pageView.TabIndex = 7;
            this.pageView.UseCompatibleStateImageBehavior = false;
            this.pageView.View = System.Windows.Forms.View.Details;
            // 
            // Doc_Name
            // 
            this.Doc_Name.Text = "Document Name";
            this.Doc_Name.Width = 172;
            // 
            // location
            // 
            this.location.Text = "Location";
            this.location.Width = 212;
            // 
            // number
            // 
            this.number.Text = "Doc number";
            this.number.Width = 77;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(304, 188);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(87, 23);
            this.button2.TabIndex = 8;
            this.button2.Text = "Create LOEP";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // LoepCreator
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(508, 218);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.pageView);
            this.Controls.Add(this.button1);
            this.Name = "LoepCreator";
            this.Text = "LoepCreator";
            this.Load += new System.EventHandler(this.LoepCreator_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ListView pageView;
        private System.Windows.Forms.ColumnHeader number;
        private System.Windows.Forms.ColumnHeader Doc_Name;
        private System.Windows.Forms.ColumnHeader location;
        private System.Windows.Forms.Button button2;
    }
}