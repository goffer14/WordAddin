namespace eDocs_Editor
{
    partial class monitoringFrm
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
            this.label7 = new System.Windows.Forms.Label();
            this.date_rev = new System.Windows.Forms.TextBox();
            this.page_rev = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.label7.Location = new System.Drawing.Point(6, 10);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(40, 16);
            this.label7.TabIndex = 20;
            this.label7.Text = "Use:";
            // 
            // date_rev
            // 
            this.date_rev.Location = new System.Drawing.Point(93, 67);
            this.date_rev.Name = "date_rev";
            this.date_rev.Size = new System.Drawing.Size(150, 22);
            this.date_rev.TabIndex = 19;
            // 
            // page_rev
            // 
            this.page_rev.Location = new System.Drawing.Point(93, 38);
            this.page_rev.Name = "page_rev";
            this.page_rev.Size = new System.Drawing.Size(150, 22);
            this.page_rev.TabIndex = 18;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(168, 95);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 23;
            this.button1.Text = "Start";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 41);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(64, 16);
            this.label1.TabIndex = 24;
            this.label1.Text = "Revision:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 70);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(40, 16);
            this.label2.TabIndex = 25;
            this.label2.Text = "Date:";
            // 
            // monitoringFrm
            // 
            //this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            //this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(260, 124);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.date_rev);
            this.Controls.Add(this.page_rev);
            this.Name = "monitoringFrm";
            this.Text = "monitoringFrm";
            this.Load += new System.EventHandler(this.monitoringFrm_Load);

            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F); //IMPORTANT
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label7;
        public System.Windows.Forms.TextBox date_rev;
        public System.Windows.Forms.TextBox page_rev;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
    }
}