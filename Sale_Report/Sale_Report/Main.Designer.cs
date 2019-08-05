namespace Sale_Report
{
    partial class Main
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
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.masterToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.groupShipToToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.partGroupOEMToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.reportToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.summarySaleReportsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.oEMServicePartGroupToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.masterToolStripMenuItem,
            this.reportToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(884, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // masterToolStripMenuItem
            // 
            this.masterToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.groupShipToToolStripMenuItem,
            this.partGroupOEMToolStripMenuItem,
            this.oEMServicePartGroupToolStripMenuItem});
            this.masterToolStripMenuItem.Name = "masterToolStripMenuItem";
            this.masterToolStripMenuItem.Size = new System.Drawing.Size(55, 20);
            this.masterToolStripMenuItem.Text = "Master";
            // 
            // groupShipToToolStripMenuItem
            // 
            this.groupShipToToolStripMenuItem.Name = "groupShipToToolStripMenuItem";
            this.groupShipToToolStripMenuItem.Size = new System.Drawing.Size(200, 22);
            this.groupShipToToolStripMenuItem.Text = "Ship To Group";
            this.groupShipToToolStripMenuItem.Click += new System.EventHandler(this.groupShipToToolStripMenuItem_Click);
            // 
            // partGroupOEMToolStripMenuItem
            // 
            this.partGroupOEMToolStripMenuItem.Name = "partGroupOEMToolStripMenuItem";
            this.partGroupOEMToolStripMenuItem.Size = new System.Drawing.Size(200, 22);
            this.partGroupOEMToolStripMenuItem.Text = "OEM Part Group";
            this.partGroupOEMToolStripMenuItem.Click += new System.EventHandler(this.partGroupOEMToolStripMenuItem_Click);
            // 
            // reportToolStripMenuItem
            // 
            this.reportToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.summarySaleReportsToolStripMenuItem});
            this.reportToolStripMenuItem.Name = "reportToolStripMenuItem";
            this.reportToolStripMenuItem.Size = new System.Drawing.Size(54, 20);
            this.reportToolStripMenuItem.Text = "Report";
            // 
            // summarySaleReportsToolStripMenuItem
            // 
            this.summarySaleReportsToolStripMenuItem.Name = "summarySaleReportsToolStripMenuItem";
            this.summarySaleReportsToolStripMenuItem.Size = new System.Drawing.Size(192, 22);
            this.summarySaleReportsToolStripMenuItem.Text = "Summary Sale Reports";
            this.summarySaleReportsToolStripMenuItem.Click += new System.EventHandler(this.summarySaleReportsToolStripMenuItem_Click);
            // 
            // oEMServicePartGroupToolStripMenuItem
            // 
            this.oEMServicePartGroupToolStripMenuItem.Name = "oEMServicePartGroupToolStripMenuItem";
            this.oEMServicePartGroupToolStripMenuItem.Size = new System.Drawing.Size(200, 22);
            this.oEMServicePartGroupToolStripMenuItem.Text = "OEM Service Part Group";
            this.oEMServicePartGroupToolStripMenuItem.Click += new System.EventHandler(this.oEMServicePartGroupToolStripMenuItem_Click);
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(884, 592);
            this.Controls.Add(this.menuStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.IsMdiContainer = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Main";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Sale Report";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem reportToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem masterToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem groupShipToToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem partGroupOEMToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem summarySaleReportsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem oEMServicePartGroupToolStripMenuItem;
    }
}

