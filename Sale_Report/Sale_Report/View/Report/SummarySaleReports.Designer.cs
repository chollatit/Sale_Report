namespace Sale_Report.View.Report
{
    partial class SummarySaleReports
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
            this.components = new System.ComponentModel.Container();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStripLabel1 = new System.Windows.Forms.ToolStripLabel();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripLabel2 = new System.Windows.Forms.ToolStripLabel();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripLabel3 = new System.Windows.Forms.ToolStripLabel();
            this.txtDateTime = new System.Windows.Forms.ToolStripLabel();
            this.txtSchema = new System.Windows.Forms.ToolStripLabel();
            this.txtNameSchema = new System.Windows.Forms.ToolStripLabel();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnExport = new System.Windows.Forms.Button();
            this.btnSearch = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.dtpYearMonth = new System.Windows.Forms.DateTimePicker();
            this.panel1 = new System.Windows.Forms.Panel();
            this.Month = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.dgvSearchResult = new System.Windows.Forms.DataGridView();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.panel3 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.clbReport = new System.Windows.Forms.CheckedListBox();
            this.panel4 = new System.Windows.Forms.Panel();
            this.lblSelectAll = new System.Windows.Forms.LinkLabel();
            this.lblUnSelectAll = new System.Windows.Forms.LinkLabel();
            this.toolStrip1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSearchResult)).BeginInit();
            this.panel3.SuspendLayout();
            this.panel4.SuspendLayout();
            this.SuspendLayout();
            // 
            // toolStrip1
            // 
            this.toolStrip1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.toolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripLabel1,
            this.toolStripSeparator1,
            this.toolStripLabel2,
            this.toolStripSeparator2,
            this.toolStripLabel3,
            this.txtDateTime,
            this.txtSchema,
            this.txtNameSchema});
            this.toolStrip1.Location = new System.Drawing.Point(0, 533);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(850, 25);
            this.toolStrip1.TabIndex = 3;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // toolStripLabel1
            // 
            this.toolStripLabel1.Name = "toolStripLabel1";
            this.toolStripLabel1.Size = new System.Drawing.Size(63, 22);
            this.toolStripLabel1.Text = "Version 1.0";
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 25);
            // 
            // toolStripLabel2
            // 
            this.toolStripLabel2.Name = "toolStripLabel2";
            this.toolStripLabel2.Size = new System.Drawing.Size(192, 22);
            this.toolStripLabel2.Text = "T_SHIP_TR && T_FRST_SO_HEAD_TR";
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 25);
            // 
            // toolStripLabel3
            // 
            this.toolStripLabel3.Name = "toolStripLabel3";
            this.toolStripLabel3.Size = new System.Drawing.Size(40, 22);
            this.toolStripLabel3.Text = "Date : ";
            // 
            // txtDateTime
            // 
            this.txtDateTime.Name = "txtDateTime";
            this.txtDateTime.Size = new System.Drawing.Size(0, 22);
            // 
            // txtSchema
            // 
            this.txtSchema.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.txtSchema.Name = "txtSchema";
            this.txtSchema.Size = new System.Drawing.Size(0, 22);
            // 
            // txtNameSchema
            // 
            this.txtNameSchema.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.txtNameSchema.Name = "txtNameSchema";
            this.txtNameSchema.Size = new System.Drawing.Size(64, 22);
            this.txtNameSchema.Text = "SCHEMA : ";
            // 
            // timer1
            // 
            this.timer1.Interval = 1000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.panel4);
            this.groupBox1.Controls.Add(this.panel3);
            this.groupBox1.Controls.Add(this.btnExport);
            this.groupBox1.Controls.Add(this.btnSearch);
            this.groupBox1.Controls.Add(this.panel2);
            this.groupBox1.Controls.Add(this.panel1);
            this.groupBox1.Location = new System.Drawing.Point(11, 4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(827, 194);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Search";
            // 
            // btnExport
            // 
            this.btnExport.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.btnExport.Location = new System.Drawing.Point(738, 135);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(83, 41);
            this.btnExport.TabIndex = 4;
            this.btnExport.Text = "Export";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // btnSearch
            // 
            this.btnSearch.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.btnSearch.Location = new System.Drawing.Point(647, 135);
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Size = new System.Drawing.Size(85, 41);
            this.btnSearch.TabIndex = 3;
            this.btnSearch.Text = "Search";
            this.btnSearch.UseVisualStyleBackColor = true;
            this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.SystemColors.ControlLight;
            this.panel2.Controls.Add(this.dtpYearMonth);
            this.panel2.Location = new System.Drawing.Point(153, 20);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(224, 40);
            this.panel2.TabIndex = 2;
            // 
            // dtpYearMonth
            // 
            this.dtpYearMonth.CustomFormat = "MMM yyyy";
            this.dtpYearMonth.DropDownAlign = System.Windows.Forms.LeftRightAlignment.Right;
            this.dtpYearMonth.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.dtpYearMonth.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpYearMonth.Location = new System.Drawing.Point(12, 5);
            this.dtpYearMonth.Name = "dtpYearMonth";
            this.dtpYearMonth.Size = new System.Drawing.Size(200, 29);
            this.dtpYearMonth.TabIndex = 1;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.panel1.Controls.Add(this.Month);
            this.panel1.Location = new System.Drawing.Point(7, 20);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(140, 40);
            this.panel1.TabIndex = 0;
            // 
            // Month
            // 
            this.Month.AutoSize = true;
            this.Month.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.Month.Location = new System.Drawing.Point(61, 8);
            this.Month.Name = "Month";
            this.Month.Size = new System.Drawing.Size(68, 24);
            this.Month.TabIndex = 0;
            this.Month.Text = "Month:";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.dgvSearchResult);
            this.groupBox2.Location = new System.Drawing.Point(11, 204);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(827, 326);
            this.groupBox2.TabIndex = 5;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Result";
            // 
            // dgvSearchResult
            // 
            this.dgvSearchResult.AllowUserToAddRows = false;
            this.dgvSearchResult.AllowUserToDeleteRows = false;
            this.dgvSearchResult.AllowUserToOrderColumns = true;
            this.dgvSearchResult.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvSearchResult.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvSearchResult.Location = new System.Drawing.Point(3, 16);
            this.dgvSearchResult.Name = "dgvSearchResult";
            this.dgvSearchResult.ReadOnly = true;
            this.dgvSearchResult.Size = new System.Drawing.Size(821, 307);
            this.dgvSearchResult.TabIndex = 0;
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.WorkerReportsProgress = true;
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            this.backgroundWorker1.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker1_ProgressChanged);
            this.backgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker1_RunWorkerCompleted);
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.panel3.Controls.Add(this.label1);
            this.panel3.Location = new System.Drawing.Point(7, 66);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(140, 40);
            this.panel3.TabIndex = 5;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.label1.Location = new System.Drawing.Point(61, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(71, 24);
            this.label1.TabIndex = 0;
            this.label1.Text = "Report:";
            // 
            // clbReport
            // 
            this.clbReport.FormattingEnabled = true;
            this.clbReport.Items.AddRange(new object[] {
            "Report 1 #Over all sales result",
            "Report 2 #OEM sales result",
            "Report 3 #PMSP sales result",
            "Report 4 #Overall accumulate sales result",
            "Report 5 #OEM accumulate sales result",
            "Report 6 #PMSP accumulate sales result",
            "Report 7 #Over all sales forecast",
            "Report 8 #OEM sales forecast",
            "Report 9 #PMSP sales forecast",
            "Report 10 #Overall accumulate sales forecast",
            "Report 11 #OEM accumulate sales forecast",
            "Report 12 #PMSP accumulate sales forecast"});
            this.clbReport.Location = new System.Drawing.Point(12, 5);
            this.clbReport.Name = "clbReport";
            this.clbReport.Size = new System.Drawing.Size(268, 94);
            this.clbReport.TabIndex = 0;
            // 
            // panel4
            // 
            this.panel4.BackColor = System.Drawing.SystemColors.ControlLight;
            this.panel4.Controls.Add(this.lblUnSelectAll);
            this.panel4.Controls.Add(this.lblSelectAll);
            this.panel4.Controls.Add(this.clbReport);
            this.panel4.Location = new System.Drawing.Point(153, 66);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(293, 122);
            this.panel4.TabIndex = 7;
            // 
            // lblSelectAll
            // 
            this.lblSelectAll.AutoSize = true;
            this.lblSelectAll.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.lblSelectAll.Location = new System.Drawing.Point(12, 104);
            this.lblSelectAll.Name = "lblSelectAll";
            this.lblSelectAll.Size = new System.Drawing.Size(74, 16);
            this.lblSelectAll.TabIndex = 1;
            this.lblSelectAll.TabStop = true;
            this.lblSelectAll.Text = "Select All";
            this.lblSelectAll.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lblSelectAll_LinkClicked);
            // 
            // lblUnSelectAll
            // 
            this.lblUnSelectAll.AutoSize = true;
            this.lblUnSelectAll.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.lblUnSelectAll.Location = new System.Drawing.Point(92, 104);
            this.lblUnSelectAll.Name = "lblUnSelectAll";
            this.lblUnSelectAll.Size = new System.Drawing.Size(91, 16);
            this.lblUnSelectAll.TabIndex = 2;
            this.lblUnSelectAll.TabStop = true;
            this.lblUnSelectAll.Text = "Unselect All";
            this.lblUnSelectAll.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lblUnSelectAll_LinkClicked);
            // 
            // SummarySaleReports
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(850, 558);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.toolStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.MaximizeBox = false;
            this.Name = "SummarySaleReports";
            this.Text = "Summary Sale Reports";
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvSearchResult)).EndInit();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.panel4.ResumeLayout(false);
            this.panel4.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripLabel toolStripLabel1;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripLabel toolStripLabel2;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripLabel toolStripLabel3;
        private System.Windows.Forms.ToolStripLabel txtDateTime;
        private System.Windows.Forms.ToolStripLabel txtSchema;
        private System.Windows.Forms.ToolStripLabel txtNameSchema;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.DateTimePicker dtpYearMonth;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label Month;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Button btnSearch;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.DataGridView dgvSearchResult;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckedListBox clbReport;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.LinkLabel lblSelectAll;
        private System.Windows.Forms.LinkLabel lblUnSelectAll;
    }
}