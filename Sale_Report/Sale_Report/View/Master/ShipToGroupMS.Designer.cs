namespace Sale_Report.View.Master
{
    partial class ShipToGroupMS
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
            this.groupBox2 = new System.Windows.Forms.GroupBox();
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
            this.rdbDelShipto = new System.Windows.Forms.RadioButton();
            this.rdbChgShipto = new System.Windows.Forms.RadioButton();
            this.rdbAddShipto = new System.Windows.Forms.RadioButton();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.tbxShiptoCD = new System.Windows.Forms.TextBox();
            this.cmbShiptoMS = new System.Windows.Forms.ComboBox();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnUpdate = new System.Windows.Forms.Button();
            this.tbxShiptoDesc = new System.Windows.Forms.TextBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.dgvListShipto = new System.Windows.Forms.DataGridView();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.rdbOtherCls = new System.Windows.Forms.RadioButton();
            this.rdbPMSPCls = new System.Windows.Forms.RadioButton();
            this.rdbOEMCls = new System.Windows.Forms.RadioButton();
            this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.groupBox2.SuspendLayout();
            this.toolStrip1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel1.SuspendLayout();
            this.groupBox4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvListShipto)).BeginInit();
            this.groupBox5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.toolStrip1);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.groupBox2.Location = new System.Drawing.Point(0, 519);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(850, 39);
            this.groupBox2.TabIndex = 4;
            this.groupBox2.TabStop = false;
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
            this.toolStrip1.Location = new System.Drawing.Point(3, 11);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(844, 25);
            this.toolStrip1.TabIndex = 1;
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
            this.toolStripLabel2.Size = new System.Drawing.Size(126, 22);
            this.toolStripLabel2.Text = "T_IS_SALE_SHIPTO_MS";
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
            this.groupBox1.Controls.Add(this.rdbDelShipto);
            this.groupBox1.Controls.Add(this.rdbChgShipto);
            this.groupBox1.Controls.Add(this.rdbAddShipto);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(168, 57);
            this.groupBox1.TabIndex = 5;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Process Cls.";
            // 
            // rdbDelShipto
            // 
            this.rdbDelShipto.AutoSize = true;
            this.rdbDelShipto.Location = new System.Drawing.Point(114, 20);
            this.rdbDelShipto.Name = "rdbDelShipto";
            this.rdbDelShipto.Size = new System.Drawing.Size(41, 17);
            this.rdbDelShipto.TabIndex = 2;
            this.rdbDelShipto.Text = "Del";
            this.rdbDelShipto.UseVisualStyleBackColor = true;
            this.rdbDelShipto.Click += new System.EventHandler(this.rdbDelShipto_Click);
            // 
            // rdbChgShipto
            // 
            this.rdbChgShipto.AutoSize = true;
            this.rdbChgShipto.Enabled = false;
            this.rdbChgShipto.Location = new System.Drawing.Point(63, 20);
            this.rdbChgShipto.Name = "rdbChgShipto";
            this.rdbChgShipto.Size = new System.Drawing.Size(44, 17);
            this.rdbChgShipto.TabIndex = 1;
            this.rdbChgShipto.Text = "Chg";
            this.rdbChgShipto.UseVisualStyleBackColor = true;
            this.rdbChgShipto.Click += new System.EventHandler(this.rdbChgShipto_Click);
            // 
            // rdbAddShipto
            // 
            this.rdbAddShipto.AutoSize = true;
            this.rdbAddShipto.Location = new System.Drawing.Point(12, 20);
            this.rdbAddShipto.Name = "rdbAddShipto";
            this.rdbAddShipto.Size = new System.Drawing.Size(44, 17);
            this.rdbAddShipto.TabIndex = 0;
            this.rdbAddShipto.Text = "Add";
            this.rdbAddShipto.UseVisualStyleBackColor = true;
            this.rdbAddShipto.Click += new System.EventHandler(this.rdbAddShipto_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.tbxShiptoCD);
            this.groupBox3.Controls.Add(this.cmbShiptoMS);
            this.groupBox3.Controls.Add(this.btnCancel);
            this.groupBox3.Controls.Add(this.btnUpdate);
            this.groupBox3.Controls.Add(this.tbxShiptoDesc);
            this.groupBox3.Controls.Add(this.panel2);
            this.groupBox3.Controls.Add(this.panel1);
            this.groupBox3.Location = new System.Drawing.Point(13, 76);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(694, 116);
            this.groupBox3.TabIndex = 6;
            this.groupBox3.TabStop = false;
            // 
            // tbxShiptoCD
            // 
            this.tbxShiptoCD.Enabled = false;
            this.tbxShiptoCD.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.tbxShiptoCD.Location = new System.Drawing.Point(145, 20);
            this.tbxShiptoCD.Name = "tbxShiptoCD";
            this.tbxShiptoCD.Size = new System.Drawing.Size(153, 22);
            this.tbxShiptoCD.TabIndex = 7;
            this.tbxShiptoCD.Visible = false;
            // 
            // cmbShiptoMS
            // 
            this.cmbShiptoMS.Enabled = false;
            this.cmbShiptoMS.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.cmbShiptoMS.FormattingEnabled = true;
            this.cmbShiptoMS.Location = new System.Drawing.Point(145, 20);
            this.cmbShiptoMS.Name = "cmbShiptoMS";
            this.cmbShiptoMS.Size = new System.Drawing.Size(153, 24);
            this.cmbShiptoMS.TabIndex = 6;
            this.cmbShiptoMS.SelectedIndexChanged += new System.EventHandler(this.cmbShiptoMS_SelectedIndexChanged);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(613, 76);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 5;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnUpdate
            // 
            this.btnUpdate.Enabled = false;
            this.btnUpdate.Location = new System.Drawing.Point(532, 76);
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.Size = new System.Drawing.Size(75, 23);
            this.btnUpdate.TabIndex = 4;
            this.btnUpdate.Text = "Update";
            this.btnUpdate.UseVisualStyleBackColor = true;
            this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
            // 
            // tbxShiptoDesc
            // 
            this.tbxShiptoDesc.Enabled = false;
            this.tbxShiptoDesc.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.tbxShiptoDesc.Location = new System.Drawing.Point(145, 48);
            this.tbxShiptoDesc.Name = "tbxShiptoDesc";
            this.tbxShiptoDesc.Size = new System.Drawing.Size(543, 22);
            this.tbxShiptoDesc.TabIndex = 3;
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.SystemColors.ControlDark;
            this.panel2.Controls.Add(this.label2);
            this.panel2.Location = new System.Drawing.Point(7, 48);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(132, 22);
            this.panel2.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.label2.Location = new System.Drawing.Point(42, 3);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(87, 16);
            this.label2.TabIndex = 0;
            this.label2.Text = "Ship to Desc.";
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.ControlDark;
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(7, 20);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(132, 22);
            this.panel1.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.label1.Location = new System.Drawing.Point(58, 3);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(71, 16);
            this.label1.TabIndex = 0;
            this.label1.Text = "Ship to CD";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.dgvListShipto);
            this.groupBox4.Location = new System.Drawing.Point(12, 199);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(695, 314);
            this.groupBox4.TabIndex = 7;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "List Ship to Already Added";
            // 
            // dgvListShipto
            // 
            this.dgvListShipto.AllowUserToAddRows = false;
            this.dgvListShipto.AllowUserToDeleteRows = false;
            this.dgvListShipto.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvListShipto.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvListShipto.Location = new System.Drawing.Point(3, 16);
            this.dgvListShipto.MultiSelect = false;
            this.dgvListShipto.Name = "dgvListShipto";
            this.dgvListShipto.ReadOnly = true;
            this.dgvListShipto.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvListShipto.Size = new System.Drawing.Size(689, 295);
            this.dgvListShipto.TabIndex = 0;
            this.dgvListShipto.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvListShipto_CellDoubleClick);
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.rdbOtherCls);
            this.groupBox5.Controls.Add(this.rdbPMSPCls);
            this.groupBox5.Controls.Add(this.rdbOEMCls);
            this.groupBox5.Location = new System.Drawing.Point(187, 13);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(227, 56);
            this.groupBox5.TabIndex = 8;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Group Cls.";
            // 
            // rdbOtherCls
            // 
            this.rdbOtherCls.AutoSize = true;
            this.rdbOtherCls.Enabled = false;
            this.rdbOtherCls.Location = new System.Drawing.Point(149, 19);
            this.rdbOtherCls.Name = "rdbOtherCls";
            this.rdbOtherCls.Size = new System.Drawing.Size(51, 17);
            this.rdbOtherCls.TabIndex = 3;
            this.rdbOtherCls.TabStop = true;
            this.rdbOtherCls.Text = "Other";
            this.rdbOtherCls.UseVisualStyleBackColor = true;
            this.rdbOtherCls.Click += new System.EventHandler(this.rdbOtherCls_Click);
            // 
            // rdbPMSPCls
            // 
            this.rdbPMSPCls.AutoSize = true;
            this.rdbPMSPCls.Enabled = false;
            this.rdbPMSPCls.Location = new System.Drawing.Point(88, 19);
            this.rdbPMSPCls.Name = "rdbPMSPCls";
            this.rdbPMSPCls.Size = new System.Drawing.Size(55, 17);
            this.rdbPMSPCls.TabIndex = 1;
            this.rdbPMSPCls.TabStop = true;
            this.rdbPMSPCls.Text = "PMSP";
            this.rdbPMSPCls.UseVisualStyleBackColor = true;
            this.rdbPMSPCls.Click += new System.EventHandler(this.rdbPMSPCls_Click);
            // 
            // rdbOEMCls
            // 
            this.rdbOEMCls.AutoSize = true;
            this.rdbOEMCls.Enabled = false;
            this.rdbOEMCls.Location = new System.Drawing.Point(33, 19);
            this.rdbOEMCls.Name = "rdbOEMCls";
            this.rdbOEMCls.Size = new System.Drawing.Size(49, 17);
            this.rdbOEMCls.TabIndex = 0;
            this.rdbOEMCls.TabStop = true;
            this.rdbOEMCls.Text = "OEM";
            this.rdbOEMCls.UseVisualStyleBackColor = true;
            this.rdbOEMCls.Click += new System.EventHandler(this.rdbOEMCls_Click);
            // 
            // ShipToGroupMS
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(850, 558);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.groupBox2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.MaximizeBox = false;
            this.Name = "ShipToGroupMS";
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "ShipTo Group Master";
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvListShipto)).EndInit();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton rdbDelShipto;
        private System.Windows.Forms.RadioButton rdbAddShipto;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.TextBox tbxShiptoDesc;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnUpdate;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.DataGridView dgvListShipto;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.RadioButton rdbPMSPCls;
        private System.Windows.Forms.RadioButton rdbOEMCls;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripLabel toolStripLabel1;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripLabel toolStripLabel2;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripLabel toolStripLabel3;
        private System.Windows.Forms.ToolStripLabel txtDateTime;
        private System.Windows.Forms.ToolStripLabel txtSchema;
        private System.Windows.Forms.ToolStripLabel txtNameSchema;
        private System.Windows.Forms.ComboBox cmbShiptoMS;
        private System.Windows.Forms.RadioButton rdbChgShipto;
        private System.Windows.Forms.BindingSource bindingSource1;
        private System.Windows.Forms.TextBox tbxShiptoCD;
        private System.Windows.Forms.RadioButton rdbOtherCls;
    }
}