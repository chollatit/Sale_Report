using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Sale_Report.View.Master;
using Sale_Report.View.Report;

namespace Sale_Report
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        private void groupShipToToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.MdiChildren.Count() > 0)
                this.ActiveMdiChild.Close();

            ShipToGroupMS shipto = new ShipToGroupMS();
            shipto.MdiParent = this;
            shipto.Show();
            shipto.WindowState = FormWindowState.Minimized;
            shipto.WindowState = FormWindowState.Maximized;
        }

        private void partGroupOEMToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.MdiChildren.Count() > 0)
                this.ActiveMdiChild.Close();

            OEMPartGroupMS oemPart = new OEMPartGroupMS();
            oemPart.MdiParent = this;
            oemPart.Show();
            oemPart.WindowState = FormWindowState.Minimized;
            oemPart.WindowState = FormWindowState.Maximized;
        }

        private void summarySaleReportsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.MdiChildren.Count() > 0)
                this.ActiveMdiChild.Close();

            SummarySaleReports summarySaleReport = new SummarySaleReports();
            summarySaleReport.MdiParent = this;
            summarySaleReport.Show();
            summarySaleReport.WindowState = FormWindowState.Minimized;
            summarySaleReport.WindowState = FormWindowState.Maximized;
        }

        private void oEMServicePartGroupToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.MdiChildren.Count() > 0)
                this.ActiveMdiChild.Close();

            OEMServicePartGroupMS oemPart = new OEMServicePartGroupMS();
            oemPart.MdiParent = this;
            oemPart.Show();
            oemPart.WindowState = FormWindowState.Minimized;
            oemPart.WindowState = FormWindowState.Maximized;
        }
    }
}
