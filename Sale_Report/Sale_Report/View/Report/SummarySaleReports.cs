using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Sale_Report.Model;
using Sale_Report.View.Loading;

namespace Sale_Report.View.Report
{
    public partial class SummarySaleReports : Form
    {
        SaleResult saleResult = new SaleResult();

        DataSet dsActualThisMonth = new DataSet();
        DataSet dsFrstThisMonth = new DataSet();
        DataSet dsActualPrevMonth = new DataSet();
        DataSet dsFrstPrevMonth = new DataSet();

        string processCls = "";

        bool flgSearchComplete = true;

        public SummarySaleReports()
        {
            InitializeComponent();

            txtNameSchema.Text = "SCHEMA : " + saleResult.obj.oracle.User;
            txtDateTime.Text = DateTime.Now.ToString();
            timer1.Start();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            txtDateTime.Text = DateTime.Now.ToString();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            processCls = "search";
            summaryProcess();
        }

        private void summaryProcess()
        {
            backgroundWorker1.RunWorkerAsync();
            saleResult.obj.oracle.frmLoad = new Frm_Loading();
            saleResult.obj.oracle.frmLoad.ShowDialog();


            if (flgSearchComplete)
            {
                dgvSearchResult.DataSource = dsActualThisMonth.Tables[0];
            }
            else
            {
                MessageBox.Show("No data.", "Result", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                if (processCls == "search")
                {
                    string yearThisMonth = dtpYearMonth.Value.ToString("yyyyMM");
                    string yearPrevMonth = dtpYearMonth.Value.AddMonths(-1).ToString("yyyyMM");
                    DataSet dsTmp1 = null;
                    DataSet dsTmp2 = null;

                    // ACTUAL RESULT
                    string[] groupDesc = { "OTHER", "PMSP", "OEM" };
                    for (int i = 0; i < groupDesc.Length; i++)
                    {
                        dsTmp1 = null;
                        dsTmp2 = null;
                        dsTmp1 = saleResult.selectOverAllActual(yearThisMonth, groupDesc[i]);
                        dsTmp2 = saleResult.selectOverAllActual(yearPrevMonth, groupDesc[i]);

                        dsActualThisMonth.Merge(dsTmp1);
                        dsActualPrevMonth.Merge(dsTmp2);
                    }


                    // FORECAST RESULT
                    string[] groupDesc2 = { "OTHER", "PMSP", "OEM" };
                    for (int i = 0; i < groupDesc2.Length; i++)
                    {
                        if (i == 1)
                        {
                            dsTmp1 = null;
                            dsTmp2 = null;
                            dsTmp1 = saleResult.selectOverAllPMSPForecast(yearThisMonth);
                            dsTmp2 = saleResult.selectOverAllPMSPForecast(yearPrevMonth);
                        }
                        else
                        {
                            dsTmp1 = null;
                            dsTmp2 = null;
                            dsTmp1 = saleResult.selectOverAllForecast(yearThisMonth, groupDesc2[i]);
                            dsTmp2 = saleResult.selectOverAllForecast(yearPrevMonth, groupDesc2[i]);

                            if (dsTmp1.Tables[0].Rows.Count == 0)
                            {
                                DataRow dr = dsTmp1.Tables[0].NewRow();

                                dr["groupdesc"] = groupDesc2[i];
                                dr["i_frst_year_mnth"] = yearThisMonth;
                                dr["i_amt_frst"] = 0;

                                dsTmp1.Tables[0].Rows.Add(dr);
                            }

                            if (dsTmp2.Tables[0].Rows.Count == 0)
                            {
                                DataRow dr = dsTmp2.Tables[0].NewRow();

                                dr["groupdesc"] = groupDesc2[i];
                                dr["i_frst_year_mnth"] = yearPrevMonth;
                                dr["i_amt_frst"] = 0;

                                dsTmp2.Tables[0].Rows.Add(dr);
                            }
                        }

                        dsFrstThisMonth.Merge(dsTmp1);
                        dsFrstPrevMonth.Merge(dsTmp2);

                    }
                }
                else if (processCls == "export")
                { 
                    
                }
            }
            catch (Exception ex)
            {
                flgSearchComplete = false;

                MessageBox.Show(ex.Message);
                if (saleResult.obj.oracle.frmLoad != null && saleResult.obj.oracle.frmLoad.Visible)
                {
                    saleResult.obj.oracle.frmLoad.Dispose();
                }
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (saleResult.obj.oracle.frmLoad != null && saleResult.obj.oracle.frmLoad.Visible)
            {
                saleResult.obj.oracle.frmLoad.Dispose();
            }
        }
    }
}
