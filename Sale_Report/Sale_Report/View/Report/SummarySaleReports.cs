﻿using System;
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
using System.IO;
using OfficeOpenXml;
using System.Diagnostics;

namespace Sale_Report.View.Report
{
    public partial class SummarySaleReports : Form
    {
        SaleResult saleResult = new SaleResult();
        AccumulateResult accResult = new AccumulateResult();
        SaleForecast saleForecast = new SaleForecast();
        AccumulateForecast accForcast = new AccumulateForecast();

        // REPORT1
        DataSet dsOverallActualThisMonth = new DataSet();
        DataSet dsOverallFrstThisMonth = new DataSet();
        DataSet dsOverallActualPrevMonth = new DataSet();
        DataSet dsOverallFrstPrevMonth = new DataSet();

        // REPORT2
        DataSet dsProjOEMActualThisMonth = new DataSet();
        DataSet dsProjOEMActualPrevMonth = new DataSet();
        DataSet dsProjOEMFrstThisMonth = new DataSet();
        DataSet dsProjOEMFrstPrevMonth = new DataSet();

        // REPORT3
        DataSet dsProjPMSPActualThisMonth = new DataSet();
        DataSet dsProjPMSPActualPrevMonth = new DataSet();
        DataSet dsProjPMSPForecastThisMonth = new DataSet();
        DataSet dsProjPMSPForecastPrevMonth = new DataSet();

        // REPORT4
        DataSet dsOverallYearOAP = new DataSet();
        DataSet dsAccumulateOAP = new DataSet();
        DataSet dsAccActualResult = new DataSet();

        // REPORT5
        DataSet dsProjOEMYearOAP = new DataSet();
        DataSet dsProjOEMAccumulateOAP = new DataSet();
        DataSet dsProjOEMAccActualResult = new DataSet();

        // REPORT6
        DataSet dsProjPMSPYearOAP = new DataSet();
        DataSet dsProjPMSPAccumulateOAP = new DataSet();
        DataSet dsProjPMSPAccActualResult = new DataSet();

        //Report7
        DataSet dsOverallSaleForecat = new DataSet();

        //Report8
        DataSet dsOEMSaleForecast = new DataSet();

        //Report9
        DataSet dsPMSPSaleForecast = new DataSet();

        //Report10
        DataSet dsAccOverallOAP = new DataSet();
        DataSet dsAccOAPPrevYear = new DataSet();
        DataSet dsAccOAPThisYear = new DataSet();

        //Report11
        DataSet dsOEMAccOverallOAP = new DataSet();
        DataSet dsOEMAccOAPPrevYear = new DataSet();
        DataSet dsOEMAccOAPThisYear = new DataSet();

        //Report12
        DataSet dsPMSPAccOverallOAP = new DataSet();
        DataSet dsPMSPAccOAPPrevYear = new DataSet();
        DataSet dsPMSPAccOAPThisYear = new DataSet();

        string PATH_Program = Directory.GetCurrentDirectory();
        string pathDesktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        string fileNameExcelExport = "sale_report";
        string SAVE_PATH = "";
        string processCls = "";

        bool flgSearchComplete = true;
        bool flgExportComplete = true;

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
                dgvSearchResult.DataSource = dsOverallActualThisMonth.Tables[0];
            }
            else
            {
                MessageBox.Show("No data.", "Result", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            if (dsOverallActualThisMonth.Tables.Count > 0)
            {
                processCls = "export";
                exportExcelProcess();
            }
            else
            {
                btnSearch.PerformClick();

                processCls = "export";
                exportExcelProcess();
            }

            //processCls = "export";
            //exportExcelProcess();
        }

        private void exportExcelProcess()
        {
            backgroundWorker1.RunWorkerAsync();
            saleResult.obj.oracle.frmLoad = new Frm_Loading();
            saleResult.obj.oracle.frmLoad.ShowDialog();

            if (flgExportComplete)
            {
                MessageBox.Show("Export Complete.", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
                // Open file
                Process.Start(SAVE_PATH);
            }
            else
            {
                MessageBox.Show("Export Incomplete.", "Result", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                flgSearchComplete = true;
                flgExportComplete = true;

                if (processCls == "search")
                {
                    processOverallSalesResult();
                    processAccumulateSalesResult();
                    processOverallSalesForecast();
                    processAccumulateSalesForecast();
                }
                else if (processCls == "export")
                {
                    exportExcel();
                }
            }
            catch (Exception ex)
            {
                flgSearchComplete = false;
                flgExportComplete = false;

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

        private void processOverallSalesResult()
        {
            string yearThisMonth = dtpYearMonth.Value.ToString("yyyyMM");
            string yearPrevMonth = dtpYearMonth.Value.AddMonths(-1).ToString("yyyyMM");

            report1(yearThisMonth, yearPrevMonth);
            report2(yearThisMonth, yearPrevMonth);
            report3(yearThisMonth, yearPrevMonth);
        }

        #region OVERALL SALE RESULT

        private void report1(string yearThisMonth, string yearPrevMonth)
        {
            DataSet dsTmp1 = null;
            DataSet dsTmp2 = null;

            // ACTUAL
            string[] groupDesc = { "OTHER", "PMSP", "OEM" };
            for (int i = 0; i < groupDesc.Length; i++)
            {
                dsTmp1 = null;
                dsTmp2 = null;
                dsTmp1 = saleResult.selectOverAllActual(yearThisMonth, groupDesc[i]);
                dsTmp2 = saleResult.selectOverAllActual(yearPrevMonth, groupDesc[i]);

                dsOverallActualThisMonth.Merge(dsTmp1);
                dsOverallActualPrevMonth.Merge(dsTmp2);
            }

            // FORECAST
            for (int i = 0; i < groupDesc.Length; i++)
            {
                dsTmp1 = null;
                dsTmp2 = null;
                if (i == 1)
                {
                    dsTmp1 = saleResult.selectOverAllPMSPForecast(yearThisMonth);
                    dsTmp2 = saleResult.selectOverAllPMSPForecast(yearPrevMonth);
                }
                else
                {
                    dsTmp1 = saleResult.selectOverAllForecast(yearThisMonth, groupDesc[i]);
                    dsTmp2 = saleResult.selectOverAllForecast(yearPrevMonth, groupDesc[i]);

                    if (dsTmp1.Tables[0].Rows.Count == 0)
                    {
                        DataRow dr = dsTmp1.Tables[0].NewRow();

                        dr["groupdesc"] = groupDesc[i];
                        dr["i_frst_year_mnth"] = yearThisMonth;
                        dr["i_amt_frst"] = 0;

                        dsTmp1.Tables[0].Rows.Add(dr);
                    }

                    if (dsTmp2.Tables[0].Rows.Count == 0)
                    {
                        DataRow dr = dsTmp2.Tables[0].NewRow();

                        dr["groupdesc"] = groupDesc[i];
                        dr["i_frst_year_mnth"] = yearPrevMonth;
                        dr["i_amt_frst"] = 0;

                        dsTmp2.Tables[0].Rows.Add(dr);
                    }
                }

                dsOverallFrstThisMonth.Merge(dsTmp1);
                dsOverallFrstPrevMonth.Merge(dsTmp2);
            }
        }

        private void report2(string yearThisMonth, string yearPrevMonth)
        {
            DataSet dsTmp1 = null;
            DataSet dsTmp2 = null;

            //ACTUAL
            string[] groupDesc = { "OTHER", "NC Data Sales", "492B", "Hiace", "OEM", "640A" };
            string[] projectDesc = { "", "", "'PJ09'", "'PJ03', 'PJ10'", "", "'PJ05', 'PJ08'" };
            for (int i = 0; i < groupDesc.Length; i++)
            {
                dsTmp1 = null;
                dsTmp2 = null;

                if (i == 0)
                {
                    dsTmp1 = saleResult.selectProjectOEMOtherActual(yearThisMonth);
                    dsTmp2 = saleResult.selectProjectOEMOtherActual(yearPrevMonth);
                }
                else if (i == 1)
                {
                    dsTmp1 = saleResult.selectProjectOEMNCDataActual(yearThisMonth);
                    dsTmp2 = saleResult.selectProjectOEMNCDataActual(yearPrevMonth);
                }
                else if (i == 4)
                {
                    dsTmp1 = saleResult.selectProjectOEMOEMPartActual(yearThisMonth);
                    dsTmp2 = saleResult.selectProjectOEMOEMPartActual(yearPrevMonth);
                }
                else
                {
                    dsTmp1 = saleResult.selectProjectOEMActual(yearThisMonth, groupDesc[i], projectDesc[i]);
                    dsTmp2 = saleResult.selectProjectOEMActual(yearPrevMonth, groupDesc[i], projectDesc[i]);
                }

                dsProjOEMActualThisMonth.Merge(dsTmp1);
                dsProjOEMActualPrevMonth.Merge(dsTmp2);
            }

            //FORECAST
            for (int i = 0; i < groupDesc.Length; i++)
            {
                dsTmp1 = null;
                dsTmp2 = null;

                if (i == 1)
                {
                    dsTmp1 = saleResult.selectProjectOEMNCDataForecast(yearThisMonth);
                    dsTmp2 = saleResult.selectProjectOEMNCDataForecast(yearPrevMonth);
                }
                else if (i == 4)
                {
                    dsTmp1 = saleResult.selectProjectOEMOEMPartForecast(yearThisMonth);
                    dsTmp2 = saleResult.selectProjectOEMOEMPartForecast(yearPrevMonth);
                }
                else
                {
                    dsTmp1 = saleResult.selectProjectOEMForecast(yearThisMonth, groupDesc[i], projectDesc[i]);
                    dsTmp2 = saleResult.selectProjectOEMForecast(yearPrevMonth, groupDesc[i], projectDesc[i]);
                }

                if (dsTmp1.Tables[0].Rows.Count == 0)
                {
                    DataRow dr = dsTmp1.Tables[0].NewRow();

                    dr["groupdesc"] = groupDesc[i];
                    dr["i_frst_year_mnth"] = yearThisMonth;
                    dr["i_amt_frst"] = 0;

                    dsTmp1.Tables[0].Rows.Add(dr);
                }

                if (dsTmp2.Tables[0].Rows.Count == 0)
                {
                    DataRow dr = dsTmp2.Tables[0].NewRow();

                    dr["groupdesc"] = groupDesc[i];
                    dr["i_frst_year_mnth"] = yearPrevMonth;
                    dr["i_amt_frst"] = 0;

                    dsTmp2.Tables[0].Rows.Add(dr);
                }

                dsProjOEMFrstThisMonth.Merge(dsTmp1);
                dsProjOEMFrstPrevMonth.Merge(dsTmp2);
            }
        }

        private void report3(string yearThisMonth, string yearPrevMonth)
        {
            DataSet dsTmp1 = null;
            DataSet dsTmp2 = null;

            dsTmp1 = saleResult.selectProjectPMSP640ServiceActual(yearThisMonth);
            dsTmp2 = saleResult.selectProjectPMSP640ServiceActual(yearPrevMonth);
            dsProjPMSPActualThisMonth.Merge(dsTmp1);
            dsProjPMSPActualPrevMonth.Merge(dsTmp2);

            dsTmp1 = null;
            dsTmp2 = null;
            dsTmp1 = saleResult.selectProjectPMSPActual(yearThisMonth);
            dsTmp2 = saleResult.selectProjectPMSPActual(yearPrevMonth);
            dsProjPMSPActualThisMonth.Merge(dsTmp1);
            dsProjPMSPActualPrevMonth.Merge(dsTmp2);

            dsTmp1 = null;
            dsTmp2 = null;
            dsTmp1 = saleResult.selectProjectPMSP640ServiceForecast(yearThisMonth);
            dsTmp2 = saleResult.selectProjectPMSP640ServiceForecast(yearPrevMonth);
            dsProjPMSPForecastThisMonth.Merge(dsTmp1);
            dsProjPMSPForecastPrevMonth.Merge(dsTmp2);

            dsTmp1 = null;
            dsTmp2 = null;
            dsTmp1 = saleResult.selectProjectPMSPForecast(yearThisMonth);
            dsTmp2 = saleResult.selectProjectPMSPForecast(yearPrevMonth);
            dsProjPMSPForecastThisMonth.Merge(dsTmp1);
            dsProjPMSPForecastPrevMonth.Merge(dsTmp2);
        }

        #endregion

        private void processAccumulateSalesResult()
        {
            int month = Convert.ToInt32(dtpYearMonth.Value.ToString("MM"));

            string yearMonthBegin = "";
            string yearMonthEnd = dtpYearMonth.Value.ToString("MMyyyy");

            if (month >= 4 && month <= 12)
            {
                yearMonthBegin = "04" + dtpYearMonth.Value.ToString("yyyy");
            }
            else if (month >= 1 && month <= 3)
            {
                yearMonthBegin = "04" + dtpYearMonth.Value.AddYears(-1).ToString("yyyy");
            }

            report4(yearMonthBegin, yearMonthEnd);
            report5(yearMonthBegin, yearMonthEnd);
            report6(yearMonthBegin, yearMonthEnd);
        }

        #region ACCUMULATE SALES RESULT

        private void report4(string yearMonthBegin, string yearMonthEnd)
        {
            DataSet dsTmp = null;

            // ALL OAP
            string[] groupCD = { "OTHER", "PMSP", "OEM" };
            for (int i = 0; i < groupCD.Length; i++)
            {
                dsTmp = null;
                //dsTmp = accResult.selectOverallYearOAP(yearMonthEnd, groupCD[i]);


                if (i != 1)
                {
                    dsTmp = accResult.selectOverallYearOAP(yearMonthEnd, groupCD[i]);
                }
                else
                {
                    //dsTmp = accResult.selectOverallPMSPYearOAP(yearMonthEnd, groupCD[i]);
                    DataSet dsPMSPQty = accResult.selectOverallPMSPQtyYearOAP(yearMonthEnd);
                    DataSet dsPMSPUp = accResult.selectOverallPMSPUnitPriceYearOAP(yearMonthEnd);
                    dsTmp = this.calculateOverallPMSPSaleYearOAP(dsPMSPQty, dsPMSPUp);
                }


                dsOverallYearOAP.Merge(dsTmp);
            }

            // ACC ACTAUL
            for (int i = 0; i < groupCD.Length; i++)
            {
                dsTmp = null;
                dsTmp = accResult.selectOverAllActual(yearMonthBegin, yearMonthEnd, groupCD[i]);

                dsAccActualResult.Merge(dsTmp);
            }

            // ACC OAP PLAN

            for (int i = 0; i < groupCD.Length; i++)
            {
                dsTmp = null;
                //dsTmp = accResult.selectAccOverallOAP(yearMonthBegin, yearMonthEnd, groupCD[i]);


                if (i != 1)
                {
                    dsTmp = accResult.selectAccOverallOAP(yearMonthBegin, yearMonthEnd, groupCD[i]);
                }
                else
                {
                    //dsTmp = accResult.selectAccOverallPMSPOAP(yearMonthBegin, yearMonthEnd, groupCD[i]);
                    DataSet dsAccPmspQty = accResult.selectAccOverallPMSPQtyOAP(yearMonthBegin, yearMonthEnd, groupCD[i]);
                    DataSet dsAccPmspUp = accResult.selectOverallPMSPUnitPriceYearOAP(yearMonthEnd);
                    dsTmp = this.calculateAccumulatePMSPSaleYearOAP(dsAccPmspQty, dsAccPmspUp);
                }


                dsAccumulateOAP.Merge(dsTmp);
            }
        }

        private DataSet calculateOverallPMSPSaleYearOAP(DataSet dsPmspQty, DataSet dsPmspUp)
        {
            try
            {
                DataTable dtTmp = new DataTable();
                dtTmp = dsOverallYearOAP.Tables[0].Clone();
                DataRow drTmp = dtTmp.NewRow();

                drTmp["GROUPCD"] = "PMSP";
                drTmp["I_YEAR_MNTH"] = dsPmspQty.Tables[0].Rows[0]["i_year_mnth"].ToString();

                double sumAmt = 0.0;
                string strCmd = "";
                int qty = 0;
                double up = 0.0;

                for (int i = 0; i < dsPmspQty.Tables[0].Rows.Count; i++)
                {
                    strCmd = "";
                    strCmd = "trim(i_item_cd) = '" + dsPmspQty.Tables[0].Rows[i]["i_item_cd"].ToString().Trim() + "'";

                    DataRow[] dr = dsPmspUp.Tables[0].Select(strCmd);
                    qty = Convert.ToInt32(dsPmspQty.Tables[0].Rows[i]["i_qty"].ToString());

                    if (dr != null)
                    {
                        up = Convert.ToDouble(dsPmspUp.Tables[0].Rows[i]["i_up"].ToString());
                    }
                    else
                    {
                        up = 0.0;
                    }

                    sumAmt += (qty * up);
                }

                drTmp["I_AMT"] = sumAmt;

                dtTmp.Rows.Add(drTmp);

                DataSet dsTmp = new DataSet();
                dsTmp.Tables.Add(dtTmp);

                return dsTmp;
            }
            catch
            {
                return null;
            }
        }

        private DataSet calculateAccumulatePMSPSaleYearOAP(DataSet dsPmspQty, DataSet dsPmspUp)
        {
            try
            {
                DataTable dtTmp = new DataTable();
                dtTmp = dsAccumulateOAP.Tables[0].Clone();
                DataRow drTmp = dtTmp.NewRow();

                drTmp["GROUPCD"] = "PMSP";
                drTmp["I_YEAR_MNTH"] = dsPmspQty.Tables[0].Rows[0]["i_year_mnth"].ToString();

                double sumAmt = 0.0;
                string strCmd = "";
                int qty = 0;
                double up = 0.0;

                for (int i = 0; i < dsPmspQty.Tables[0].Rows.Count; i++)
                {
                    strCmd = "";
                    strCmd = "trim(i_item_cd) = '" + dsPmspQty.Tables[0].Rows[i]["i_item_cd"].ToString().Trim() + "'";

                    DataRow[] dr = dsPmspUp.Tables[0].Select(strCmd);
                    qty = Convert.ToInt32(dsPmspQty.Tables[0].Rows[i]["i_qty"].ToString());

                    if (dr != null)
                    {
                        up = Convert.ToDouble(dsPmspUp.Tables[0].Rows[i]["i_up"].ToString());
                    }
                    else
                    {
                        up = 0.0;
                    }

                    sumAmt += (qty * up);
                }

                drTmp["I_AMT"] = sumAmt;

                dtTmp.Rows.Add(drTmp);

                DataSet dsTmp = new DataSet();
                dsTmp.Tables.Add(dtTmp);

                return dsTmp;
            }
            catch
            {
                return null;
            }
        }

        private void report5(string yearMonthBegin, string yearMonthEnd)
        {
            DataSet dsTmp = null;

            // ALL OAP
            //string[] groupCD = { "OTHER", "NC Data Sales", "492B", "Hiace", "OEM", "640A" };
            //string[] projectDesc = { "", "", "'PJ09'", "'PJ03', 'PJ10'", "", "'PJ05', 'PJ08'" };
            string[] groupCD = { "UAM NG Scrap Sales", "NC Data Sales", "OTHER", "492B", "Hiace", "OEM", "640A" };
            string[] projectDesc = { "", "", "", "'PJ09'", "'PJ03', 'PJ10'", "", "'PJ05', 'PJ08'" };

            for (int i = 0; i < groupCD.Length; i++)
            {
                dsTmp = null;

                if (i == 2)
                {
                    dsTmp = accResult.selectOverallYearOAP(yearMonthEnd, "OTHER");
                }
                else if (i == 1)
                {
                    dsTmp = accResult.selectProjecjOEMNCDataOAP(yearMonthEnd);
                }
                else if (i == 5)
                {
                    dsTmp = accResult.selectProjectOEMOEMPartYearOAP(yearMonthEnd);
                }
                else if (i == 0)
                {
                    dsTmp = accResult.selectProjecjOEMUAMOAP(yearMonthEnd);
                }
                else
                {
                    dsTmp = accResult.selectProjectOEMYearOAP(yearMonthEnd, groupCD[i], projectDesc[i]);
                }

                dsProjOEMYearOAP.Merge(dsTmp);
            }

            // ACC ACTAUL
            for (int i = 0; i < groupCD.Length; i++)
            {
                dsTmp = null;
                if (i == 2)
                {
                    dsTmp = accResult.selectOverAllActual(yearMonthBegin, yearMonthEnd, "OTHER");
                }
                else if (i == 1)
                {
                    dsTmp = accResult.selectProjecjOEMNCDataOAP(yearMonthEnd);
                }
                else if (i == 5)
                {
                    dsTmp = accResult.selectProjectOEMOEMPartActual(yearMonthBegin, yearMonthEnd);
                }
                else if (i == 0)
                {
                    dsTmp = accResult.selectProjectOEMUAMActual(yearMonthBegin, yearMonthEnd, groupCD[i]);
                }
                else
                {
                    dsTmp = accResult.selectProjectOEMActual(yearMonthBegin, yearMonthEnd, groupCD[i], projectDesc[i]);
                }

                dsProjOEMAccActualResult.Merge(dsTmp);
            }

            // ACC OAP PLAN

            for (int i = 0; i < groupCD.Length; i++)
            {
                dsTmp = null;
                //dsTmp = accResult.selectAccOverallOAP(yearMonthBegin, yearMonthEnd, groupCD[i]);


                if (i == 0)
                {
                    dsTmp = accResult.selectProjecjOEMUAMOAP(yearMonthEnd);
                }
                else if (i == 1)
                {
                    dsTmp = accResult.selectProjecjOEMNCDataOAP(yearMonthEnd);
                }
                else if (i == 2)
                {
                    dsTmp = accResult.selectAccProjectOEMOtherOAP(yearMonthBegin, yearMonthEnd, groupCD[i]);
                }
                else if (i == 5)
                {
                    dsTmp = accResult.selectAccProjectOEMOEMPartOAP(yearMonthBegin, yearMonthEnd, groupCD[i]);
                }
                else
                {
                    dsTmp = accResult.selectAccProjectOEMOAP(yearMonthBegin, yearMonthEnd, groupCD[i], projectDesc[i]);
                }

                dsProjOEMAccumulateOAP.Merge(dsTmp);
            }

        }

        private void report6(string yearMonthBegin, string yearMonthEnd)
        {
            DataSet dsTmp = new DataSet();
            DataTable dt1 = new DataTable();
            dt1.Columns.Add("GROUPCD", typeof(String));
            dt1.Columns.Add("I_YEAR_MNTH", typeof(String));
            dt1.Columns.Add("I_AMT", typeof(Double));

            dsTmp.Tables.Add(dt1);

            // ALL OAP
            DataSet dsPMSPQty = accResult.selectOverallPMSPQtyYearOAP(yearMonthEnd);
            DataSet dsPMSPUp = accResult.selectOverallPMSPUnitPriceYearOAP(yearMonthEnd);
            dsTmp = this.calculateProjectPMSPSaleYearOAP(dsPMSPQty, dsPMSPUp, dsTmp);
            dsProjPMSPYearOAP.Merge(dsTmp);

            // ACC ACTAUL
            dsTmp = null;
            dsTmp = accResult.selectOverAllActual(yearMonthBegin, yearMonthEnd, "PMSP");
            dsProjPMSPAccActualResult.Merge(dsTmp);

            // ACC OAP PLAN
            dsTmp = null;
            dsTmp = new DataSet();

            DataTable dt2 = new DataTable();
            dt2.Columns.Add("GROUPCD", typeof(String));
            dt2.Columns.Add("I_YEAR_MNTH", typeof(String));
            dt2.Columns.Add("I_AMT", typeof(Double));

            dsTmp.Tables.Add(dt2);

            DataSet dsAccPmspQty = accResult.selectAccOverallPMSPQtyOAP(yearMonthBegin, yearMonthEnd, "PMSP");
            DataSet dsAccPmspUp = accResult.selectOverallPMSPUnitPriceYearOAP(yearMonthEnd);
            dsTmp = this.calculateProjectPMSPSaleAccOAP(dsAccPmspQty, dsAccPmspUp, dsTmp);
            dsProjPMSPAccumulateOAP.Merge(dsTmp);
        }

        private DataSet calculateProjectPMSPSaleYearOAP(DataSet dsPmspQty, DataSet dsPmspUp, DataSet dsMaster)
        {
            try
            {
                DataTable dtTmp = new DataTable();
                dtTmp = dsMaster.Tables[0].Clone();
                DataRow drTmp = dtTmp.NewRow();

                drTmp["GROUPCD"] = "PMSP";
                drTmp["I_YEAR_MNTH"] = dsPmspQty.Tables[0].Rows[0]["i_year_mnth"].ToString();

                double sumAmt = 0.0;
                string strCmd = "";
                int qty = 0;
                double up = 0.0;

                for (int i = 0; i < dsPmspQty.Tables[0].Rows.Count; i++)
                {
                    strCmd = "";
                    strCmd = "trim(i_item_cd) = '" + dsPmspQty.Tables[0].Rows[i]["i_item_cd"].ToString().Trim() + "'";

                    DataRow[] dr = dsPmspUp.Tables[0].Select(strCmd);
                    qty = Convert.ToInt32(dsPmspQty.Tables[0].Rows[i]["i_qty"].ToString());

                    if (dr != null)
                    {
                        up = Convert.ToDouble(dsPmspUp.Tables[0].Rows[i]["i_up"].ToString());
                    }
                    else
                    {
                        up = 0.0;
                    }

                    sumAmt += (qty * up);
                }

                drTmp["I_AMT"] = sumAmt;

                dtTmp.Rows.Add(drTmp);

                DataSet dsTmp = new DataSet();
                dsTmp.Tables.Add(dtTmp);

                return dsTmp;
            }
            catch
            {
                return null;
            }
        }

        private DataSet calculateProjectPMSPSaleAccOAP(DataSet dsPmspQty, DataSet dsPmspUp, DataSet dsMaster)
        {
            try
            {
                DataTable dtTmp = new DataTable();
                dtTmp = dsMaster.Tables[0].Clone();
                DataRow drTmp = dtTmp.NewRow();

                drTmp["GROUPCD"] = "PMSP";
                drTmp["I_YEAR_MNTH"] = dsPmspQty.Tables[0].Rows[0]["i_year_mnth"].ToString();

                double sumAmt = 0.0;
                string strCmd = "";
                int qty = 0;
                double up = 0.0;

                for (int i = 0; i < dsPmspQty.Tables[0].Rows.Count; i++)
                {
                    strCmd = "";
                    strCmd = "trim(i_item_cd) = '" + dsPmspQty.Tables[0].Rows[i]["i_item_cd"].ToString().Trim() + "'";

                    DataRow[] dr = dsPmspUp.Tables[0].Select(strCmd);
                    qty = Convert.ToInt32(dsPmspQty.Tables[0].Rows[i]["i_qty"].ToString());

                    if (dr != null)
                    {
                        up = Convert.ToDouble(dsPmspUp.Tables[0].Rows[i]["i_up"].ToString());
                    }
                    else
                    {
                        up = 0.0;
                    }

                    sumAmt += (qty * up);
                }

                drTmp["I_AMT"] = sumAmt;

                dtTmp.Rows.Add(drTmp);

                DataSet dsTmp = new DataSet();
                dsTmp.Tables.Add(dtTmp);

                return dsTmp;
            }
            catch
            {
                return null;
            }
        }

        #endregion

        private void processOverallSalesForecast()
        {
            string yearMonthPrev = dtpYearMonth.Value.AddMonths(-1).ToString("yyyyMM");
            string yearMonth = dtpYearMonth.Value.ToString("yyyyMM");
            string year = dtpYearMonth.Value.ToString("yyyy");
            int month = dtpYearMonth.Value.Month;
            int prevMonth = dtpYearMonth.Value.AddMonths(-1).Month;

            report7(year, yearMonth, yearMonthPrev, month, prevMonth);
            report8(month, year, yearMonth, yearMonthPrev);
            report9(year, yearMonth, yearMonthPrev, month, prevMonth);
        }

        #region OVERALL MONTHLY SALES FORECAST

        private void report7(string year, string yearMonth, string yearMonthPrev, int month, int prevMonth)
        {
            // PREV FORECAST
            DataSet dsTmp = new DataSet();
            dsTmp.Tables.Add(this.createTableSaleForecast());
            DataSet dsFrstQty = saleForecast.selectForecastQty(yearMonthPrev, "ALL", "ALL");
            DataSet dsFrstUp = saleForecast.selectForecastUnitPrice(yearMonthPrev, "ALL", "ALL");

            dsTmp = this.calculateOverallSaleForecast("Previous Forecast", dsFrstQty, dsFrstUp, dsTmp, prevMonth);
            dsOverallSaleForecat.Merge(dsTmp);

            // LAST FORECAST
            dsTmp = null;
            dsTmp = new DataSet();
            dsTmp.Tables.Add(this.createTableSaleForecast());
            dsFrstQty = null;
            dsFrstUp = null;
            dsFrstQty = saleForecast.selectForecastQty(yearMonth, "ALL", "ALL");
            dsFrstUp = saleForecast.selectForecastUnitPrice(yearMonth, "ALL", "ALL");
            dsTmp = this.calculateOverallSaleForecast("Lastest Forecast", dsFrstQty, dsFrstUp, dsTmp, month);
            dsOverallSaleForecat.Merge(dsTmp);

            // ACTUAL RESULT
            dsTmp = null;
            dsTmp = new DataSet();
            dsTmp.Tables.Add(this.createTableSaleForecast());
            dsTmp = this.calculateActualResult("Actual Result", month, year, yearMonth, dsTmp, "ALL", "PMSP");
            dsOverallSaleForecat.Merge(dsTmp);

            // ANNUAL PLAN OAP
            dsTmp = null;
            dsTmp = new DataSet();
            dsTmp.Tables.Add(this.createTableSaleForecast());

            dsFrstQty = null;
            dsFrstQty = saleForecast.selectMonthlyOAP(year, "ALL");
            dsTmp = this.calculateAnnualOAP("Annual Plan OAP", dsFrstQty, dsFrstUp, dsTmp);
            dsOverallSaleForecat.Merge(dsTmp);
        }

        private DataSet calculateOverallSaleForecast(string type, DataSet dsPmspQty, DataSet dsPmspUp, DataSet dsMaster, int month)
        {
            try
            {
                DataTable dtTmp = new DataTable();
                dtTmp = dsMaster.Tables[0].Clone();
                DataRow drTmp = dtTmp.NewRow();

                drTmp["TYPE"] = type;
                for (int i = 1; i <= 12; i++)
                {
                    drTmp["M" + (i).ToString()] = 0.0;
                }

                string strCmd = "";
                int[] qty = { 0, 0, 0, 0, 0, 0 };
                double up = 0.0;

                for (int i = 0; i < dsPmspQty.Tables[0].Rows.Count; i++)
                {
                    strCmd = "";
                    strCmd = "trim(i_item_cd) = '" + dsPmspQty.Tables[0].Rows[i]["i_item_cd"].ToString().Trim() + "'";

                    DataRow[] dr = dsPmspUp.Tables[0].Select(strCmd);
                    qty[0] = Convert.ToInt32(dsPmspQty.Tables[0].Rows[i]["qty_m1"].ToString());
                    qty[1] = Convert.ToInt32(dsPmspQty.Tables[0].Rows[i]["qty_m2"].ToString());
                    qty[2] = Convert.ToInt32(dsPmspQty.Tables[0].Rows[i]["qty_m3"].ToString());
                    qty[3] = Convert.ToInt32(dsPmspQty.Tables[0].Rows[i]["qty_m4"].ToString());
                    qty[4] = Convert.ToInt32(dsPmspQty.Tables[0].Rows[i]["qty_m5"].ToString());
                    qty[5] = Convert.ToInt32(dsPmspQty.Tables[0].Rows[i]["qty_m6"].ToString());

                    if (dr.Count() > 0)
                    {
                        up = Convert.ToDouble(dr[0]["i_up"].ToString());
                    }
                    else
                    {
                        up = 0.0;
                    }

                    for (int j = 0; j < 6; j++)
                    {
                        int index = month - 3 + j;
                        double amtTmp = Convert.ToDouble(drTmp["M" + (index).ToString()].ToString());
                        drTmp["M" + (index).ToString()] = amtTmp + (qty[j] * up);
                    }
                }

                dtTmp.Rows.Add(drTmp);

                DataSet dsTmp = new DataSet();
                dsTmp.Tables.Add(dtTmp);

                return dsTmp;
            }
            catch
            {
                return null;
            }
        }

        private DataSet calculateActualResult(string type, int month, string year, string yearMonth, DataSet dsMaster, string typeSearch, string dlCode)
        {
            try
            {
                DataTable dtTmp = new DataTable();
                dtTmp = dsMaster.Tables[0].Clone();
                DataRow drTmp = dtTmp.NewRow();

                drTmp["TYPE"] = type;
                for (int i = 1; i <= 12; i++)
                {
                    drTmp["M" + (i).ToString()] = 0.0;
                }

                for (int i = 0; i < month - 3; i++)
                {
                    DataSet dsActual = saleForecast.selectActualResult(year + (i + 4).ToString("00"), typeSearch, dlCode);
                    drTmp["M" + (i + 1).ToString()] = Convert.ToDouble(dsActual.Tables[0].Rows[0]["i_amt"].ToString());
                }

                dtTmp.Rows.Add(drTmp);

                DataSet dsTmp = new DataSet();
                dsTmp.Tables.Add(dtTmp);

                return dsTmp;
            }
            catch
            {
                return null;
            }
        }

        private DataSet calculateAnnualOAP(string type, DataSet dsPmspQty, DataSet dsPmspUp, DataSet dsMaster)
        {
            try
            {
                DataTable dtTmp = new DataTable();
                dtTmp = dsMaster.Tables[0].Clone();
                DataRow drTmp = dtTmp.NewRow();

                drTmp["TYPE"] = type;
                for (int i = 1; i <= 12; i++)
                {
                    drTmp["M" + (i).ToString()] = 0.0;
                }

                int[] qty = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
                string strCmd = "";
                double up = 0.0;
                for (int i = 0; i < dsPmspQty.Tables[0].Rows.Count; i++)
                {
                    strCmd = "";
                    strCmd = "trim(i_item_cd) = '" + dsPmspQty.Tables[0].Rows[i]["i_item_cd"].ToString().Trim() + "'";

                    DataRow[] dr = dsPmspUp.Tables[0].Select(strCmd);

                    for (int j = 1; j <= 12; j++)
                    {
                        qty[j - 1] = Convert.ToInt32(dsPmspQty.Tables[0].Rows[i]["i_mnth_plan_qty" + (j).ToString()].ToString());
                    }

                    if (dr.Count() > 0)
                    {
                        up = Convert.ToDouble(dr[0]["i_up"].ToString());
                    }
                    else
                    {
                        up = 0.0;
                    }

                    for (int j = 1; j <= 12; j++)
                    {
                        double amtTmp = Convert.ToDouble(drTmp["M" + (j).ToString()].ToString());
                        drTmp["M" + (j).ToString()] = (amtTmp + (qty[j - 1] * up));
                    }
                }

                dtTmp.Rows.Add(drTmp);

                DataSet dsTmp = new DataSet();
                dsTmp.Tables.Add(dtTmp);

                return dsTmp;
            }
            catch
            {
                return null;
            }
        }

        private void report8(int month, string year, string yearMonth, string yearMonthPrev)
        {
            string[] groupDesc = { "NC Data Sales", "OTHER", "OEM", "492B", "HIACE", "640A", "UAM NG Scrap Sales", "OAP" };
            string[] dlGroup = { "''", "'8-1-005-3'", "'8-1-002-1'", "'8-1-023-3'", "'8-1-004-3'", "'8-1-001-3', '8-1-011-1', '8-1-010-1', '8-1-002-3', '8-1-009-1'", "''", "''" };
            DataSet dsTmp = new DataSet();
            DataSet dsFrstQty = null;
            DataSet dsFrstUp = null;

            // Prev of Prev Month
            dsTmp.Tables.Add(this.createTableOEMSaleForecast());

            for (int i = 0; i < groupDesc.Length; i++)
            {
                DataRow dr = dsTmp.Tables[0].NewRow();
                dr["type"] = groupDesc[i];
                dr["i_amt_prev_month"] = 0.0;

                dsTmp.Tables[0].Rows.Add(dr);
            }

            dsOEMSaleForecast.Merge(dsTmp);

            // Prev Month
            for (int i = 0; i < groupDesc.Length; i++)
            {
                dsFrstQty = null;
                dsFrstUp = null;
                dsFrstQty = saleForecast.selectForecastQty(yearMonthPrev, "OEM", dlGroup[i]);
                dsFrstUp = saleForecast.selectForecastUnitPrice(yearMonthPrev, "OEM", dlGroup[i]);

                DataSet ds = this.calculateOEMSalePrevForecast(groupDesc[i], dsFrstQty, dsFrstUp, dsTmp);
                updateCelData("prev", ds);
                //dsOEMSaleForecast.Merge(dsTmp);
            }

            // Month
            for (int i = 0; i < groupDesc.Length; i++)
            {
                dsFrstQty = null;
                dsFrstUp = null;
                dsFrstQty = saleForecast.selectForecastQty(yearMonth, "OEM", dlGroup[i]);
                dsFrstUp = saleForecast.selectForecastUnitPrice(yearMonth, "OEM", dlGroup[i]);

                DataSet ds = this.calculateOEMSaleForecast(groupDesc[i], dsFrstQty, dsFrstUp, dsTmp);
                updateCelData("month", ds);
                //dsOEMSaleForecast.Merge(ds);
            }

            // Actual
            for (int i = 0; i < groupDesc.Length; i++)
            {
                int prevMonth = month - 1;
                DataSet dsResult = new DataSet();
                for (int j = 0; j < 6; j++)
                {
                    DataSet dsActual = null;
                    if (j == 0)
                    {
                        dsActual = saleForecast.selectActualResult(year + (prevMonth).ToString("00"), "OEM", dlGroup[i]);
                    }
                    else
                    {
                        dsActual = saleForecast.selectActualResult(year + (month + (j - 1)).ToString("00"), "OEM", dlGroup[i]);
                    }

                    DataSet ds = this.calculateOEMActualResult(groupDesc[i], j, dsActual, dsTmp);

                    updateCelData("actual", ds);
                    //dsOEMSaleForecast.Merge(ds);
                }
            }

            // OAP
            dsFrstQty = null;
            dsFrstUp = null;
            dsFrstQty = saleForecast.selectMonthlyOAP(year, "ALL");
            dsFrstUp = saleForecast.selectForecastUnitPrice(yearMonthPrev, "ALL", "ALL");

            DataSet ds1 = this.calculateOEMAnnualOAP("OAP", dsFrstQty, dsFrstUp, dsTmp, month);
            updateCelData("oap", ds1);
            //dsOEMSaleForecast.Merge(ds1);
        }

        private void updateCelData(string process, DataSet dsData)
        {
            for (int i = 0; i < dsData.Tables[0].Rows.Count; i++)
            {
                DataRow dr = dsOEMSaleForecast.Tables[0].Select("type='" + dsData.Tables[0].Rows[i]["type"].ToString() + "'").FirstOrDefault();
                if (dr != null)
                {
                    if (process == "prev")
                    {
                        dr["i_amt_last_month"] = dsData.Tables[0].Rows[i]["i_amt_last_month"];
                        dr["i_amt_prev_month_n"] = dsData.Tables[0].Rows[i]["i_amt_prev_month_n"];
                        dr["i_amt_prev_month_n1"] = dsData.Tables[0].Rows[i]["i_amt_prev_month_n1"];
                        dr["i_amt_prev_month_n2"] = dsData.Tables[0].Rows[i]["i_amt_prev_month_n2"];
                        dr["i_amt_prev_month_n3"] = dsData.Tables[0].Rows[i]["i_amt_prev_month_n3"];
                        dr["i_amt_prev_month_n4"] = dsData.Tables[0].Rows[i]["i_amt_prev_month_n4"];
                    }
                    else if (process == "month")
                    {
                        dr["i_amt_last_month_n"] = dsData.Tables[0].Rows[i]["i_amt_last_month_n"];
                        dr["i_amt_last_month_n1"] = dsData.Tables[0].Rows[i]["i_amt_last_month_n1"];
                        dr["i_amt_last_month_n2"] = dsData.Tables[0].Rows[i]["i_amt_last_month_n2"];
                        dr["i_amt_last_month_n3"] = dsData.Tables[0].Rows[i]["i_amt_last_month_n3"];
                        dr["i_amt_last_month_n4"] = dsData.Tables[0].Rows[i]["i_amt_last_month_n4"];
                    }
                    else if (process == "actual")
                    {
                        if (dsData.Tables[0].Rows[i]["i_amt_actual_month"].ToString().Trim().Length > 0)
                            dr["i_amt_actual_month"] = dsData.Tables[0].Rows[i]["i_amt_actual_month"];
                        if (dsData.Tables[0].Rows[i]["i_amt_actual_month_n"].ToString().Trim().Length > 0)
                            dr["i_amt_actual_month_n"] = dsData.Tables[0].Rows[i]["i_amt_actual_month_n"];
                        if (dsData.Tables[0].Rows[i]["i_amt_actual_month_n1"].ToString().Trim().Length > 0)
                            dr["i_amt_actual_month_n1"] = dsData.Tables[0].Rows[i]["i_amt_actual_month_n1"];
                        if (dsData.Tables[0].Rows[i]["i_amt_actual_month_n2"].ToString().Trim().Length > 0)
                            dr["i_amt_actual_month_n2"] = dsData.Tables[0].Rows[i]["i_amt_actual_month_n2"];
                        if (dsData.Tables[0].Rows[i]["i_amt_actual_month_n3"].ToString().Trim().Length > 0)
                            dr["i_amt_actual_month_n3"] = dsData.Tables[0].Rows[i]["i_amt_actual_month_n3"];
                        if (dsData.Tables[0].Rows[i]["i_amt_actual_month_n4"].ToString().Trim().Length > 0)
                            dr["i_amt_actual_month_n4"] = dsData.Tables[0].Rows[i]["i_amt_actual_month_n4"];

                    }
                    else if (process == "oap")
                    {

                        dr["i_amt_prev_month"] = dsData.Tables[0].Rows[i]["i_amt_prev_month"];
                        dr["i_amt_last_month"] = dsData.Tables[0].Rows[i]["i_amt_last_month"];
                        dr["i_amt_actual_month"] = dsData.Tables[0].Rows[i]["i_amt_actual_month"];

                        dr["i_amt_prev_month_n"] = dsData.Tables[0].Rows[i]["i_amt_prev_month_n"];
                        dr["i_amt_last_month_n"] = dsData.Tables[0].Rows[i]["i_amt_last_month_n"];
                        dr["i_amt_actual_month_n"] = dsData.Tables[0].Rows[i]["i_amt_actual_month_n"];

                        dr["i_amt_prev_month_n1"] = dsData.Tables[0].Rows[i]["i_amt_prev_month_n1"];
                        dr["i_amt_last_month_n1"] = dsData.Tables[0].Rows[i]["i_amt_last_month_n1"];
                        dr["i_amt_actual_month_n1"] = dsData.Tables[0].Rows[i]["i_amt_actual_month_n1"];

                        dr["i_amt_prev_month_n2"] = dsData.Tables[0].Rows[i]["i_amt_prev_month_n2"];
                        dr["i_amt_last_month_n2"] = dsData.Tables[0].Rows[i]["i_amt_last_month_n2"];
                        dr["i_amt_actual_month_n2"] = dsData.Tables[0].Rows[i]["i_amt_actual_month_n2"];

                        dr["i_amt_prev_month_n3"] = dsData.Tables[0].Rows[i]["i_amt_prev_month_n3"];
                        dr["i_amt_last_month_n3"] = dsData.Tables[0].Rows[i]["i_amt_last_month_n3"];
                        dr["i_amt_actual_month_n3"] = dsData.Tables[0].Rows[i]["i_amt_actual_month_n3"];

                        dr["i_amt_prev_month_n4"] = dsData.Tables[0].Rows[i]["i_amt_prev_month_n4"];
                        dr["i_amt_last_month_n4"] = dsData.Tables[0].Rows[i]["i_amt_last_month_n4"];
                        dr["i_amt_actual_month_n4"] = dsData.Tables[0].Rows[i]["i_amt_actual_month_n4"];
                    }
                }
            }
        }

        private DataTable createTableOEMSaleForecast()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("type", typeof(String));
            dt.Columns.Add("i_amt_prev_month", typeof(Double));
            dt.Columns.Add("i_amt_last_month", typeof(Double));
            dt.Columns.Add("i_amt_actual_month", typeof(Double));
            dt.Columns.Add("i_amt_prev_month_n", typeof(Double));
            dt.Columns.Add("i_amt_last_month_n", typeof(Double));
            dt.Columns.Add("i_amt_actual_month_n", typeof(Double));
            dt.Columns.Add("i_amt_prev_month_n1", typeof(Double));
            dt.Columns.Add("i_amt_last_month_n1", typeof(Double));
            dt.Columns.Add("i_amt_actual_month_n1", typeof(Double));
            dt.Columns.Add("i_amt_prev_month_n2", typeof(Double));
            dt.Columns.Add("i_amt_last_month_n2", typeof(Double));
            dt.Columns.Add("i_amt_actual_month_n2", typeof(Double));
            dt.Columns.Add("i_amt_prev_month_n3", typeof(Double));
            dt.Columns.Add("i_amt_last_month_n3", typeof(Double));
            dt.Columns.Add("i_amt_actual_month_n3", typeof(Double));
            dt.Columns.Add("i_amt_prev_month_n4", typeof(Double));
            dt.Columns.Add("i_amt_last_month_n4", typeof(Double));
            dt.Columns.Add("i_amt_actual_month_n4", typeof(Double));

            return dt;
        }

        private DataSet calculateOEMSalePrevForecast(string type, DataSet dsPmspQty, DataSet dsPmspUp, DataSet dsMaster)
        {
            try
            {
                DataTable dtTmp = new DataTable();
                dtTmp = dsMaster.Tables[0].Clone();
                DataRow drTmp = dtTmp.NewRow();

                drTmp["TYPE"] = type;
                for (int i = 0; i < 6; i++)
                {
                    if (i == 0)
                    {
                        drTmp["i_amt_last_month"] = 0.0;
                    }
                    else if (i == 1)
                    {
                        drTmp["i_amt_prev_month_n"] = 0.0;
                    }
                    else
                    {
                        drTmp["i_amt_prev_month_n" + (i - 1).ToString()] = 0.0;
                    }
                }

                string strCmd = "";
                double[] amt = { 0.0, 0.0, 0.0, 0.0, 0.0, 0.0 };
                double up = 0.0;

                for (int i = 0; i < dsPmspQty.Tables[0].Rows.Count; i++)
                {
                    strCmd = "";
                    strCmd = "trim(i_item_cd) = '" + dsPmspQty.Tables[0].Rows[i]["i_item_cd"].ToString().Trim() + "'";

                    DataRow[] dr = dsPmspUp.Tables[0].Select(strCmd);
                    if (dr.Count() > 0)
                    {
                        up = Convert.ToDouble(dr[0]["i_up"].ToString());
                    }
                    else
                    {
                        up = 0.0;
                    }

                    amt[0] += (Convert.ToInt32(dsPmspQty.Tables[0].Rows[i]["qty_m1"].ToString()) * up);
                    amt[1] += (Convert.ToInt32(dsPmspQty.Tables[0].Rows[i]["qty_m2"].ToString()) * up);
                    amt[2] += (Convert.ToInt32(dsPmspQty.Tables[0].Rows[i]["qty_m3"].ToString()) * up);
                    amt[3] += (Convert.ToInt32(dsPmspQty.Tables[0].Rows[i]["qty_m4"].ToString()) * up);
                    amt[4] += (Convert.ToInt32(dsPmspQty.Tables[0].Rows[i]["qty_m5"].ToString()) * up);
                    amt[5] += (Convert.ToInt32(dsPmspQty.Tables[0].Rows[i]["qty_m6"].ToString()) * up);
                }

                for (int j = 0; j < 6; j++)
                {
                    if (j == 0)
                    {
                        drTmp["i_amt_last_month"] = amt[j];
                    }
                    else if (j == 1)
                    {
                        drTmp["i_amt_prev_month_n"] = amt[j];
                    }
                    else
                    {
                        drTmp["i_amt_prev_month_n" + (j - 1).ToString()] = amt[j];
                    }
                }

                dtTmp.Rows.Add(drTmp);

                DataSet dsTmp = new DataSet();
                dsTmp.Tables.Add(dtTmp);

                return dsTmp;
            }
            catch
            {
                return null;
            }
        }

        private DataSet calculateOEMSaleForecast(string type, DataSet dsPmspQty, DataSet dsPmspUp, DataSet dsMaster)
        {
            try
            {
                DataTable dtTmp = new DataTable();
                dtTmp = dsMaster.Tables[0].Clone();
                DataRow drTmp = dtTmp.NewRow();

                drTmp["TYPE"] = type;
                for (int i = 0; i < 5; i++)
                {
                    if (i == 0)
                    {
                        drTmp["i_amt_last_month_n"] = 0.0;
                    }
                    else
                    {
                        drTmp["i_amt_last_month_n" + (i).ToString()] = 0.0;
                    }
                }

                string strCmd = "";
                double[] amt = { 0.0, 0.0, 0.0, 0.0, 0.0 };
                double up = 0.0;

                for (int i = 0; i < dsPmspQty.Tables[0].Rows.Count; i++)
                {
                    strCmd = "";
                    strCmd = "trim(i_item_cd) = '" + dsPmspQty.Tables[0].Rows[i]["i_item_cd"].ToString().Trim() + "'";

                    DataRow[] dr = dsPmspUp.Tables[0].Select(strCmd);
                    if (dr.Count() > 0)
                    {
                        up = Convert.ToDouble(dr[0]["i_up"].ToString());
                    }
                    else
                    {
                        up = 0.0;
                    }

                    amt[0] += (Convert.ToInt32(dsPmspQty.Tables[0].Rows[i]["qty_m1"].ToString()) * up);
                    amt[1] += (Convert.ToInt32(dsPmspQty.Tables[0].Rows[i]["qty_m2"].ToString()) * up);
                    amt[2] += (Convert.ToInt32(dsPmspQty.Tables[0].Rows[i]["qty_m3"].ToString()) * up);
                    amt[3] += (Convert.ToInt32(dsPmspQty.Tables[0].Rows[i]["qty_m4"].ToString()) * up);
                    amt[4] += (Convert.ToInt32(dsPmspQty.Tables[0].Rows[i]["qty_m5"].ToString()) * up);
                }

                for (int j = 0; j < 5; j++)
                {
                    if (j == 0)
                    {
                        drTmp["i_amt_last_month_n"] = amt[j];
                    }
                    else
                    {
                        drTmp["i_amt_prev_month_n" + (j).ToString()] = amt[j];
                    }
                }

                dtTmp.Rows.Add(drTmp);

                DataSet dsTmp = new DataSet();
                dsTmp.Tables.Add(dtTmp);

                return dsTmp;
            }
            catch
            {
                return null;
            }
        }

        private DataSet calculateOEMActualResult(string type, int index, DataSet dsActual, DataSet dsMaster)
        {
            try
            {
                DataTable dtTmp = new DataTable();
                dtTmp = dsMaster.Tables[0].Clone();
                DataRow drTmp = dtTmp.NewRow();

                drTmp["TYPE"] = type;

                if (index == 0)
                {
                    try
                    {
                        drTmp["i_amt_actual_month"] = Convert.ToDouble(dsActual.Tables[0].Rows[0]["i_amt"].ToString());
                    }
                    catch
                    {
                        drTmp["i_amt_actual_month"] = 0.0;
                    }
                }
                else if (index == 1)
                {
                    try
                    {
                        drTmp["i_amt_actual_month_n"] = Convert.ToDouble(dsActual.Tables[0].Rows[0]["i_amt"].ToString());
                    }
                    catch
                    {
                        drTmp["i_amt_actual_month_n"] = 0.0;
                    }
                }
                else
                {
                    try
                    {
                        drTmp["i_amt_actual_month_n" + (index - 1).ToString()] = Convert.ToDouble(dsActual.Tables[0].Rows[0]["i_amt"].ToString());
                    }
                    catch
                    {
                        drTmp["i_amt_actual_month_n" + (index - 1).ToString()] = 0.0;
                    }
                }

                dtTmp.Rows.Add(drTmp);

                DataSet dsTmp = new DataSet();
                dsTmp.Tables.Add(dtTmp);

                return dsTmp;
            }
            catch
            {
                return null;
            }
        }

        private DataSet calculateOEMAnnualOAP(string type, DataSet dsPmspQty, DataSet dsPmspUp, DataSet dsMaster, int month)
        {
            try
            {
                int monthBegin = 0;
                if (month >= 4 && month <= 12)
                {
                    monthBegin = month - 4;
                }
                else if (month >= 1 && month <= 3)
                {
                    monthBegin = month + 9;
                }

                DataTable dtTmp = new DataTable();
                dtTmp = dsMaster.Tables[0].Clone();
                DataRow drTmp = dtTmp.NewRow();

                drTmp["TYPE"] = type;
                for (int i = 0; i < 6; i++)
                {
                    if (i == 0)
                    {
                        drTmp["i_amt_prev_month"] = 0.0;
                        drTmp["i_amt_last_month"] = 0.0;
                        drTmp["i_amt_actual_month"] = 0.0;
                    }
                    else if (i == 1)
                    {
                        drTmp["i_amt_prev_month_n"] = 0.0;
                        drTmp["i_amt_last_month_n"] = 0.0;
                        drTmp["i_amt_actual_month_n"] = 0.0;
                    }
                    else
                    {
                        drTmp["i_amt_prev_month_n" + (i - 1).ToString()] = 0.0;
                        drTmp["i_amt_last_month_n" + (i - 1).ToString()] = 0.0;
                        drTmp["i_amt_actual_month_n" + (i - 1).ToString()] = 0.0;
                    }
                }

                double[] qty = { 0.0, 0.0, 0.0, 0.0, 0.0, 0.0 };
                string strCmd = "";
                double up = 0.0;
                for (int i = 0; i < dsPmspQty.Tables[0].Rows.Count; i++)
                {
                    strCmd = "";
                    strCmd = "trim(i_item_cd) = '" + dsPmspQty.Tables[0].Rows[i]["i_item_cd"].ToString().Trim() + "'";

                    DataRow[] dr = dsPmspUp.Tables[0].Select(strCmd);
                    if (dr.Count() > 0)
                    {
                        up = Convert.ToDouble(dr[0]["i_up"].ToString());
                    }
                    else
                    {
                        up = 0.0;
                    }

                    for (int j = 0; j < 6; j++)
                    {
                        qty[j] += (Convert.ToInt32(dsPmspQty.Tables[0].Rows[i]["i_mnth_plan_qty" + (monthBegin + j).ToString()].ToString()) * up);
                    }
                }

                for (int j = 0; j < 6; j++)
                {
                    if (j == 0)
                    {
                        drTmp["i_amt_prev_month"] = qty[j];
                        drTmp["i_amt_last_month"] = qty[j];
                        drTmp["i_amt_actual_month"] = qty[j];
                    }
                    else if (j == 1)
                    {
                        drTmp["i_amt_prev_month_n"] = qty[j];
                        drTmp["i_amt_last_month_n"] = qty[j];
                        drTmp["i_amt_actual_month_n"] = qty[j];
                    }
                    else
                    {
                        drTmp["i_amt_prev_month_n" + (j - 1).ToString()] = qty[j];
                        drTmp["i_amt_last_month_n" + (j - 1).ToString()] = qty[j];
                        drTmp["i_amt_actual_month_n" + (j - 1).ToString()] = qty[j];
                    }
                }

                dtTmp.Rows.Add(drTmp);

                DataSet dsTmp = new DataSet();
                dsTmp.Tables.Add(dtTmp);

                return dsTmp;
            }
            catch
            {
                return null;
            }
        }

        private void report9(string year, string yearMonth, string yearMonthPrev, int month, int prevMonth)
        {
            // PREV FORECAST
            DataSet dsTmp = new DataSet();
            dsTmp.Tables.Add(this.createTableSaleForecast());
            DataSet dsFrstQty = saleForecast.selectForecastQty(yearMonthPrev, "PMSP", "PMSP");
            DataSet dsFrstUp = saleForecast.selectForecastUnitPrice(yearMonthPrev, "PMSP", "PMSP");

            dsTmp = this.calculateOverallSaleForecast("Previous Forecast", dsFrstQty, dsFrstUp, dsTmp, prevMonth);
            dsPMSPSaleForecast.Merge(dsTmp);

            // LAST FORECAST
            dsTmp = null;
            dsTmp = new DataSet();
            dsTmp.Tables.Add(this.createTableSaleForecast());
            dsFrstQty = null;
            dsFrstUp = null;
            dsFrstQty = saleForecast.selectForecastQty(yearMonth, "PMSP", "PMSP");
            dsFrstUp = saleForecast.selectForecastUnitPrice(yearMonth, "PMSP", "PMSP");

            dsTmp = this.calculateOverallSaleForecast("Lastest Forecast", dsFrstQty, dsFrstUp, dsTmp, month);
            dsPMSPSaleForecast.Merge(dsTmp);

            // ACTUAL RESULT
            dsTmp = null;
            dsTmp = new DataSet();
            dsTmp.Tables.Add(this.createTableSaleForecast());
            dsTmp = this.calculateActualResult("Actual Result", month, year, yearMonth, dsTmp, "PMSP", "PMSP");
            dsPMSPSaleForecast.Merge(dsTmp);

            // ANNUAL PLAN OAP
            dsTmp = null;
            dsTmp = new DataSet();
            dsTmp.Tables.Add(this.createTableSaleForecast());

            dsFrstQty = null;
            dsFrstQty = saleForecast.selectMonthlyOAP(year, "PMSP");
            dsTmp = this.calculateAnnualOAP("Annual Plan OAP", dsFrstQty, dsFrstUp, dsTmp);
            dsPMSPSaleForecast.Merge(dsTmp);
        }

        private DataTable createTableSaleForecast()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("type", typeof(String));
            dt.Columns.Add("m1", typeof(Double));
            dt.Columns.Add("m2", typeof(Double));
            dt.Columns.Add("m3", typeof(Double));
            dt.Columns.Add("m4", typeof(Double));
            dt.Columns.Add("m5", typeof(Double));
            dt.Columns.Add("m6", typeof(Double));
            dt.Columns.Add("m7", typeof(Double));
            dt.Columns.Add("m8", typeof(Double));
            dt.Columns.Add("m9", typeof(Double));
            dt.Columns.Add("m10", typeof(Double));
            dt.Columns.Add("m11", typeof(Double));
            dt.Columns.Add("m12", typeof(Double));

            return dt;
        }



        #endregion

        private void processAccumulateSalesForecast()
        {

            string yearMonth = dtpYearMonth.Value.ToString("yyyyMM");
            string year = dtpYearMonth.Value.ToString("yyyy");
            int month = dtpYearMonth.Value.Month;
            int prevMonth = dtpYearMonth.Value.AddMonths(-1).Month;
            string yearMonthPrev = year + month.ToString("00");

            report10(year, month, prevMonth);
            report11(year, month, prevMonth);
            report12(year, month, prevMonth);
        }

        #region ACCUMULATE FORECAST

        private void report10(string year, int month, int prevMonth)
        {
            // All OAP
            DataSet dsTmp = new DataSet();
            DataSet dsOAPQty = accForcast.selectAccQtyOverallOAP(year, "all");
            DataSet dsOAPUP = accForcast.selectAccUPOverallOAP(year);

            dsTmp.Tables.Add(this.createTableAccOverallOAP());
            dsTmp = this.calculateAccOverallOAP("OAP", dsOAPQty, dsOAPUP, dsTmp);
            dsAccOverallOAP.Merge(dsTmp);

            // PREV OAP
            dsTmp = null;
            dsTmp = new DataSet();
            dsOAPQty = null;

            dsOAPQty = accForcast.selectAccQtyOverallOAP(year, "prev");
            dsTmp.Tables.Add(this.createTableAccOverallOAP());
            dsTmp = this.calculateAccOverallOAP("Prev Year OAP", dsOAPQty, dsOAPUP, dsTmp);
            dsAccOAPPrevYear.Merge(dsTmp);

            // THIS OAP
            dsTmp = null;
            dsTmp = new DataSet();
            dsOAPQty = null;

            dsOAPQty = accForcast.selectAccQtyOverallOAP(year, "this");
            dsTmp.Tables.Add(this.createTableAccOverallOAP());
            dsTmp = this.calculateAccOverallOAP("This Year OAP", dsOAPQty, dsOAPUP, dsTmp);
            dsAccOAPThisYear.Merge(dsTmp);
        }

        private DataTable createTableAccOverallOAP()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("type", typeof(String));
            dt.Columns.Add("i_amt", typeof(Double));

            return dt;
        }

        private DataSet calculateAccOverallOAP(string type, DataSet dsFrstQty, DataSet dsFrstUP, DataSet dsMaster)
        {
            try
            {
                DataTable dt = new DataTable();
                dt = dsMaster.Tables[0].Clone();
                DataRow drTmp = dt.NewRow();

                drTmp["TYPE"] = type;

                int qty = 0;
                double up = 0.0;
                string strCmd = "";
                double sumAmt = 0.0;

                for (int i = 0; i < dsFrstQty.Tables[0].Rows.Count; i++)
                {
                    qty = Convert.ToInt32(dsFrstQty.Tables[0].Rows[i]["I_QTY"].ToString());

                    strCmd = "";
                    strCmd = "trim(i_item_cd) = '" + dsFrstQty.Tables[0].Rows[i]["I_ITEM_CD"].ToString() + "'";
                    DataRow[] dr = dsFrstUP.Tables[0].Select(strCmd);

                    if (dr.Count() > 0)
                    {
                        up = Convert.ToDouble(dr[0]["I_UP"].ToString());
                    }
                    else
                    {
                        up = 0.0;
                    }

                    sumAmt += (qty * up);
                }

                drTmp["I_AMT"] = sumAmt;
                dt.Rows.Add(drTmp);

                DataSet dsTmp = new DataSet();
                dsTmp.Tables.Add(dt);

                return dsTmp;
            }
            catch
            {
                return null;
            }
        }

        private void report11(string year, int month, int prevMonth)
        {

            DataSet dsTmp = new DataSet();

            string[] groupDesc = { "NC Data Sales", "OTHER", "492B", "HIACE", "OEM", "640A", "CPJT" };
            string[] dlCode = { "''", "'8-1-005-3'", "'8-1-023-3'", "'8-1-004-3'", "'8-1-002-1'", "'8-1-001-3', '8-1-011-1', '8-1-010-1', '8-1-002-3', '8-1-009-1'", "''" };
            DataSet dsOAPQty = null;
            DataSet dsOAPUP = accForcast.selectAccUPOverallOAP(year);

            for (int i = 0; i < groupDesc.Length; i++)
            {
                // OEM All OAP
                dsTmp = null;
                dsTmp = new DataSet();
                dsOAPQty = null;
                dsTmp.Tables.Add(this.createTableAccOverallOAP());

                dsOAPQty = accForcast.selectAccOEMQtyOAP(year, "all", dlCode[i]);
                dsTmp = this.calculateAccOverallOAP(groupDesc[i], dsOAPQty, dsOAPUP, dsTmp);
                dsOEMAccOverallOAP.Merge(dsTmp);

                // OEM PREV OAP
                dsTmp = null;
                dsTmp = new DataSet();
                dsOAPQty = null;
                dsTmp.Tables.Add(this.createTableAccOverallOAP());

                dsOAPQty = accForcast.selectAccOEMQtyOAP(year, "prev", dlCode[i]);
                dsTmp = this.calculateAccOverallOAP(groupDesc[i], dsOAPQty, dsOAPUP, dsTmp);
                dsOEMAccOAPPrevYear.Merge(dsTmp);

                // OEM THIS OAP
                dsTmp = null;
                dsTmp = new DataSet();
                dsOAPQty = null;
                dsTmp.Tables.Add(this.createTableAccOverallOAP());

                dsOAPQty = accForcast.selectAccOEMQtyOAP(year, "this", dlCode[i]);
                dsTmp = this.calculateAccOverallOAP(groupDesc[i], dsOAPQty, dsOAPUP, dsTmp);
                dsOEMAccOAPThisYear.Merge(dsTmp);
            }
        }

        private void report12(string year, int month, int prevMonth)
        {
            DataSet dsTmp = new DataSet();
            DataSet dsOAPQty = null;
            DataSet dsOAPUP = accForcast.selectAccUPOverallOAP(year);

            // PMSP ALL OAP
            dsTmp.Tables.Add(this.createTableAccOverallOAP());
            dsOAPQty = accForcast.selectAccOEMQtyOAP(year, "all", "'8-1-003-3'");
            dsTmp = this.calculateAccOverallOAP("PMSP", dsOAPQty, dsOAPUP, dsTmp);
            dsPMSPAccOverallOAP.Merge(dsTmp);

            // PMSP PREV OAP
            dsTmp = null;
            dsOAPQty = null;
            dsTmp = new DataSet();

            dsTmp.Tables.Add(this.createTableAccOverallOAP());
            dsOAPQty = accForcast.selectAccOEMQtyOAP(year, "prev", "'8-1-003-3'");
            dsTmp = this.calculateAccOverallOAP("PMSP", dsOAPQty, dsOAPUP, dsTmp);
            dsPMSPAccOAPPrevYear.Merge(dsTmp);

            // THIS PREV OAP
            dsTmp = null;
            dsOAPQty = null;
            dsTmp = new DataSet();

            dsTmp.Tables.Add(this.createTableAccOverallOAP());
            dsOAPQty = accForcast.selectAccOEMQtyOAP(year, "this", "'8-1-003-3'");
            dsTmp = this.calculateAccOverallOAP("PMSP", dsOAPQty, dsOAPUP, dsTmp);
            dsPMSPAccOAPThisYear.Merge(dsTmp);
        }

        #endregion

        private void exportExcel()
        {
            string[] yearThisMonth = dtpYearMonth.Value.ToString("dd/MM/yyyy").Split('/');
            string[] yearPrevMonth = dtpYearMonth.Value.AddMonths(-1).ToString("dd/MM/yyyy").Split('/');

            DateTime thisMonth = DateTime.Parse("01/" + yearThisMonth[1] + "/" + yearThisMonth[2]);
            DateTime prevMonth = DateTime.Parse("01/" + yearPrevMonth[1] + "/" + yearPrevMonth[2]); ;

            FileInfo template = new FileInfo(PATH_Program + @"\Template\Template-Sale-Report.xlsx");
            using (var package = new ExcelPackage(template))
            {
                var workBook = package.Workbook;


                //1.OVERALL RESULT
                var workSheet1 = workBook.Worksheets[1];
                workSheet1.Cells["B1"].Value = prevMonth;
                workSheet1.Cells["D1"].Value = thisMonth;
                for (int i = 0; i < dsOverallFrstPrevMonth.Tables[0].Rows.Count; i++)
                {
                    workSheet1.Cells["B" + (i + 3).ToString()].Value = Convert.ToDouble(dsOverallFrstPrevMonth.Tables[0].Rows[i]["i_amt_frst"].ToString()) / 1000000;
                    workSheet1.Cells["C" + (i + 3).ToString()].Value = Convert.ToDouble(dsOverallActualPrevMonth.Tables[0].Rows[i]["i_amt_sale"].ToString()) / 1000000;
                    workSheet1.Cells["D" + (i + 3).ToString()].Value = Convert.ToDouble(dsOverallFrstThisMonth.Tables[0].Rows[i]["i_amt_frst"].ToString()) / 1000000;
                    workSheet1.Cells["E" + (i + 3).ToString()].Value = Convert.ToDouble(dsOverallActualThisMonth.Tables[0].Rows[i]["i_amt_sale"].ToString()) / 1000000;
                }

                //2.OVERALL PROJECT OEM
                var workSheet2 = workBook.Worksheets[2];
                workSheet2.Cells["B1"].Value = prevMonth;
                workSheet2.Cells["D1"].Value = thisMonth;
                for (int i = 0; i < dsProjOEMActualThisMonth.Tables[0].Rows.Count; i++)
                {
                    workSheet2.Cells["B" + (i + 3).ToString()].Value = Convert.ToDouble(dsProjOEMFrstPrevMonth.Tables[0].Rows[i]["i_amt_frst"].ToString()) / 1000000;
                    workSheet2.Cells["C" + (i + 3).ToString()].Value = Convert.ToDouble(dsProjOEMActualPrevMonth.Tables[0].Rows[i]["i_amt_sale"].ToString()) / 1000000;
                    workSheet2.Cells["D" + (i + 3).ToString()].Value = Convert.ToDouble(dsProjOEMFrstThisMonth.Tables[0].Rows[i]["i_amt_frst"].ToString()) / 1000000;
                    workSheet2.Cells["E" + (i + 3).ToString()].Value = Convert.ToDouble(dsProjOEMActualThisMonth.Tables[0].Rows[i]["i_amt_sale"].ToString()) / 1000000;
                }

                //3.OVERALL PROJECT PMSP
                var workSheet3 = workBook.Worksheets[3];
                workSheet3.Cells["B1"].Value = prevMonth;
                workSheet3.Cells["D1"].Value = thisMonth;
                for (int i = 0; i < dsProjPMSPActualThisMonth.Tables[0].Rows.Count; i++)
                {
                    workSheet3.Cells["B" + (i + 3).ToString()].Value = Convert.ToDouble(dsProjPMSPForecastPrevMonth.Tables[0].Rows[i]["i_amt_frst"].ToString()) / 1000000;
                    workSheet3.Cells["C" + (i + 3).ToString()].Value = Convert.ToDouble(dsProjPMSPActualPrevMonth.Tables[0].Rows[i]["i_amt_sale"].ToString()) / 1000000;
                    workSheet3.Cells["D" + (i + 3).ToString()].Value = Convert.ToDouble(dsProjPMSPForecastThisMonth.Tables[0].Rows[i]["i_amt_frst"].ToString()) / 1000000;
                    workSheet3.Cells["E" + (i + 3).ToString()].Value = Convert.ToDouble(dsProjPMSPActualThisMonth.Tables[0].Rows[i]["i_amt_sale"].ToString()) / 1000000;
                }

                //4.Acc Overall Sale Result
                var workSheet4 = workBook.Worksheets[4];
                for (int i = 0; i < dsOverallYearOAP.Tables[0].Rows.Count; i++)
                {
                    workSheet4.Cells["B" + (i + 2).ToString()].Value = Convert.ToDouble(dsOverallYearOAP.Tables[0].Rows[i]["i_amt"].ToString()) / 1000000;
                    workSheet4.Cells["D" + (i + 2).ToString()].Value = Convert.ToDouble(dsAccumulateOAP.Tables[0].Rows[i]["i_amt"].ToString()) / 1000000;
                    workSheet4.Cells["E" + (i + 2).ToString()].Value = Convert.ToDouble(dsAccActualResult.Tables[0].Rows[i]["i_amt"].ToString()) / 1000000;
                }

                //5.Acc OEM Sale Result
                var workSheet5 = workBook.Worksheets[5];
                for (int i = 0; i < dsProjOEMYearOAP.Tables[0].Rows.Count; i++)
                {
                    workSheet5.Cells["B" + (i + 2).ToString()].Value = Convert.ToDouble(dsProjOEMYearOAP.Tables[0].Rows[i]["i_amt"].ToString()) / 1000000;
                    workSheet5.Cells["D" + (i + 2).ToString()].Value = Convert.ToDouble(dsProjOEMAccumulateOAP.Tables[0].Rows[i]["i_amt"].ToString()) / 1000000;
                    workSheet5.Cells["E" + (i + 2).ToString()].Value = Convert.ToDouble(dsProjOEMAccActualResult.Tables[0].Rows[i]["i_amt"].ToString()) / 1000000;
                }

                //6.Acc PMSP Sale Result
                var workSheet6 = workBook.Worksheets[6];
                for (int i = 0; i < dsProjPMSPYearOAP.Tables[0].Rows.Count; i++)
                {
                    workSheet6.Cells["B" + (i + 2).ToString()].Value = Convert.ToDouble(dsProjPMSPYearOAP.Tables[0].Rows[i]["i_amt"].ToString()) / 1000000;
                    workSheet6.Cells["D" + (i + 2).ToString()].Value = Convert.ToDouble(dsProjPMSPAccumulateOAP.Tables[0].Rows[i]["i_amt"].ToString()) / 1000000;
                    workSheet6.Cells["E" + (i + 2).ToString()].Value = Convert.ToDouble(dsProjPMSPAccActualResult.Tables[0].Rows[i]["i_amt"].ToString()) / 1000000;
                }

                //7.Overall Monthly sales forecast
                var workSheet7 = workBook.Worksheets[7];
                for (int i = 0; i < dsOverallSaleForecat.Tables[0].Rows.Count; i++)
                {
                    workSheet7.Cells["B" + (i + 2).ToString()].Value = Convert.ToDouble(dsOverallSaleForecat.Tables[0].Rows[i]["m1"].ToString()) / 1000000;
                    workSheet7.Cells["C" + (i + 2).ToString()].Value = Convert.ToDouble(dsOverallSaleForecat.Tables[0].Rows[i]["m2"].ToString()) / 1000000;
                    workSheet7.Cells["D" + (i + 2).ToString()].Value = Convert.ToDouble(dsOverallSaleForecat.Tables[0].Rows[i]["m3"].ToString()) / 1000000;
                    workSheet7.Cells["E" + (i + 2).ToString()].Value = Convert.ToDouble(dsOverallSaleForecat.Tables[0].Rows[i]["m4"].ToString()) / 1000000;
                    workSheet7.Cells["F" + (i + 2).ToString()].Value = Convert.ToDouble(dsOverallSaleForecat.Tables[0].Rows[i]["m5"].ToString()) / 1000000;
                    workSheet7.Cells["G" + (i + 2).ToString()].Value = Convert.ToDouble(dsOverallSaleForecat.Tables[0].Rows[i]["m6"].ToString()) / 1000000;
                    workSheet7.Cells["H" + (i + 2).ToString()].Value = Convert.ToDouble(dsOverallSaleForecat.Tables[0].Rows[i]["m7"].ToString()) / 1000000;
                    workSheet7.Cells["I" + (i + 2).ToString()].Value = Convert.ToDouble(dsOverallSaleForecat.Tables[0].Rows[i]["m8"].ToString()) / 1000000;
                    workSheet7.Cells["J" + (i + 2).ToString()].Value = Convert.ToDouble(dsOverallSaleForecat.Tables[0].Rows[i]["m9"].ToString()) / 1000000;
                    workSheet7.Cells["K" + (i + 2).ToString()].Value = Convert.ToDouble(dsOverallSaleForecat.Tables[0].Rows[i]["m10"].ToString()) / 1000000;
                    workSheet7.Cells["L" + (i + 2).ToString()].Value = Convert.ToDouble(dsOverallSaleForecat.Tables[0].Rows[i]["m11"].ToString()) / 1000000;
                    workSheet7.Cells["M" + (i + 2).ToString()].Value = Convert.ToDouble(dsOverallSaleForecat.Tables[0].Rows[i]["m12"].ToString()) / 1000000;
                }

                //8.OEM Monthly sales forecast
                var workSheet8 = workBook.Worksheets[8];
                for (int i = 0; i < dsOEMSaleForecast.Tables[0].Rows.Count; i++)
                {
                    workSheet7.Cells["B" + (i + 3).ToString()].Value = Convert.ToDouble(dsOEMSaleForecast.Tables[0].Rows[i]["i_amt_prev_month"].ToString()) / 1000000;
                    workSheet7.Cells["C" + (i + 3).ToString()].Value = Convert.ToDouble(dsOEMSaleForecast.Tables[0].Rows[i]["i_amt_last_month"].ToString()) / 1000000;
                    workSheet7.Cells["D" + (i + 3).ToString()].Value = Convert.ToDouble(dsOEMSaleForecast.Tables[0].Rows[i]["i_amt_actual_month"].ToString()) / 1000000;
                    workSheet7.Cells["E" + (i + 3).ToString()].Value = Convert.ToDouble(dsOEMSaleForecast.Tables[0].Rows[i]["i_amt_prev_month_n"].ToString()) / 1000000;
                    workSheet7.Cells["F" + (i + 3).ToString()].Value = Convert.ToDouble(dsOEMSaleForecast.Tables[0].Rows[i]["i_amt_last_month_n"].ToString()) / 1000000;
                    workSheet7.Cells["G" + (i + 3).ToString()].Value = Convert.ToDouble(dsOEMSaleForecast.Tables[0].Rows[i]["i_amt_actual_month_n"].ToString()) / 1000000;
                    workSheet7.Cells["H" + (i + 3).ToString()].Value = Convert.ToDouble(dsOEMSaleForecast.Tables[0].Rows[i]["i_amt_prev_month_n1"].ToString()) / 1000000;
                    workSheet7.Cells["I" + (i + 3).ToString()].Value = Convert.ToDouble(dsOEMSaleForecast.Tables[0].Rows[i]["i_amt_last_month_n1"].ToString()) / 1000000;
                    workSheet7.Cells["J" + (i + 3).ToString()].Value = Convert.ToDouble(dsOEMSaleForecast.Tables[0].Rows[i]["i_amt_actual_month_n1"].ToString()) / 1000000;
                    workSheet7.Cells["K" + (i + 3).ToString()].Value = Convert.ToDouble(dsOEMSaleForecast.Tables[0].Rows[i]["i_amt_prev_month_n2"].ToString()) / 1000000;
                    workSheet7.Cells["L" + (i + 3).ToString()].Value = Convert.ToDouble(dsOEMSaleForecast.Tables[0].Rows[i]["i_amt_last_month_n2"].ToString()) / 1000000;
                    workSheet7.Cells["M" + (i + 3).ToString()].Value = Convert.ToDouble(dsOEMSaleForecast.Tables[0].Rows[i]["i_amt_actual_month_n2"].ToString()) / 1000000;
                    workSheet7.Cells["N" + (i + 3).ToString()].Value = Convert.ToDouble(dsOEMSaleForecast.Tables[0].Rows[i]["i_amt_prev_month_n3"].ToString()) / 1000000;
                    workSheet7.Cells["O" + (i + 3).ToString()].Value = Convert.ToDouble(dsOEMSaleForecast.Tables[0].Rows[i]["i_amt_last_month_n3"].ToString()) / 1000000;
                    workSheet7.Cells["P" + (i + 3).ToString()].Value = Convert.ToDouble(dsOEMSaleForecast.Tables[0].Rows[i]["i_amt_actual_month_n3"].ToString()) / 1000000;
                    workSheet7.Cells["Q" + (i + 3).ToString()].Value = Convert.ToDouble(dsOEMSaleForecast.Tables[0].Rows[i]["i_amt_prev_month_n4"].ToString()) / 1000000;
                    workSheet7.Cells["R" + (i + 3).ToString()].Value = Convert.ToDouble(dsOEMSaleForecast.Tables[0].Rows[i]["i_amt_last_month_n4"].ToString()) / 1000000;
                    workSheet7.Cells["S" + (i + 3).ToString()].Value = Convert.ToDouble(dsOEMSaleForecast.Tables[0].Rows[i]["i_amt_actual_month_n4"].ToString()) / 1000000;
                }


                //9.PMSP Monthly sales forecast
                var workSheet9 = workBook.Worksheets[9];
                for (int i = 0; i < dsPMSPSaleForecast.Tables[0].Rows.Count; i++)
                {
                    workSheet9.Cells["B" + (i + 2).ToString()].Value = Convert.ToDouble(dsPMSPSaleForecast.Tables[0].Rows[i]["m1"].ToString()) / 1000000;
                    workSheet9.Cells["C" + (i + 2).ToString()].Value = Convert.ToDouble(dsPMSPSaleForecast.Tables[0].Rows[i]["m2"].ToString()) / 1000000;
                    workSheet9.Cells["D" + (i + 2).ToString()].Value = Convert.ToDouble(dsPMSPSaleForecast.Tables[0].Rows[i]["m3"].ToString()) / 1000000;
                    workSheet9.Cells["E" + (i + 2).ToString()].Value = Convert.ToDouble(dsPMSPSaleForecast.Tables[0].Rows[i]["m4"].ToString()) / 1000000;
                    workSheet9.Cells["F" + (i + 2).ToString()].Value = Convert.ToDouble(dsPMSPSaleForecast.Tables[0].Rows[i]["m5"].ToString()) / 1000000;
                    workSheet9.Cells["G" + (i + 2).ToString()].Value = Convert.ToDouble(dsPMSPSaleForecast.Tables[0].Rows[i]["m6"].ToString()) / 1000000;
                    workSheet9.Cells["H" + (i + 2).ToString()].Value = Convert.ToDouble(dsPMSPSaleForecast.Tables[0].Rows[i]["m7"].ToString()) / 1000000;
                    workSheet9.Cells["I" + (i + 2).ToString()].Value = Convert.ToDouble(dsPMSPSaleForecast.Tables[0].Rows[i]["m8"].ToString()) / 1000000;
                    workSheet9.Cells["J" + (i + 2).ToString()].Value = Convert.ToDouble(dsPMSPSaleForecast.Tables[0].Rows[i]["m9"].ToString()) / 1000000;
                    workSheet9.Cells["K" + (i + 2).ToString()].Value = Convert.ToDouble(dsPMSPSaleForecast.Tables[0].Rows[i]["m10"].ToString()) / 1000000;
                    workSheet9.Cells["L" + (i + 2).ToString()].Value = Convert.ToDouble(dsPMSPSaleForecast.Tables[0].Rows[i]["m11"].ToString()) / 1000000;
                    workSheet9.Cells["M" + (i + 2).ToString()].Value = Convert.ToDouble(dsPMSPSaleForecast.Tables[0].Rows[i]["m12"].ToString()) / 1000000;
                }


                //10.Accumulate Overall Forecast
                var workSheet10 = workBook.Worksheets[10];
                workSheet10.Cells["B2"].Value = Convert.ToDouble(dsAccOverallOAP.Tables[0].Rows[0]["i_amt"].ToString()) / 1000000;
                workSheet10.Cells["D2"].Value = Convert.ToDouble(dsAccOAPPrevYear.Tables[0].Rows[0]["i_amt"].ToString()) / 1000000;
                workSheet10.Cells["E2"].Value = Convert.ToDouble(dsAccOAPThisYear.Tables[0].Rows[0]["i_amt"].ToString()) / 1000000;

                //11.OEM Accumulate Forecast
                var workSheet11 = workBook.Worksheets[11];
                for (int i = 0; i < dsOEMAccOverallOAP.Tables[0].Rows.Count; i++)
                {
                    workSheet11.Cells["B" + (i + 2).ToString()].Value = Convert.ToDouble(dsOEMAccOverallOAP.Tables[0].Rows[i]["i_amt"].ToString()) / 1000000;
                    workSheet11.Cells["D" + (i + 2).ToString()].Value = Convert.ToDouble(dsOEMAccOAPPrevYear.Tables[0].Rows[i]["i_amt"].ToString()) / 1000000;
                    workSheet11.Cells["E" + (i + 2).ToString()].Value = Convert.ToDouble(dsOEMAccOAPThisYear.Tables[0].Rows[i]["i_amt"].ToString()) / 1000000;
                }

                //12.PMSP Accumulate Forecast
                var workSheet12 = workBook.Worksheets[12];
                workSheet12.Cells["B2"].Value = Convert.ToDouble(dsPMSPAccOverallOAP.Tables[0].Rows[0]["i_amt"].ToString()) / 1000000;
                workSheet12.Cells["D2"].Value = Convert.ToDouble(dsPMSPAccOAPPrevYear.Tables[0].Rows[0]["i_amt"].ToString()) / 1000000;
                workSheet12.Cells["E2"].Value = Convert.ToDouble(dsPMSPAccOAPThisYear.Tables[0].Rows[0]["i_amt"].ToString()) / 1000000;

                // SAVE EXCEL
                string monthStr = dtpYearMonth.Value.ToString("yyyyMM");
                SAVE_PATH = pathDesktop + "\\" + fileNameExcelExport + "_" + monthStr + "-" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
                package.SaveAs(new FileInfo(SAVE_PATH));
            }
        }
    }
}
