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
using System.IO;
using OfficeOpenXml;
using System.Diagnostics;

namespace Sale_Report.View.Report
{
    public partial class SummarySaleReports : Form
    {
        SaleResult saleResult = new SaleResult();
        AccumulateResult accResult = new AccumulateResult();

        DataSet dsOverallActualThisMonth = new DataSet();
        DataSet dsOverallFrstThisMonth = new DataSet();
        DataSet dsOverallActualPrevMonth = new DataSet();
        DataSet dsOverallFrstPrevMonth = new DataSet();

        DataSet dsProjOEMActualThisMonth = new DataSet();
        DataSet dsProjOEMActualPrevMonth = new DataSet();
        DataSet dsProjOEMFrstThisMonth = new DataSet();
        DataSet dsProjOEMFrstPrevMonth = new DataSet();

        DataSet dsProjPMSPActualThisMonth = new DataSet();
        DataSet dsProjPMSPActualPrevMonth = new DataSet();
        DataSet dsProjPMSPForecastThisMonth = new DataSet();
        DataSet dsProjPMSPForecastPrevMonth = new DataSet();

        DataSet dsOverallYearOAP = new DataSet();
        DataSet dsAccumulateOAP = new DataSet();
        DataSet dsAccActualResult = new DataSet();

        string PATH_Program = Directory.GetCurrentDirectory();
        string pathDesktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        string fileNameExcelExport = "sale_report";
        string SAVE_PATH = "";
        string processCls = "";
        string fileNameExport = "";

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
                //dgvSearchResult.DataSource = dsOverallActualThisMonth.Tables[0];
            }
            else
            {
                MessageBox.Show("No data.", "Result", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            /*
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
             * 
             * */

            processCls = "export";
            exportExcelProcess();
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
                    //processOverallSalesResult();
                    processAccumulateSalesResult();
                    //processOverallSalesForecast();
                    //processAccumulateSalesForecast();
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

            report1OverallSaleResult(yearThisMonth, yearPrevMonth);
            report2OverallSaleProjectOEM(yearThisMonth, yearPrevMonth);
            report3OverallSaleProjectPMSP(yearThisMonth, yearPrevMonth);
        }

        #region OVERALL SALE RESULT

        private void report1OverallSaleResult(string yearThisMonth, string yearPrevMonth)
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

        private void report2OverallSaleProjectOEM(string yearThisMonth, string yearPrevMonth)
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

        private void report3OverallSaleProjectPMSP(string yearThisMonth, string yearPrevMonth)
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
            report5();
            report6();
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

        private void report5()
        { 
        
        }

        private void report6()
        { 
            
        }

        #endregion

        private void processOverallSalesForecast()
        {

        }

        private void processAccumulateSalesForecast()
        {

        }

        private void exportExcel()
        {
            string[] yearThisMonth = dtpYearMonth.Value.ToString("dd/MM/yyyy").Split('/');
            string[] yearPrevMonth = dtpYearMonth.Value.AddMonths(-1).ToString("dd/MM/yyyy").Split('/');

            DateTime thisMonth = DateTime.Parse("01/" + yearThisMonth[1] + "/" + yearThisMonth[2]);
            DateTime prevMonth = DateTime.Parse("01/" + yearPrevMonth[1] + "/" + yearPrevMonth[2]);;

            FileInfo template = new FileInfo(PATH_Program + @"\Template\Template-Sale-Report.xlsx");
            using (var package = new ExcelPackage(template))
            {
                var workBook = package.Workbook;

                /*
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

                 *
                 * */

                //4.Acc Overall Sale Result
                var workSheet4 = workBook.Worksheets[4];
                for (int i = 0; i < dsOverallYearOAP.Tables[0].Rows.Count; i++)
                {
                    workSheet4.Cells["B" + (i + 2).ToString()].Value = Convert.ToDouble(dsOverallYearOAP.Tables[0].Rows[i]["i_amt"].ToString()) / 1000000;
                    workSheet4.Cells["D" + (i + 2).ToString()].Value = Convert.ToDouble(dsAccumulateOAP.Tables[0].Rows[i]["i_amt"].ToString()) / 1000000;
                    workSheet4.Cells["E" + (i + 2).ToString()].Value = Convert.ToDouble(dsAccActualResult.Tables[0].Rows[i]["i_amt_sale"].ToString()) / 1000000;
                }

                // SAVE EXCEL
                string monthStr = dtpYearMonth.Value.ToString("yyyyMM");
                SAVE_PATH = pathDesktop + "\\" + fileNameExcelExport + "_" + monthStr + "-" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
                package.SaveAs(new FileInfo(SAVE_PATH));
            }
        }
    }
}
