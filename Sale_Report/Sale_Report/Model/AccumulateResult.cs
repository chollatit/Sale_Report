using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Sale_Report.Help;
using System.Data;

namespace Sale_Report.Model
{
    class AccumulateResult
    {
        public Sale_Report.Help.Object obj = new Sale_Report.Help.Object();

        internal DataSet selectOverallYearOAP(string yearMonth, string groupcd)
        {
            try
            {
                string strCmd = "SELECT   '" + groupcd + "' AS groupcd, ta.i_year_mnth, ";
                strCmd += "SUM (  (  ta.i_mnth_plan_qty1 ";
                strCmd += "+ ta.i_mnth_plan_qty2 ";
                strCmd += "+ ta.i_mnth_plan_qty3 ";
                strCmd += "+ ta.i_mnth_plan_qty4 ";
                strCmd += "+ ta.i_mnth_plan_qty5 ";
                strCmd += "+ ta.i_mnth_plan_qty6 ";
                strCmd += "+ ta.i_mnth_plan_qty7 ";
                strCmd += "+ ta.i_mnth_plan_qty8 ";
                strCmd += "+ ta.i_mnth_plan_qty9 ";
                strCmd += "+ ta.i_mnth_plan_qty10 ";
                strCmd += "+ ta.i_mnth_plan_qty11 ";
                strCmd += "+ ta.i_mnth_plan_qty12 ";
                strCmd += "+ ta.i_mnth_plan_qty13 ";
                strCmd += "+ ta.i_mnth_plan_qty14 ";
                strCmd += "+ ta.i_mnth_plan_qty15 ";
                strCmd += "+ ta.i_mnth_plan_qty16 ";
                strCmd += "+ ta.i_mnth_plan_qty17 ";
                strCmd += "+ ta.i_mnth_plan_qty18 ";
                strCmd += "+ ta.i_mnth_plan_qty19 ";
                strCmd += "+ ta.i_mnth_plan_qty20 ";
                strCmd += ") ";
                strCmd += "* (SELECT MAX (  i_rm_cst ";
                strCmd += "+ i_pur_item_cst1 ";
                strCmd += "+ i_pur_item_cst2 ";
                strCmd += "+ i_pur_item_cst3 ";
                strCmd += "+ i_paint_cst ";
                strCmd += "+ i_laser_proc_cst ";
                strCmd += "+ i_p_proc_cst ";
                strCmd += "+ i_w_proc_cst ";
                strCmd += "+ i_paint_proc_cst ";
                strCmd += "+ i_pack_proc_cst ";
                strCmd += "+ i_other_proc_cst ";
                strCmd += "+ i_trans_cst ";
                strCmd += "+ i_die_charge_off ";
                strCmd += "+ i_sga_profit ";
                strCmd += ") AS i_up ";
                strCmd += "FROM t_sales_detail_up_ms tb ";
                strCmd += "WHERE TRIM (ta.i_customer_item_cd) = ";
                strCmd += "TRIM (tb.i_trade_item_cd) ";
                strCmd += "AND tb.i_com_code = 'TTAT' ";
                strCmd += "AND tb.i_fac_cd = 'FAC01') ";
                strCmd += ") AS i_amt ";
                strCmd += "FROM t_is_f20m_kanban_plan_head_tr ta ";
                strCmd += "WHERE ta.i_dl_cd IN (SELECT tc.i_shipto_cd ";
                strCmd += "FROM t_is_sale_shipto_ms tc ";
                strCmd += "WHERE tc.i_group_desc = '" + groupcd + "') ";
                strCmd += "AND ta.i_year_mnth = '052019' ";
                strCmd += "GROUP BY ta.i_year_mnth";

                DataSet ds = obj.oracle.libOracle.GetData(strCmd);
                return ds;
            }
            catch
            {
                return null;
            }
        }

        internal DataSet selectOverallPMSPYearOAP(string yearMonth, string groupcd)
        {
            try
            {
                string strCmd = "SELECT   '" + groupcd + "' AS groupcd, ta.i_year_mnth, ";
                strCmd += "SUM (  (  ta.i_mnth_plan_qty1 ";
                strCmd += "+ ta.i_mnth_plan_qty2 ";
                strCmd += "+ ta.i_mnth_plan_qty3 ";
                strCmd += "+ ta.i_mnth_plan_qty4 ";
                strCmd += "+ ta.i_mnth_plan_qty5 ";
                strCmd += "+ ta.i_mnth_plan_qty6 ";
                strCmd += "+ ta.i_mnth_plan_qty7 ";
                strCmd += "+ ta.i_mnth_plan_qty8 ";
                strCmd += "+ ta.i_mnth_plan_qty9 ";
                strCmd += "+ ta.i_mnth_plan_qty10 ";
                strCmd += "+ ta.i_mnth_plan_qty11 ";
                strCmd += "+ ta.i_mnth_plan_qty12 ";
                strCmd += "+ ta.i_mnth_plan_qty13 ";
                strCmd += "+ ta.i_mnth_plan_qty14 ";
                strCmd += "+ ta.i_mnth_plan_qty15 ";
                strCmd += "+ ta.i_mnth_plan_qty16 ";
                strCmd += "+ ta.i_mnth_plan_qty17 ";
                strCmd += "+ ta.i_mnth_plan_qty18 ";
                strCmd += "+ ta.i_mnth_plan_qty19 ";
                strCmd += "+ ta.i_mnth_plan_qty20 ";
                strCmd += ") ";
                strCmd += "* (SELECT CASE ";
                strCmd += "WHEN MAX (tb.i_up) IS NOT NULL ";
                strCmd += "THEN MAX (tb.i_up) ";
                strCmd += "ELSE 0 ";
                strCmd += "END AS i_up ";
                strCmd += "FROM t_so_tr tb ";
                strCmd += "WHERE TRIM (tb.i_item_cd) = TRIM (ta.i_item_cd) ";
                strCmd += "AND TO_CHAR (tb.i_del_date, 'mmyyyy') = '" + yearMonth + "') ";
                strCmd += ") AS i_amt ";
                strCmd += "FROM t_is_f20m_kanban_plan_head_tr ta ";
                strCmd += "WHERE ta.i_dl_cd IN (SELECT tc.i_shipto_cd ";
                strCmd += "FROM t_is_sale_shipto_ms tc ";
                strCmd += "WHERE tc.i_group_desc = '" + groupcd + "') ";
                strCmd += "AND ta.i_year_mnth = '052019' ";
                strCmd += "GROUP BY ta.i_year_mnth";

                DataSet ds = obj.oracle.libOracle.GetData(strCmd);
                return ds;
            }
            catch
            {
                return null;
            }
        }

        internal DataSet selectAccOverallOAP(string yearMonthBegin, string yearMonthEnd, string groupcd)
        {
            try
            {
                int month = Convert.ToInt32(yearMonthEnd.Substring(0, 2));

                string strCmd = "SELECT   '" + groupcd + "' AS groupcd, ta.i_year_mnth, ";
                strCmd += "SUM (  ( ";

                if (month >= 4 && month <= 12)
                {
                    for (int i = 1; i <= month - 3; i++)
                    {
                        strCmd += "+ ta.i_mnth_plan_qty" + (i).ToString() + " ";
                    }
                }
                else if (month >= 1 && month <= 3)
                {
                    for (int i = 1; i <= month + 9; i++)
                    {
                        strCmd += "+ ta.i_mnth_plan_qty" + (i).ToString() + " ";
                    }
                }

                strCmd += ") ";
                strCmd += "* (SELECT MAX (  i_rm_cst ";
                strCmd += "+ i_pur_item_cst1 ";
                strCmd += "+ i_pur_item_cst2 ";
                strCmd += "+ i_pur_item_cst3 ";
                strCmd += "+ i_paint_cst ";
                strCmd += "+ i_laser_proc_cst ";
                strCmd += "+ i_p_proc_cst ";
                strCmd += "+ i_w_proc_cst ";
                strCmd += "+ i_paint_proc_cst ";
                strCmd += "+ i_pack_proc_cst ";
                strCmd += "+ i_other_proc_cst ";
                strCmd += "+ i_trans_cst ";
                strCmd += "+ i_die_charge_off ";
                strCmd += "+ i_sga_profit ";
                strCmd += ") AS i_up ";
                strCmd += "FROM t_sales_detail_up_ms tb ";
                strCmd += "WHERE TRIM (ta.i_customer_item_cd) = ";
                strCmd += "TRIM (tb.i_trade_item_cd) ";
                strCmd += "AND tb.i_com_code = 'TTAT' ";
                strCmd += "AND tb.i_fac_cd = 'FAC01') ";
                strCmd += ") AS i_amt ";
                strCmd += "FROM t_is_f20m_kanban_plan_head_tr ta ";
                strCmd += "WHERE ta.i_dl_cd IN (SELECT tc.i_shipto_cd ";
                strCmd += "FROM t_is_sale_shipto_ms tc ";
                strCmd += "WHERE tc.i_group_desc = '" + groupcd + "') ";
                strCmd += "AND ta.i_year_mnth = '052019' ";
                strCmd += "GROUP BY ta.i_year_mnth";

                DataSet ds = obj.oracle.libOracle.GetData(strCmd);
                return ds;
            }
            catch
            {
                return null;
            }
        }

        internal DataSet selectAccOverallPMSPOAP(string yearMonthBegin, string yearMonthEnd, string groupcd)
        {
            try
            {
                int month = Convert.ToInt32(yearMonthEnd.Substring(0, 2));

                string strCmd = "SELECT   '" + groupcd + "' AS groupcd, ta.i_year_mnth, ";
                strCmd += "SUM (  (  ";

                if (month >= 4 && month <= 12)
                {
                    for (int i = 1; i <= month - 3; i++)
                    {
                        strCmd += "+ ta.i_mnth_plan_qty" + (i).ToString() + " ";
                    }
                }
                else if (month >= 1 && month <= 3)
                {
                    for (int i = 1; i <= month + 9; i++)
                    {
                        strCmd += "+ ta.i_mnth_plan_qty" + (i).ToString() + " ";
                    }
                }

                strCmd += ") ";
                strCmd += "* (SELECT CASE ";
                strCmd += "WHEN MAX (tb.i_up) IS NOT NULL ";
                strCmd += "THEN MAX (tb.i_up) ";
                strCmd += "ELSE 0 ";
                strCmd += "END AS i_up ";
                strCmd += "FROM t_ship_tr tb ";
                strCmd += "WHERE TRIM (tb.i_item_cd) = TRIM (ta.i_item_cd) ";
                strCmd += "AND TO_CHAR (tb.i_ship_date, 'mmyyyy') = '" + yearMonthEnd + "') ";
                strCmd += ") AS i_amt ";
                strCmd += "FROM t_is_f20m_kanban_plan_head_tr ta ";
                strCmd += "WHERE ta.i_dl_cd IN (SELECT tc.i_shipto_cd ";
                strCmd += "FROM t_is_sale_shipto_ms tc ";
                strCmd += "WHERE tc.i_group_desc = '" + groupcd + "') ";
                strCmd += "AND ta.i_year_mnth = '" + yearMonthEnd + "' ";
                strCmd += "GROUP BY ta.i_year_mnth";

                DataSet ds = obj.oracle.libOracle.GetData(strCmd);
                return ds;
            }
            catch
            {
                return null;
            }
        }

        internal DataSet selectOverAllActual(string yearMonthBegin, string yearMonthEnd, string groupDesc)
        {
            try
            {
                string strCmd = "SELECT  '" + groupDesc + "' as groupdesc, ";
                strCmd += "SUM (ta.i_amt) AS i_amt_sale ";
                strCmd += "FROM t_ship_tr ta ";
                strCmd += "WHERE TO_CHAR (ta.i_ship_date, 'mmyyyy') between '" + yearMonthBegin + "' and '" + yearMonthEnd + "' ";
                strCmd += "AND TRIM (ta.i_del_dest_cd) IN (SELECT tb.i_shipto_cd ";
                strCmd += "FROM t_is_sale_shipto_ms tb ";
                strCmd += "WHERE tb.i_group_desc = '" + groupDesc + "') ";
                strCmd += "AND ta.i_inv_no IS NOT NULL ";

                DataSet ds = obj.oracle.libOracle.GetData(strCmd);

                return ds;
            }
            catch
            {
                return null;
            }
        }

        internal DataSet selectOverallPMSPQtyYearOAP(string yearMonth)
        {
            try
            {
                string strCmd = "SELECT 'PMSP' AS groupcd, ta.i_year_mnth, ta.i_item_cd, ";
                strCmd += "(  ta.i_mnth_plan_qty1 ";
                strCmd += "+ ta.i_mnth_plan_qty2 ";
                strCmd += "+ ta.i_mnth_plan_qty3 ";
                strCmd += "+ ta.i_mnth_plan_qty4 ";
                strCmd += "+ ta.i_mnth_plan_qty5 ";
                strCmd += "+ ta.i_mnth_plan_qty6 ";
                strCmd += "+ ta.i_mnth_plan_qty7 ";
                strCmd += "+ ta.i_mnth_plan_qty8 ";
                strCmd += "+ ta.i_mnth_plan_qty9 ";
                strCmd += "+ ta.i_mnth_plan_qty10 ";
                strCmd += "+ ta.i_mnth_plan_qty11 ";
                strCmd += "+ ta.i_mnth_plan_qty12 ";
                strCmd += "+ ta.i_mnth_plan_qty13 ";
                strCmd += "+ ta.i_mnth_plan_qty14 ";
                strCmd += "+ ta.i_mnth_plan_qty15 ";
                strCmd += "+ ta.i_mnth_plan_qty16 ";
                strCmd += "+ ta.i_mnth_plan_qty17 ";
                strCmd += "+ ta.i_mnth_plan_qty18 ";
                strCmd += "+ ta.i_mnth_plan_qty19 ";
                strCmd += "+ ta.i_mnth_plan_qty20 ";
                strCmd += ") AS i_qty ";
                strCmd += "FROM t_is_f20m_kanban_plan_head_tr ta ";
                strCmd += "WHERE ta.i_dl_cd IN (SELECT tc.i_shipto_cd ";
                strCmd += "FROM t_is_sale_shipto_ms tc ";
                strCmd += "WHERE tc.i_group_desc = 'PMSP') ";
                strCmd += "AND ta.i_year_mnth = '052019'";

                DataSet ds = obj.oracle.libOracle.GetData(strCmd);
                return ds;
            }
            catch
            {
                return null;
            }
        }

        internal DataSet selectOverallPMSPUnitPriceYearOAP(string yearMonth)
        {
            try
            {
                string strCmd = "SELECT   i_item_cd, CASE ";
                strCmd += "WHEN MAX (i_up) IS NOT NULL ";
                strCmd += "THEN MAX (i_up) ";
                strCmd += "ELSE 0 ";
                strCmd += "END AS i_up ";
                strCmd += "FROM t_so_tr ";
                strCmd += "WHERE TO_CHAR (i_del_date, 'mmyyyy') = '" + yearMonth + "' ";
                strCmd += "AND TRIM (i_del_dest_cd) = '8-1-003-3' ";
                strCmd += "GROUP BY i_item_cd";

                DataSet ds = obj.oracle.libOracle.GetData(strCmd);
                return ds;
            }
            catch
            {
                return null;
            }
        }

        internal DataSet selectAccOverallPMSPQtyOAP(string yearMonthBegin, string yearMonthEnd, string groupcd)
        {
            try
            {
                int month = Convert.ToInt32(yearMonthEnd.Substring(0, 2));
                string strCmd = "SELECT 'PMSP' AS groupcd, ta.i_year_mnth, ta.i_item_cd, ";
                strCmd += "( ";
                if (month >= 4 && month <= 12)
                {
                    for (int i = 1; i <= month - 3; i++)
                    {
                        strCmd += "+ ta.i_mnth_plan_qty" + (i).ToString() + " ";
                    }
                }
                else if (month >= 1 && month <= 3)
                {
                    for (int i = 1; i <= month + 9; i++)
                    {
                        strCmd += "+ ta.i_mnth_plan_qty" + (i).ToString() + " ";
                    }
                }
                strCmd += ") AS i_qty ";
                strCmd += "FROM t_is_f20m_kanban_plan_head_tr ta ";
                strCmd += "WHERE ta.i_dl_cd IN (SELECT tc.i_shipto_cd ";
                strCmd += "FROM t_is_sale_shipto_ms tc ";
                strCmd += "WHERE tc.i_group_desc = 'PMSP') ";
                strCmd += "AND ta.i_year_mnth = '052019'";

                DataSet ds = obj.oracle.libOracle.GetData(strCmd);
                return ds;
            }
            catch
            {
                return null;
            }
        }

    }
}
