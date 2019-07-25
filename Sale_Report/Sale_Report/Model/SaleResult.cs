using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Sale_Report.Help;
using System.Data;

namespace Sale_Report.Model
{
    class SaleResult
    {
        public Sale_Report.Help.Object obj = new Sale_Report.Help.Object();

        internal DataSet selectOverAllActual(string yearMonth, string groupDesc)
        {
            try
            {
                string strCmd = "SELECT  '" + groupDesc + "' as groupdesc, TO_CHAR (ta.i_ship_date, 'yyyymm') AS MONTH, ";
                strCmd += "SUM (ta.i_amt) AS i_amt_sale ";
                strCmd += "FROM t_ship_tr ta ";
                strCmd += "WHERE TO_CHAR (ta.i_ship_date, 'yyyymm') = '" + yearMonth + "' ";
                strCmd += "AND TRIM (ta.i_del_dest_cd) IN (SELECT tb.i_shipto_cd ";
                strCmd += "FROM t_is_sale_shipto_ms tb ";
                strCmd += "WHERE tb.i_group_desc = '" + groupDesc + "') ";
                strCmd += "AND ta.i_inv_no IS NOT NULL ";
                strCmd += "GROUP BY TO_CHAR (ta.i_ship_date, 'yyyymm')";

                DataSet ds = obj.oracle.libOracle.GetData(strCmd);

                return ds;
            }
            catch
            {
                return null;
            }
        }

        internal DataSet selectOverAllPMSPForecast(string yearMonth)
        {
            try
            {
                string strCmd = "SELECT  'PMSP' as groupdesc, ta.i_frst_year_mnth, ";
                strCmd += "SUM ";
                strCmd += "(  ta.i_mnth_frst_so_qty ";
                strCmd += "* (SELECT CASE ";
                strCmd += "WHEN MAX (tb.i_up) IS NOT NULL ";
                strCmd += "THEN MAX (tb.i_up) ";
                strCmd += "ELSE 0 ";
                strCmd += "END AS i_up ";
                strCmd += "FROM t_ship_tr tb ";
                strCmd += "WHERE tb.i_item_cd = ta.i_item_cd ";
                strCmd += "AND TO_CHAR (tb.i_ship_date, 'yyyymm') = '" + yearMonth + "') ";
                strCmd += ") AS i_amt_frst ";
                strCmd += "FROM t_frst_so_head_tr ta ";
                strCmd += "WHERE ta.i_frst_year_mnth = '" + yearMonth + "' AND ta.i_del_dest_cd = '8-1-003-3' ";
                strCmd += "GROUP BY ta.i_frst_year_mnth";

                DataSet ds = obj.oracle.libOracle.GetData(strCmd);

                return ds;
            }
            catch
            {
                return null;
            }
        }

        internal DataSet selectOverAllForecast(string yearMonth, string groupDesc)
        {
            try
            {
                string strCmd = "SELECT '" + groupDesc + "' as groupdesc,  ta.i_frst_year_mnth, ";
                strCmd += "SUM (  ta.i_mnth_frst_so_qty ";
                strCmd += "* (  tb.i_rm_cst ";
                strCmd += "+ tb.i_pur_item_cst1 ";
                strCmd += "+ tb.i_pur_item_cst2 ";
                strCmd += "+ tb.i_pur_item_cst3 ";
                strCmd += "+ tb.i_paint_cst ";
                strCmd += "+ tb.i_laser_proc_cst ";
                strCmd += "+ tb.i_p_proc_cst ";
                strCmd += "+ tb.i_w_proc_cst ";
                strCmd += "+ tb.i_paint_proc_cst ";
                strCmd += "+ tb.i_pack_proc_cst ";
                strCmd += "+ tb.i_other_proc_cst ";
                strCmd += "+ tb.i_trans_cst ";
                strCmd += "+ tb.i_die_charge_off ";
                strCmd += "+ tb.i_sga_profit ";
                strCmd += ") ";
                strCmd += ") AS i_amt_frst ";
                strCmd += "FROM t_frst_so_head_tr ta LEFT JOIN t_sales_detail_up_ms tb ";
                strCmd += "ON TRIM (ta.i_trade_item_cd) = TRIM (tb.i_trade_item_cd) ";
                strCmd += "AND TRIM (ta.i_customer_cd) = TRIM (tb.i_customer_cd) ";
                strCmd += "AND TO_NUMBER (ta.i_frst_year_mnth) >= ";
                strCmd += "TO_NUMBER (TO_CHAR (tb.i_start_date, 'yyyymm')) ";
                strCmd += "AND TO_NUMBER (ta.i_frst_year_mnth) <= ";
                strCmd += "TO_NUMBER (TO_CHAR (tb.i_end_date, 'yyyymm')) ";
                strCmd += "AND tb.i_com_code = 'TTAT' ";
                strCmd += "AND tb.i_fac_cd = 'FAC01' ";
                strCmd += "WHERE ta.i_frst_year_mnth = '" + yearMonth + "' ";
                strCmd += "AND TRIM (ta.i_del_dest_cd) IN (SELECT tc.i_shipto_cd ";
                strCmd += "FROM t_is_sale_shipto_ms tc ";
                strCmd += "WHERE tc.i_group_desc = '" + groupDesc + "') ";
                strCmd += "GROUP BY ta.i_frst_year_mnth";

                DataSet ds = obj.oracle.libOracle.GetData(strCmd);

                return ds;
            }
            catch
            {
                return null;
            }
        }

        internal DataSet selectProjectOEMActual(string yearMonth, string groupDesc, string projectDesc)
        {
            try
            {
                string strCmd = "SELECT   '" + projectDesc + "' AS projectdesc, TO_CHAR (ta.i_ship_date, 'yyyymm') AS MONTH, ";
                strCmd += "SUM (ta.i_amt) AS i_amt_sale ";
                strCmd += "FROM t_ship_tr ta INNER JOIN t_pm_ms tc ";
                strCmd += "ON ta.i_item_cd = tc.i_item_cd ";
                strCmd += "AND TRIM (tc.i_item_type3) IN ('PJ05', 'PJ08') ";
                strCmd += "WHERE TO_CHAR (ta.i_ship_date, 'yyyymm') = '" + yearMonth + "' ";
                strCmd += "AND TRIM (ta.i_del_dest_cd) IN (SELECT tb.i_shipto_cd ";
                strCmd += "FROM t_is_sale_shipto_ms tb ";
                strCmd += "WHERE tb.i_group_desc = '" + groupDesc + "') ";
                strCmd += "AND ta.i_inv_no IS NOT NULL ";
                strCmd += "GROUP BY TO_CHAR (ta.i_ship_date, 'yyyymm')";

                DataSet ds = obj.oracle.libOracle.GetData(strCmd);
                return ds;
            }
            catch
            {
                return null;
            }
        }

        internal DataSet selectProjectOEMOtherActual(string yearMonth)
        {
            try
            {
                string strCmd = "SELECT   'OTHER' AS projectdesc, TO_CHAR (ta.i_ship_date, 'yyyymm') AS MONTH, ";
                strCmd += "SUM (ta.i_amt) AS i_amt_sale ";
                strCmd += "FROM t_ship_tr ta ";
                strCmd += "WHERE TO_CHAR (ta.i_ship_date, 'yyyymm') = '" + yearMonth + "' ";
                strCmd += "AND TRIM (ta.i_del_dest_cd) IN (SELECT tb.i_shipto_cd ";
                strCmd += "FROM t_is_sale_shipto_ms tb ";
                strCmd += "WHERE tb.i_group_desc = 'OTHER') ";
                strCmd += "AND ta.i_inv_no IS NOT NULL ";
                strCmd += "GROUP BY TO_CHAR (ta.i_ship_date, 'yyyymm')";

                DataSet ds = obj.oracle.libOracle.GetData(strCmd);
                return ds;
            }
            catch
            {
                return null;
            }
        }

        internal DataSet selectProjectOEMOEMPartActual(string yearMonth)
        {
            try
            {
                string strCmd = "SELECT   'OEM' AS projectdesc, TO_CHAR (ta.i_ship_date, 'yyyymm') AS MONTH, ";
                strCmd += "SUM (ta.i_amt) AS i_amt_sale ";
                strCmd += "FROM t_ship_tr ta ";
                strCmd += "WHERE TO_CHAR (ta.i_ship_date, 'yyyymm') = '" + yearMonth + "' ";
                strCmd += "AND TRIM (ta.i_item_cd) IN (SELECT TRIM (tb.i_item_cd) ";
                strCmd += "FROM t_is_sale_oem_part tb) ";
                strCmd += "AND ta.i_inv_no IS NOT NULL ";
                strCmd += "GROUP BY TO_CHAR (ta.i_ship_date, 'yyyymm')";

                DataSet ds = obj.oracle.libOracle.GetData(strCmd);
                return ds;
            }
            catch
            {
                return null;
            }
        }

        internal DataSet selectProjectOEMNCDataActual(string yearMonth)
        {
            try
            {
                string strCmd = "";

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
