using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Sale_Report.Help;
using System.Data;

namespace Sale_Report.Model
{
    class SaleForecast
    {
        public Sale_Report.Help.Object obj = new Sale_Report.Help.Object();

        internal DataSet selectForecastQty(string yearMonth, string type, string dlCodeGroup)
        {
            try
            {
                string strCmd = "SELECT ta.i_frst_year_mnth, ta.i_item_cd, ta.i_mnth_frst_so_qty AS qty_m1, ";
                strCmd += "ta.i_mnth_frst_so_qty_1later AS qty_m2, ";
                strCmd += "ta.i_mnth_frst_so_qty_2later AS qty_m3, ";
                strCmd += "tb.i_mnth_frst_so_qty_3later AS qty_m4, ";
                strCmd += "tb.i_mnth_frst_so_qty_4later AS qty_m5, ";
                strCmd += "tb.i_mnth_frst_so_qty_5later AS qty_m6 ";
                strCmd += "FROM t_frst_so_head_tr ta INNER JOIN t_frst_so_head_tr_add tb ";
                strCmd += "ON ta.i_frst_so_no = tb.i_frst_so_no ";
                strCmd += "AND ta.i_frst_year_mnth = tb.i_frst_year_mnth ";
                strCmd += "WHERE ta.i_frst_year_mnth = '" + yearMonth + "' ";
                if (type == "PMSP")
                {
                    strCmd += "AND ta.i_del_dest_cd = '8-1-003-3'";
                }
                else if (type == "OEM")
                {
                    strCmd += "AND ta.i_del_dest_cd IN (" + dlCodeGroup + ") ";
                }

                DataSet ds = obj.oracle.libOracle.GetData(strCmd);
                return ds;
            }
            catch
            {
                return null;
            }
        }

        internal DataSet selectForecastUnitPrice(string yearMonth, string type, string dlCodeGroup)
        {
            try
            {
                string strCmd = "SELECT   i_item_cd, CASE ";
                strCmd += "WHEN MAX (i_up) IS NOT NULL ";
                strCmd += "THEN MAX (i_up) ";
                strCmd += "ELSE 0 ";
                strCmd += "END AS i_up ";
                strCmd += "FROM t_so_tr ";
                strCmd += "WHERE TO_CHAR (i_del_date, 'yyyymm') = '" + yearMonth + "' ";
                if (type == "PMSP")
                {
                    strCmd += "AND t_so_tr.i_del_dest_cd = '8-1-003-3' ";
                }
                else if (type == "OEM")
                {
                    strCmd += "AND t_so_tr.i_del_dest_cd IN (" + dlCodeGroup + ") ";
                }

                strCmd += "GROUP BY i_item_cd";

                DataSet ds = obj.oracle.libOracle.GetData(strCmd);
                return ds;
            }
            catch
            {
                return null;
            }
        }

        internal DataSet selectActualResult(string yearMonth, string type, string dlCodeGroup)
        {
            try
            {
                string strCmd = "SELECT '" + yearMonth + "' AS i_year_mnth, SUM (ta.i_amt) AS i_amt ";
                strCmd += "FROM t_ship_tr ta ";
                strCmd += "WHERE TO_CHAR (ta.i_ship_date, 'yyyymm') = '" + yearMonth + "' ";

                if (type == "PMSP")
                {
                    strCmd += "AND ta.i_del_dest_cd = '8-1-003-3' ";
                }
                else if (type == "OEM")
                {
                    strCmd += "AND ta.i_del_dest_cd IN (" + dlCodeGroup + ") ";
                }

                strCmd += "AND ta.i_inv_no IS NOT NULL";

                DataSet ds = obj.oracle.libOracle.GetData(strCmd);
                return ds;
            }
            catch
            {
                return null;
            }
        }

        internal DataSet selectMonthlyOAP(string year, string type)
        {
            try
            {
                string strCmd = "SELECT ta.i_year_mnth, ta.i_item_cd, ta.i_mnth_plan_qty1, ta.i_mnth_plan_qty2, ";
                strCmd += "ta.i_mnth_plan_qty3, ta.i_mnth_plan_qty4, ta.i_mnth_plan_qty5, ";
                strCmd += "ta.i_mnth_plan_qty6, ta.i_mnth_plan_qty7, ta.i_mnth_plan_qty8, ";
                strCmd += "ta.i_mnth_plan_qty9, ta.i_mnth_plan_qty10, ta.i_mnth_plan_qty11, ";
                strCmd += "ta.i_mnth_plan_qty12, ta.i_mnth_plan_qty13, ta.i_mnth_plan_qty14, ";
                strCmd += "ta.i_mnth_plan_qty15, ta.i_mnth_plan_qty16, ta.i_mnth_plan_qty17, ";
                strCmd += "ta.i_mnth_plan_qty18, ta.i_mnth_plan_qty19, ta.i_mnth_plan_qty20 ";
                strCmd += "FROM t_is_f20m_kanban_plan_head_tr ta ";
                strCmd += "WHERE ta.i_year_mnth LIKE '%" + year + "%'";
                if (type == "PMSP")
                {
                    strCmd += "AND ta.i_dl_cd IN (SELECT tc.i_shipto_cd ";
                    strCmd += "FROM t_is_sale_shipto_ms tc ";
                    strCmd += "WHERE tc.i_group_desc = 'PMSP')";
                }
                else if (type == "OEM")
                {
                    strCmd += "AND ta.i_dl_cd IN (SELECT tc.i_shipto_cd ";
                    strCmd += "FROM t_is_sale_shipto_ms tc ";
                    strCmd += "WHERE tc.i_group_desc = 'OEM' OR tc.i_group_desc = 'OTHER')";
                }

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
