using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Sale_Report.Help;
using System.Data;

namespace Sale_Report.Model
{
    class AccumulateForecast
    {
        public Sale_Report.Help.Object obj = new Sale_Report.Help.Object();

        internal DataSet selectAccQtyOverallOAP(string year, string type)
        {
            try
            {
                string strCmd = "SELECT ta.i_year_mnth, ta.i_item_cd, ";
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

                if (type == "this" || type == "all")
                {
                    strCmd += "+ ta.i_mnth_plan_qty13 ";
                    strCmd += "+ ta.i_mnth_plan_qty14 ";
                    strCmd += "+ ta.i_mnth_plan_qty15 ";
                    strCmd += "+ ta.i_mnth_plan_qty16 ";
                    strCmd += "+ ta.i_mnth_plan_qty17 ";
                }
                if (type == "all")
                {
                    strCmd += "+ ta.i_mnth_plan_qty18 ";
                    strCmd += "+ ta.i_mnth_plan_qty19 ";
                    strCmd += "+ ta.i_mnth_plan_qty20 ";
                }

                strCmd += ") AS i_qty ";
                strCmd += "FROM t_is_f20m_kanban_plan_head_tr ta ";
                strCmd += "WHERE ta.i_year_mnth LIKE '%" + year + "%'";

                DataSet ds = obj.oracle.libOracle.GetData(strCmd);
                return ds;
            }
            catch
            {
                return null;
            }
        }

        internal DataSet selectAccUPOverallOAP(string year)
        {
            try
            {
                string strCmd = "SELECT   i_item_cd, CASE ";
                strCmd += "WHEN MAX (i_up) IS NOT NULL ";
                strCmd += "THEN MAX (i_up) ";
                strCmd += "ELSE 0 ";
                strCmd += "END AS i_up ";
                strCmd += "FROM t_so_tr ";
                strCmd += "WHERE TO_CHAR (i_del_date, 'mmyyyy') = '04" + year + "' ";
                strCmd += "GROUP BY i_item_cd";

                DataSet ds = obj.oracle.libOracle.GetData(strCmd);
                return ds;
            }
            catch
            {
                return null;
            }
        }

        internal DataSet selectAccOEMQtyOAP(string year, string type, string dlGroupDesc)
        {
            try
            {
                string strCmd = "SELECT ta.i_year_mnth, ta.i_item_cd, ";
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

                if (type == "this" || type == "all")
                {
                    strCmd += "+ ta.i_mnth_plan_qty13 ";
                    strCmd += "+ ta.i_mnth_plan_qty14 ";
                    strCmd += "+ ta.i_mnth_plan_qty15 ";
                    strCmd += "+ ta.i_mnth_plan_qty16 ";
                    strCmd += "+ ta.i_mnth_plan_qty17 ";
                }
                if (type == "all")
                {
                    strCmd += "+ ta.i_mnth_plan_qty18 ";
                    strCmd += "+ ta.i_mnth_plan_qty19 ";
                    strCmd += "+ ta.i_mnth_plan_qty20 ";
                }

                strCmd += ") AS i_qty ";
                strCmd += "FROM t_is_f20m_kanban_plan_head_tr ta ";
                strCmd += "WHERE ta.i_year_mnth LIKE '%" + year + "%' ";
                strCmd += "AND trim(ta.i_dl_cd) IN ( "+ dlGroupDesc + ") ";

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
