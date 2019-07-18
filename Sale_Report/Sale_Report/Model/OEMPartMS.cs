using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Sale_Report.Help;
using System.Data;

namespace Sale_Report.Model
{
    class OEMPartMS
    {
        Sale_Report.Help.Object obj = new Sale_Report.Help.Object();

        internal DataSet selectItemDetail(string itemCD)
        {
            try
            {
                string strCmd = "select i_item_cd, i_item_desc from t_pm_ms where trim(i_item_cd) = '" + itemCD + "'";
                DataSet ds = obj.oracle.libOracle.GetData(strCmd);

                return ds;
            }
            catch
            {
                return null;
            }
        }

        internal int insertOEMItem(string itemCD, string itemDesc)
        {
            try
            {
                string strCmd = "insert into T_IS_SALE_OEM_PART(i_item_cd, i_item_desc, i_create_datetime, i_update_datetime) ";
                strCmd += " values('" + itemCD + "', '" + itemDesc + "', SYSDATE, SYSDATE)";

                int result = obj.oracle.libOracle.ExecuteCommand(strCmd);

                return result;
            }
            catch
            {
                return -1;
            }
        }
    }
}
