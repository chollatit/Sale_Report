using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Sale_Report.Help;
using System.Data;

namespace Sale_Report.Model
{
    class ShiptoMS
    {
        Sale_Report.Help.Object obj = new Sale_Report.Help.Object();

        internal DataSet selectShiptoMS()
        {
            try
            {
                string strCmd = "select * from T_IS_SALE_SHIPTO_MS order by i_id asc";
                DataSet ds = obj.oracle.libOracle.GetData(strCmd);

                return ds;
            }
            catch
            {
                return null;
            }
        }

        internal int insertShiptoMS(string groupCD, string groupDesc, string shiptoCD, string shiptoDesc)
        {
            try
            {
                string strCmd = "INSERT INTO t_is_sale_shipto_ms ";
                strCmd += "(i_group_cd, i_group_desc, i_shipto_cd, i_shipto_desc, ";
                strCmd += "i_create_datetime, i_update_datetime) ";
                strCmd += "VALUES ('" + groupCD + "', '" + groupDesc + "', '" + shiptoCD + "', '" + shiptoDesc + "', ";
                strCmd += "SYSDATE, SYSDATE) ";

                int result = obj.oracle.libOracle.ExecuteCommand(strCmd);

                return result;
            }
            catch
            {
                return -1;
            }
        }

        internal DataSet selectShiptoCD()
        {
            try
            {
                string strCmd = "select i_dl_cd, i_dl_arg_desc from T_TRADE_MS where I_DL_CD LIKE '8-%' order by i_dl_cd";
                DataSet ds = obj.oracle.libOracle.GetData(strCmd);

                return ds;
            }
            catch
            {
                return null;
            }
        }

        internal int deleteShiptoMS(string shiptoCD)
        {
            try
            {
                string strCmd = "delete from t_is_sale_shipto_ms where trim(i_shipto_cd) = '" + shiptoCD.Trim() + "'";
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
