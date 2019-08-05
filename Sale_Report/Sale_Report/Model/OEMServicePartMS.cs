using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Sale_Report.Help;
using System.Data;

namespace Sale_Report.Model
{
    class OEMServicePartMS
    {
        public Sale_Report.Help.Object obj = new Sale_Report.Help.Object();

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

        internal DataSet selectOemServicePartList()
        {
            try
            {
                string strCmd = "select * from T_IS_SALE_PMSP_OEMSERVICE_PART order by i_id";
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
                string strCmd = "insert into T_IS_SALE_PMSP_OEMSERVICE_PART(i_item_cd, i_item_desc, i_create_datetime, i_update_datetime) ";
                strCmd += " values('" + itemCD + "', '" + itemDesc + "', SYSDATE, SYSDATE)";

                int result = obj.oracle.libOracle.ExecuteCommand(strCmd);

                return result;
            }
            catch
            {
                return -1;
            }
        }

        internal int deleteItemCD(string itemCD)
        {
            try
            {
                string strCmd = "delete from T_IS_SALE_PMSP_OEMSERVICE_PART where trim(i_item_cd) = '" + itemCD + "'";
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
