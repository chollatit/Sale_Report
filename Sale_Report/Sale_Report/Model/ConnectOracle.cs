using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ConnectOracle_lib;
using Sale_Report.View.Loading;

namespace Sale_Report.Model
{
    class ConnectOracle
    {
        public library libOracle;

        //Parameter Connection
        public string Host = "PRONES";
        public string User = "TCT";
        public string Password = "TCT";
        public Frm_Loading frmLoad;

        public ConnectOracle()
        {
            //Set ConnectionString
            libOracle = new library(Host, User, Password);
        }
    }
}
