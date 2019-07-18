using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ConnectOracle_lib;

namespace Sale_Report.Model
{
    class ConnectOracle
    {
        public library libOracle;

        //Parameter Connection
        public string Host = "PRONES";
        public string User = "TCT";
        public string Password = "TCT";

        public ConnectOracle()
        {
            //Set ConnectionString
            libOracle = new library(Host, User, Password);
        }
    }
}
