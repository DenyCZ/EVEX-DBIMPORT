using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EVEX_DBIMPORT
{
    public class General
    {
        internal DatabaseConnection con;

        public General()
        {
            con = new DatabaseConnection();

            if (con.isConnected())
            {
                Console.WriteLine("Connected");
            }
            else
            {
                Console.WriteLine("Not connected");
            }

        }

         
    }
}
