using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EVEX_DBIMPORT
{
    class sBrew
    {
        public int id { get; set; }
        public string name { get; set; }
        public string longName { get; set; }
        public string addInfo { get; set; }

        public sBrew(string name, string LongName, string AddInfo)
        {
            this.name = name;
            this.longName = LongName;
            this.addInfo = AddInfo;
        }
    }
}
