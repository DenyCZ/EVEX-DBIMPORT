using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EVEX_DBIMPORT
{
    class cProjectList
    {
        public int id { get; set; }
        public string description { get; set; }
        //public int? contractor { get; set; }
        public string contractor { get; set; }
        public DateTime? lunchStart { get; set; }
        public DateTime? lunchEnd { get; set; }
        public DateTime? analysisStart { get; set; }
        public DateTime? analysisEnd { get; set; } 
        public DateTime? date { get; set; }

    }
}
