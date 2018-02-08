using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EVEX_DBIMPORT
{
    public class ProjectListWS
    {
        private ExcelPackage PList;

        public ProjectListWS(ExcelPackage projectList)
        {
            this.PList = projectList;
        }

        public ExcelPackage getPList()
        {
            return this.PList;
        }
    }
}
