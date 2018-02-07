using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EVEX_DBIMPORT
{
    public class ProjectList
    {
        private ExcelPackage PList;

        public ProjectList(ExcelPackage projectList)
        {
            this.PList = projectList;
        }

        public ExcelPackage getPList()
        {
            return this.PList;
        }
    }
}
