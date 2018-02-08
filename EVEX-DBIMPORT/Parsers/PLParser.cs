using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EVEX_DBIMPORT
{
    class PLParser
    {
        private ExcelWorkbook wb;

        private cProjectList _Project;
        private cBrew _Brews;
        private ExcelUtils eUtils;

        private ExcelRange origin;
        private ExcelRange footerOrigin;

        public PLParser(ExcelWorkbook workbook)
        {
            this.wb = workbook;
            this._Project = new cProjectList();
            this._Brews = new cBrew();


            this.eUtils = new ExcelUtils();
            

            parseCells();
        }

        private void parseCells()
        {

            ExcelWorksheet ws = wb.Worksheets.First();

            string main = "Projektový list";
            string dates = "Časový harmonogram";

            ExcelAddress address = ws.Cells["A1:K12"].ToList().Find(x => x.Text == main);
            origin = ws.Cells[address.Address];

            ExcelAddress addressDates = ws.Cells["A1:C60"].ToList().Find(x => x.Text == dates);
            footerOrigin = ws.Cells[addressDates.Address];

            _Project.description = eUtils.GetString(origin, 1, 1, ws) + eUtils.GetString(origin, 1, 2, ws);
            _Project.lunchStart = eUtils.GetDate(footerOrigin, 3, 0, ws);
            _Project.lunchEnd = eUtils.GetDate(footerOrigin, 5, 0, ws);
            _Project.analysisStart = eUtils.GetDate(footerOrigin, 3, 1, ws);
            _Project.analysisEnd = eUtils.GetDate(footerOrigin, 5, 1, ws);
            _Project.contractor = eUtils.GetString(footerOrigin, 3, 3, ws);
            _Project.date = eUtils.GetDate(footerOrigin, 1, 4, ws);


            getBrews(ws);
            _Brews.brews.Add(new sBrew("Ahoj", "Ahoj", "Ahoj"));


            Console.WriteLine("Parsed");
        }

        public void getBrews(ExcelWorksheet ws)
        {
            ExcelAddress horniLimit = ws.Cells["A1:D61"].ToList().Find(x => x.Text.Trim() == "Suroviny");
            ExcelAddress dolniLimit = ws.Cells["A1:D61"].ToList().Find(x => x.Text == "Technologie      (postup lab. pokusu)");

            ExcelRange horniLimitR = ws.Cells[horniLimit.Address];
            ExcelRange dolniLimitR = ws.Cells[dolniLimit.Address];

            int? limit = eUtils.GetRow(dolniLimit) - eUtils.GetRow(horniLimit);
        }

    }
}
