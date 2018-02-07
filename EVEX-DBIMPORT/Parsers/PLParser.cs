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

        private int startColumn;
        private int startRow;
        private int lastColumn;
        private int lastRow;


        private int projektRow;
        private int cilRow;
        private int surovinyRow;
        private int technologyRow;
        private int znaceniRow;
        private int prdataRow;
        private int rozboryRow;
        private string rozborRange;
        private int harmonogramRow;
        private int kontaktRow;
        private int confirmRow;
        private int dateRow;
        private int zmenyRow;



        //ProjectTable
        private int id;
        private string description;
        private int contractor;
        private DateTime lunchStart;
        private DateTime lunchEnd;
        private DateTime analysisStart;
        private DateTime analysisEnd;
        private DateTime date;
        

        


        public PLParser(ExcelWorkbook workbook)
        {
            this.wb = workbook;

            parseCells();

            startFilling();
        }

        private void parseCells()
        {

            ExcelWorksheet ws = wb.Worksheets.First();

            for(int col = 1; col <= 15; col++)
            {
                for(int row = 1; row <= 100; row++)
                {
                    if(ws.Cells[row, col].Value != null && ws.Cells[row, col].Value.ToString() == "Projektový list")
                    {
                        this.startColumn = col;
                        this.startRow = row;
                        this.lastColumn = col + 10;
                        
                        break;
                    }
                }
            }

            for (int row = 1; row <= 100; row++)
            {
                if(ws.Cells[row, startColumn].Value != null)
                {
                    switch(ws.Cells[row, startColumn].Value.ToString().Trim()) //TODO: Really stringy?
                    {   
                        case "Projekt":
                            projektRow = row;
                            break;

                        case "Rozvrh a cíl pokusu":
                            cilRow = row;
                            break;

                        case "Várky":
                            surovinyRow = row;
                            break;

                        case "Technologie      (postup lab. pokusu)":
                            technologyRow = row;
                            break;

                        case "Značení vzorků":
                            znaceniRow = row;
                            break;

                        case "Provozní data":
                            prdataRow = row;
                            break;

                        case "Rozbory":
                            rozboryRow = row;
                            break;

                        case "Časový harmonogram":
                            harmonogramRow = row;
                            break;

                        case "Kontaktní osoby":
                            kontaktRow = row;
                            break;

                        case "PL odsouhlasili":
                            confirmRow = row;
                            break;

                        case "Datum":
                            dateRow = row;
                            break;

                        case "Přehled změn":
                            zmenyRow = row;
                            lastRow = row;
                            break;
                    }
                }
            }




            Console.WriteLine("Parsed");
        }

        private void startFilling()
        {
            fillProjectTable();
        }

        private void fillProjectTable()
        {
            ExcelWorksheet ws = wb.Worksheets[1];

            getDescription(ws);
            getLunch(ws);
            getAnalysis(ws);
            //getContractor();
            getDate(ws);

        }

        private void getDescription(ExcelWorksheet ws)
        {
            string firstline = ws.Cells[projektRow, startColumn+1].Value != null ? ws.Cells[projektRow, startColumn + 1].Value.ToString() : "";
            string secondLine = ws.Cells[projektRow + 1, startColumn + 1].Value != null ? ws.Cells[projektRow + 1, startColumn + 1].Value.ToString() : "";

            this.description = (firstline + " " + secondLine).Trim();
        }

        //TODO: Určitě chci u řádku 173,176, 183, 184, 190 DateTime.MinValue?
        private void getLunch(ExcelWorksheet ws)
        {
            string dateStart = ws.Cells[harmonogramRow, startColumn + 3].Value != null ? ws.Cells[harmonogramRow, startColumn + 3].Value.ToString() : "";
            this.lunchStart = dateStart != "" ? DateTime.Parse(dateStart) : DateTime.MinValue;

            string dateEnd = ws.Cells[harmonogramRow, startColumn + 5].Value != null ? ws.Cells[harmonogramRow, startColumn + 5].Value.ToString() : "";
            this.lunchEnd = dateEnd != "" ? DateTime.Parse(dateStart) : DateTime.MinValue;
        }

        private void getAnalysis(ExcelWorksheet ws)
        {
            string dateStart = ws.Cells[harmonogramRow + 1, startColumn + 3].Value != null ? ws.Cells[harmonogramRow + 1, startColumn + 3].Value.ToString() : "";
            string dateEnd = ws.Cells[harmonogramRow + 1, startColumn + 5].Value != null ? ws.Cells[harmonogramRow + 1, startColumn + 5].Value.ToString() : "";
            this.analysisStart = dateStart != "" ? DateTime.Parse(dateStart) : DateTime.MinValue;
            this.analysisEnd = dateEnd != "" ? DateTime.Parse(dateEnd) : DateTime.MinValue;
        }

        private void getDate(ExcelWorksheet ws)
        {
            string date = ws.Cells[dateRow, startColumn + 1].Value != null ? ws.Cells[dateRow, startColumn + 1].Value.ToString() : "";
            this.date = date != "" ? DateTime.Parse(date) : DateTime.MinValue;
        }

        //TODO: getContractor();
        private void getContractor(ExcelWorksheet ws)
        {
            string surname = ws.Cells[confirmRow, startColumn + 3].Value != null ? ws.Cells[confirmRow, startColumn + 3].Value.ToString() : "";

            int idOsoby;

            //this.contractor = idOsoby;
        }
    }
}
