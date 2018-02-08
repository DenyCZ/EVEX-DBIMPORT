using OfficeOpenXml;
using System;
using System.Globalization;

namespace EVEX_DBIMPORT
{
    public class ExcelUtils
    {


        public string GetString(ExcelRange origin, int left, int top, ExcelWorksheet ws)
        {

            string jednaBunka = origin.ToString().Contains(":") ? origin.ToString().Split(':')[0] : origin.ToString();

            char pismeno = jednaBunka.Length == 2 || jednaBunka.Length == 3 ? jednaBunka.ToCharArray()[0] : '\0';
            char cislo = jednaBunka.Length == 2 ? jednaBunka.ToCharArray()[1] : '\0';
            int cisloD = jednaBunka.Length == 3 ? int.Parse(jednaBunka.Substring(1, 2)) : 0;


            int cisloN = cislo != '\0' ? int.Parse(cislo.ToString()) + top : cisloD+top;
            int pismenoN = int.Parse(((int)pismeno).ToString()) + left;

            string final = Convert.ToChar(pismenoN) + cisloN.ToString();

            var x = ws.Cells[final].Value ?? "";

            return x.ToString();
        }

        public int? GetInt(ExcelRange origin, int left, int top, ExcelWorksheet ws)
        {
            bool isInt = int.TryParse(GetString(origin, left, top, ws), out int dump);

            return isInt ? dump as int? : null;
        }

        public float? GetFloat(ExcelRange origin, int left, int top, ExcelWorksheet ws)
        {
            bool isFloat = float.TryParse(GetString(origin, left, top, ws), out float dump);

            return isFloat ? dump as float? : null;
        }

        public double? GetDouble(ExcelRange origin, int left, int top, ExcelWorksheet ws)
        {
            bool isDouble = double.TryParse(GetString(origin, left, top, ws), out double dump);

            return isDouble ? dump as double? : null;
        }

        public bool? GetBool(ExcelRange origin, int left, int top, ExcelWorksheet ws)
        {
            bool isBool = bool.TryParse(GetString(origin, left, top, ws), out bool dump);

            return isBool ? dump as bool? : null;
        }

        public decimal? GetDecimal(ExcelRange origin, int left, int top, ExcelWorksheet ws)
        {
            bool isDecimal = decimal.TryParse(GetString(origin, left, top, ws), out decimal dump);

            return isDecimal ? dump as decimal? : null;
        }

        public long? GetLong(ExcelRange origin, int left, int top, ExcelWorksheet ws)
        {
            bool isLong = long.TryParse(GetString(origin, left, top, ws), out long dump);

            return isLong ? dump as long? : null;
        }

        public DateTime? GetDate(ExcelRange origin, int left, int top, ExcelWorksheet ws)
        {
            string[] formats = { "d.M.yyyy", "dd.MM.yyyy", "d.MM.yyyy", "dd.M.yyyy", "dd.MM.yyyy hh:mm:ss", "dd.MM.yyyy h:mm:ss", "dd.M.yyyy h:mm:ss" };

            bool isDate = DateTime.TryParseExact(GetString(origin, left, top, ws), formats , null, DateTimeStyles.None ,out DateTime dump);

            return isDate ? dump as DateTime? : null;
        }

        public int? GetRow(ExcelAddress address)
        {
            string adresa = address.ToString();
            bool isNumber;

            if(adresa.Length == 2)
            {
                isNumber = int.TryParse(adresa.Substring(1, 1), out int dump);
                return isNumber ? dump as int? : null;
            } else if(adresa.Length == 3)
            {
                isNumber = int.TryParse(adresa.Substring(1, 2), out int dump);
                return isNumber ? dump as int? : null;
            }


            return null;
        }

        public int? GetColumn(ExcelAddress address)
        {
            string adresa = address.ToString();

            char sloupec = adresa.ToCharArray()[0];
            int sloupecAscii = int.Parse(((int)sloupec).ToString())-64;

            return sloupecAscii;
        }
    }
}
