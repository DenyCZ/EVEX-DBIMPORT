using OfficeOpenXml;
using System;

namespace EVEX_DBIMPORT
{
    public class ExcelUtils
    {


        public string GetString(ExcelRange origin, int left, int top, ExcelWorksheet ws)
        {
            string jednaBunka = origin.ToString().Split(':')[0];

            char pismeno = jednaBunka.Length == 2 ? jednaBunka.ToCharArray()[0] : '\0';
            char cislo = jednaBunka.Length == 2 ? jednaBunka.ToCharArray()[1] : '\0';

            int cisloN = int.Parse(cislo.ToString()) + top;
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
    }
}
