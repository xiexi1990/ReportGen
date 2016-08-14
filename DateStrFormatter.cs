using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

using word = Microsoft.Office.Interop.Word;

namespace ReportGen
{
    class DateStrFormatter
    {
        public string GetDateStr(int beginyear, int beginmonth, int endyear, int endmonth, string tablename = "a")
        {
            DateTime begin = new DateTime(beginyear, beginmonth, 1, 0, 0, 0);
            DateTime end = new DateTime(endyear, endmonth, 1, 0, 0, 0);
            end = end.AddMonths(1).AddSeconds(-1);
            return GetDateStr(begin, end, tablename);
        }
        public string GetDateStr(int year, int month, string tablename = "a")
        {
            return GetDateStr(year, month, year, month, tablename);
        }
        public string GetDateStr(DateTime begindate, DateTime enddate, string tablename = "a")
        {
            return string.Format(@"(({0}.end_date >= to_date('{1}','yyyymmdd hh24miss') or {2}.end_date is null )and {3}.start_date <= to_date('{4}','yyyymmdd hh24miss'))", tablename, begindate.ToString("yyyyMMdd HHmmss"), tablename, tablename, enddate.ToString("yyyyMMdd HHmmss"));
        }
        public string GetDateStr(DateTime date, string tablename = "a")
        {
            return GetDateStr(date.Year, date.Month, tablename);
        }

    }
    struct MERGEINTO
    {
        public int row;
        public int col;
        public MERGEINTO(int row, int col)
        {
            this.row = row;
            this.col = col;
        }
    }
}