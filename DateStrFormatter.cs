using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OracleClient;
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
    class OracleHelper
    {
        public OracleConnection oracon;
        public bool feedback;
        public OracleHelper(OracleConnection oracon, bool feedback = false)
        {
            this.oracon = oracon;
            this.feedback = feedback;
        }
        public int GetInt32(string strsql)
        {
            return Convert.ToInt32(GetSingleValue(strsql));
        }
        public decimal GetDecimal(string strsql)
        {
            return Convert.ToDecimal(GetSingleValue(strsql));
        }
        public object GetSingleValue(string strsql)
        {
     //       TestConnection();
            OracleCommand ocmd = new OracleCommand();
            ocmd.Connection = oracon;
            ocmd.CommandText = strsql;
            if (feedback)
            {
                FeedBack(strsql);
            }
            return ocmd.ExecuteScalar();
        }
        public DataTable GetDataTable(string strsql)
        {
            OracleDataAdapter oda = new OracleDataAdapter(strsql, oracon);
            DataTable dt = new DataTable();
            if (feedback)
            {
                FeedBack(strsql);
            }
            oda.Fill(dt);
            return dt;
        }
        protected void TestConnection()
        {
            if (oracon.State != ConnectionState.Open)
            {
                throw new Exception("OracleHelper: oracle connection is not open");
            }
        }
        protected void FeedBack(string strsql)
        {
            int fdlen;
            if (strsql.Length <= 50)
                fdlen = strsql.Length;
            else
                fdlen = 50;
            System.Console.WriteLine("doing sql: " + strsql.Substring(0, fdlen) + " ...");
        }
    }
    class DataTableHelper
    {
        public object[,] DupFold2DTable_HasColHeader(int dup, object[,] t)
        {
            int rowcount = t.GetLength(0) - 1, colcount = t.GetLength(1);
            int newrowcount = rowcount / dup + 1, newcolcount = colcount * dup;
            if (rowcount % dup != 0)
                newrowcount++;
            object[,] newtable = new object[newrowcount, newcolcount];
            for (int j = 0; j < colcount; j++)
            {
                for (int k = 0; k < dup; k++ )
                    newtable[0, j + k*colcount] = t[0, j];
            }
            for (int i = 0; i < rowcount; i++)
            {
                for (int j = 0; j < colcount; j++)
                {
                    newtable[i % (newrowcount - 1) + 1, j + (i / (newrowcount - 1))*colcount] = t[i + 1, j];
                }
            }
            return newtable;
        }
        public object[,] DataTableTo2DTable(DataTable dt)
        {
            object[,] t = new object[dt.Rows.Count + 1, dt.Columns.Count];
            for (int j = 0; j < t.GetLength(1); j++)
                t[0, j] = dt.Columns[j].ToString();
            for (int i = 0; i < t.GetLength(0) - 1; i++)
            {
                for (int j = 0; j < t.GetLength(1); j++)
                {
                    t[i + 1, j] = dt.Rows[i][j];
                }
            }
            return t;
        }
        public object[] ExtractRowByLeftFirstCol_WithoutKey(DataTable dt, string key)
        {
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (dt.Rows[i][0].ToString() == key)
                {
                    object[] r = new object[dt.Columns.Count - 1];
                    for (int j = 1; j < dt.Columns.Count; j++)
                    {
                        r[j - 1] = dt.Rows[i][j];
                    }
                    return r;
                }
            }
            return null;
        }
        public object ExtractRowByLeftFirstCol_SingleValue(DataTable dt, string key)
        {
            object[] r = ExtractRowByLeftFirstCol_WithoutKey(dt, key);
            if (r == null)
                return null;
            return r[0];
        }
        public int? ExtractRowByLeftFirstCol_Int(DataTable dt, string key)
        {
            object r = ExtractRowByLeftFirstCol_SingleValue(dt, key);
            if (r == null)
                return null;
            return Convert.ToInt32(r);
        }
        public decimal? ExtractRowByLeftFirstCol_Decimal(DataTable dt, string key)
        {
            object r = ExtractRowByLeftFirstCol_SingleValue(dt, key);
            if (r == null)
                return null;
            return Convert.ToDecimal(r);
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
