using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OracleClient;

namespace ReportGen
{
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
    
   
    
}

