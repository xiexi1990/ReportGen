#define doprogram

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OracleClient;
using System.Data;
using word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.IO;


namespace ReportGen
{
    partial class A
    {
        string[] __scilist = { "形变", "重力", "地磁", "地电", "流体" };
        string[] __ablist = { "观测系统事件", "自然环境事件", "场地环境事件", "人为干扰事件", "地球物理事件", "不明原因事件" };
        string[] __unitnamelist = { "安徽省", "北京市", "重庆市", "震防中心", "福建省", "广东省", "甘肃省", "广西壮族自治区", "河南省", "湖北省", "河北省", "海南省", "黑龙江", "湖南省", "地壳应力研究所", "地震预测研究所", "地质研究所", "地球物理研究所", "吉林省", "江苏省", "江西省", "辽宁省", "内蒙古自治区", "宁夏回族自治区", "青海省", "四川省", "山东省", "上海市", "陕西省", "山西省", "天津市", "新疆维吾尔自治区", "西藏自治区", "云南省", "浙江省" };
        string[] __abbrunitnamelist = { "安徽", "北京", "重庆", "震防中心", "福建", "广东", "甘肃", "广西", "河南", "湖北", "河北", "海南", "黑龙江", "湖南", "地壳所", "预测所", "地质所", "地球所", "吉林", "江苏", "江西", "辽宁", "内蒙古", "宁夏", "青海", "四川", "山东", "上海", "陕西", "山西", "天津", "新疆", "西藏", "云南", "浙江" };
        string[] __unitcodelist = { "AH", "BJ", "CQ", "DPC", "FJ", "GD", "GS", "GX", "HA", "HB", "HE", "HI", "HL", "HN", "ICD", "IES", "IGL", "IGP", "JL", "JS", "JX", "LN", "NM", "NX", "QH", "SC", "SD", "SH", "SN", "SX", "TJ", "XJ", "XZ", "YN", "ZJ" };
        OracleConnection oracon, oracon2;
        word.Application wordapp;
        word.Document worddoc;
        public TableAdder ta = new TableAdder();
        DateStrFormatter dsf = new DateStrFormatter();
        DateTime the_date = new DateTime(2015, 10, 1, 0, 0, 0);
        DateTime the_month_begin, the_month_end;
        string datestr_abid, datestr;
        string bitemfilterstr = "(bitem != '无' and bitem != '?' and bitem != '??' and bitem != '???' and bitem != '????' and bitem != '?????' )";
        string type2filterstr = "(type2_name != '其它分量受影响，本分量正常')";
        string strsql, tmpstr;
        OracleHelper orahlper;
        DataTableHelper dthlper = new DataTableHelper();
        public A()
        {
            oracon = new OracleConnection("server = 127.0.0.1/orcx; user id = qzdata; password = xie51");
            oracon2 = new OracleConnection("server = 10.5.67.11/pdbqz; user id = qzdata; password = qz9401tw");
            wordapp = new word.Application();
            worddoc = new word.Document();
            worddoc = wordapp.Documents.Add();
     //       wordapp.Visible = true;
            ta.wordapp = wordapp;
            ta.worddoc = worddoc;
            datestr = dsf.GetDateStr(the_date);
            datestr_abid = "(" + datestr + "and a.ab_id >=1 and a.ab_id <= 7)";
            oracon2.Open();
            orahlper = new OracleHelper(oracon2);
            orahlper.feedback = true;

            the_month_begin = new DateTime(the_date.Year, the_date.Month, 1, 0, 0, 0);
            the_month_end = the_month_begin.AddMonths(1).AddSeconds(-1);

        }
        ~A()
        {
            oracon.Close();
            worddoc.SaveAs2("d:\\doc5_2");
            wordapp.Visible = true;
            //     worddoc.Close();
            //     wordapp.Quit();
            Console.WriteLine("successfully generated at " + DateTime.Now.ToString());
        }
        
        public DataTable Get_表3_1_2(DateTime t)
        {
            DataTable 各bitem运行总套数 = orahlper.GetDataTable(@"select bitem, count(bitem) from (select distinct a.stationid, a.pointid, c.bitem from qzdata.qz_abnormity_evalist a, qzdata.qz_dict_stationinstruments b , qzdata.qz_abnormity_instrinfo c where a.science != '辅助' and a.stationid = b.stationid and a.pointid = b.pointid and B.INSTRCODE = C.INSTRCODE
) where _BITEMFILTER group by bitem order by bitem".Replace("_BITEMFILTER", bitemfilterstr));
            string sql各bitem各ab_id套数 = @"select bitem, sum(decode(ab_id, '2', 1, 0)) ab2, 
sum(decode(ab_id, '3', 1, 0)) ab3,
sum(decode(ab_id, '4', 1, 0)) ab4,
sum(decode(ab_id, '5', 1, 0)) ab5,
sum(decode(ab_id, '6', 1, 0)) ab6,
sum(decode(ab_id, '7', 1, 0)) ab7
from(
select distinct a.stationid, a.pointid, a.bitem, a.ab_id from qzdata.qz_abnormity_log a, qzdata.qz_abnormity_evalist b  where _DATE and a.stationid = b.stationid and a.pointid = b.pointid and B.SCIENCE != '辅助' and a.ab_id >=2 and a.ab_id <= 7
) where _BITEMFILTER group by bitem order by bitem".Replace("_BITEMFILTER", bitemfilterstr);
            DataTable 各bitem各ab_id套数 = orahlper.GetDataTable(sql各bitem各ab_id套数.Replace("_DATE", dsf.GetDateStr(t)));
            各bitem各ab_id套数.Columns.Add("total", Type.GetType("System.Decimal"));
            for (int i = 0; i < 各bitem各ab_id套数.Rows.Count; i++)
            {
                各bitem各ab_id套数.Rows[i]["total"] = dthlper.ExtractRowByLeftFirstCol_Int(各bitem运行总套数, 各bitem各ab_id套数.Rows[i][0].ToString()) ?? 0;
            }
            DataView 表3_1_2view = 各bitem各ab_id套数.DefaultView;
            表3_1_2view.RowFilter = "total > 0";
            表3_1_2view.Sort = "total desc";
            return 表3_1_2view.ToTable();
        }

        public DataTable Get_表3_1_3(DateTime t)
        {
            DataTable 各仪器类型运行总套数 = orahlper.GetDataTable(@"select name, count(name) from(
select distinct a.stationid, a.pointid, A.INSTRUTYPE||A.INSTRUNAME name from qzdata.qz_pj_evalist a where A.SCIENCE != '辅助'
) group by name  order by name");
            DataTable 各仪器类型各ab_id套数 = orahlper.GetDataTable(@"select name, sum(decode(ab_id, '2', 1, 0)) ab2,
sum(decode(ab_id, '3', 1, 0)) ab3,
sum(decode(ab_id, '4', 1, 0)) ab4,
sum(decode(ab_id, '5', 1, 0)) ab5,
sum(decode(ab_id, '6', 1, 0)) ab6,
sum(decode(ab_id, '7', 1, 0)) ab7
 from  
 
(
select distinct a.stationid, a.pointid, a.ab_id, C.INSTRTYPE||C.INSTRNAME name from qzdata.qz_abnormity_log a, qzdata.qz_abnormity_evalist b, qzdata.qz_abnormity_instrinfo c where _DATE and a.stationid = b.stationid and a.pointid = b.pointid and B.SCIENCE != '辅助' and a.ab_id >=2 and a.ab_id <= 7 and a.instrcode = C.INSTRCODE
) group by name order by name".Replace("_DATE", dsf.GetDateStr(t)));

            各仪器类型各ab_id套数.Columns.Add("total", Type.GetType("System.Decimal"));
            for (int i = 0; i < 各仪器类型各ab_id套数.Rows.Count; i++)
            {
                object tmp = dthlper.ExtractRowByLeftFirstCol_SingleValue(各仪器类型运行总套数, 各仪器类型各ab_id套数.Rows[i][0].ToString());
                if (tmp == null)
                    tmp = 0;
                各仪器类型各ab_id套数.Rows[i]["total"] = Convert.ToInt32(tmp);
            }
            DataView 表3_1_3view = 各仪器类型各ab_id套数.DefaultView;
            表3_1_3view.RowFilter = "total > 0";
            表3_1_3view.Sort = "total desc";
            return 表3_1_3view.ToTable();
        }
     
    
        
    }
   
    class Program
    {
        static void Main(string[] args)
        {
#if doprogram

            //         wordapp.Selection.TypeParagraph();
            //        ta.AddTable(wordapp, worddoc, oracon, "select log_id, instrcode, stationid from qzdata.qz_abnormity_log where rownum <= 5");
            //     Console.WriteLine(ocmd.ExecuteScalar());

            A a = new A();
      //      a.ta.enable = false;
       //     a.试验段();
            a.第1_2工作质量();
            a.第3章();
            return;
#endif
        
        }
    }
}
