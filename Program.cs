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
        bool is_year = true;
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
        int the_year_begin_int = 2015,
            the_month_begin_int = 1,
            the_year_end_int = 2015,
            the_month_end_int = 12;
        DateTime the_date = new DateTime(2015, 3, 1, 0, 0, 0);
        DateTime the_month_begin, the_month_end;
        string datestr_abid, datestr;
        string bitemfilterstr = "(bitem != '无' and bitem != '?' and bitem != '??' and bitem != '???' and bitem != '????' and bitem != '?????' )";
        string type2filterstr = "(type2_name != '其它分量受影响，本分量正常')";
        string abidstr = "(a.ab_id >=2 and a.ab_id <=7)";
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

            worddoc.SpellingChecked = false;
            worddoc.ShowSpellingErrors = false;

     //       wordapp.Visible = true;
            ta.wordapp = wordapp;
            ta.worddoc = worddoc;
            if (is_year)
            {
                datestr = dsf.GetDateStr(the_year_begin_int, the_month_begin_int, the_year_end_int, the_month_end_int);
            }
            else
            {
                datestr = dsf.GetDateStr(the_date);
            }
            
       //     datestr_abid = "(" + datestr + "and a.ab_id >=1 and a.ab_id <= 7)";
            datestr_abid = "(" + datestr + "and" + abidstr + ")";
            oracon2.Open();
            orahlper = new OracleHelper(oracon2);
            orahlper.feedback = true;

            the_month_begin = new DateTime(the_date.Year, the_date.Month, 1, 0, 0, 0);
            the_month_end = the_month_begin.AddMonths(1).AddSeconds(-1);

        }
        ~A()
        {
            oracon.Close();
            worddoc.SaveAs2("d:\\year2_6");
            wordapp.Visible = true;
            //     worddoc.Close();
            //     wordapp.Quit();
            Console.WriteLine("successfully generated at " + DateTime.Now.ToString());
        }
        public void Year()
        {
            string[] __abanaly = { "观测系统故障分析", "自然环境干扰分析", "场地环境影响分析", "人为干扰分析", "地球物理事件分析", "不明原因事件分析" };
            string[] __abname2 = { "观测系统故障", "自然环境干扰", "场地环境影响", "人为干扰", "地球物理事件", "不明原因事件" };
            const int ab_end = 7;
            for (int ab = 2; ab <= ab_end; ab++)
            {
                DataTable year_1 = Get_year_1(the_year_begin_int, the_month_begin_int, the_year_end_int, the_month_end_int);

                DataTable year_1比率 = year_1.Copy();
                for (int i = 0; i < year_1.Rows.Count; i++)
                {
                    for (int j = 1; j <= 6; j++)
                    {
                        year_1比率.Rows[i][j] = Convert.ToDecimal(year_1.Rows[i][j]) / Convert.ToDecimal(year_1.Rows[i][7]);
                    }
                }

                DataView year_1view = new DataView(year_1),
                    year_1比率view = new DataView(year_1比率);
                year_1view.Sort =
                    year_1比率view.Sort =
                     year_1.Columns[ab - 1].ColumnName + " desc";
            //    year_1比率view.RowFilter = "total >= 50";

                wordapp.Selection.ParagraphFormat.set_Style("标题 3");
                wordapp.Selection.TypeText(string.Format("3.{0}.2 {1}对各区域台网的影响", ab-1, __abname2[ab - 2]) + Environment.NewLine);
                tmpstr = string.Format("2015年，全国前兆台网存在{0}较多的区域台网有", __abname2[ab - 2]);
                for (int i = 0; i < 10; i++)
                {
                    tmpstr += string.Format("{0}（{1}套）、", year_1view[i][0], year_1view[i][ab - 1]);
                }
                tmpstr = tmpstr.Remove(tmpstr.Length - 1) + string.Format("；全国前兆台网存在{0}比例较多的区域台网有", __abname2[ab-2]);
                for (int i = 0; i < 10; i++)
                {
                    tmpstr += string.Format("{0}（{1}%）、", year_1比率view[i][0], Math.Round(Convert.ToDecimal(year_1比率view[i][ab - 1]) * 100, 1));
                }
               
                tmpstr = tmpstr.Remove(tmpstr.Length - 1) + "）。" + Environment.NewLine;
                wordapp.Selection.ParagraphFormat.set_Style("正文");
                wordapp.Selection.TypeText(tmpstr);

                
                DataTable 表3_1_2_year = new DataTable();
                表3_1_2_year.Columns.Add("单位名称");
                for (int i = 1; i <= 12; i++)
                {
                    表3_1_2_year.Columns.Add(string.Format("{0}月", i));
                }
                表3_1_2_year.Columns.Add("受影响仪器数", typeof(decimal));
                表3_1_2_year.Columns.Add("运行总数", typeof(decimal));
                表3_1_2_year.Columns.Add("影响比例（%）");
                DataTable tmp各省局总数 = orahlper.GetDataTable(@"select unitname 单位名称, count(unitname) 运行总数 from (select distinct b.unitname, a.stationid, a.pointid from qzdata.qz_abnormity_evalist a, qzdata.qz_abnormity_units b where a.unitcode = b.unit_code)  group by unitname order by unitname");
                表3_1_2_year.Merge(tmp各省局总数);

                string sql各省局各月套数 = string.Format(@"select unitname, count(unitname) total from(
select distinct a.stationid, a.pointid, c.unitname from qzdata.qz_abnormity_log a, qzdata.qz_abnormity_evalist b, qzdata.qz_abnormity_units c where b.unitcode = c.unit_code and _DATE and a.stationid = b.stationid and a.pointid = b.pointid and B.SCIENCE != '辅助' and a.ab_id = {0}
) group by unitname order by unitname", ab);
                DataTable 各省局全年套数 = orahlper.GetDataTable(sql各省局各月套数.Replace("_DATE", dsf.GetDateStr(the_year_begin_int, the_month_begin_int, the_year_end_int, the_month_end_int))); ;
                DataTable[] 各省局各月套数 = new DataTable[12];
                for (int i = 1; i <= 12; i++)
                {
                    各省局各月套数[i - 1] = orahlper.GetDataTable(sql各省局各月套数.Replace("_DATE", dsf.GetDateStr(the_year_begin_int, i, the_year_end_int, i)));
                    for (int j = 0; j < 表3_1_2_year.Rows.Count; j++)
                    {    
                    表3_1_2_year.Rows[j][i] = dthlper.ExtractRowByLeftFirstCol_Int(各省局各月套数[i - 1], 表3_1_2_year.Rows[j][0].ToString()) ?? 0;
                    }
                }
                for (int i = 0; i < 表3_1_2_year.Rows.Count; i++)
                {
                    表3_1_2_year.Rows[i]["受影响仪器数"] = dthlper.ExtractRowByLeftFirstCol_Int(各省局全年套数, 表3_1_2_year.Rows[i][0].ToString()) ?? 0;
                    表3_1_2_year.Rows[i]["影响比例（%）"] = Math.Round(Convert.ToDecimal(表3_1_2_year.Rows[i]["受影响仪器数"]) * 100 / Convert.ToDecimal(表3_1_2_year.Rows[i]["运行总数"]), 1);
                }
                                
                ta.AddTable(表3_1_2_year, (string[])null, (int[])null, string.Format("表3.{0}.2   2015年全国地震前兆台网{1}统计（分区域）",ab-1, __abname2[ab - 2]));

            }
        }
        public DataTable Get_year_1(int beginyear, int beginmonth, int endyear, int endmonth)
        {
            DataTable 各省局运行总套数 = orahlper.GetDataTable(@"select unitname, count(unitname) from (select distinct b.unitname, a.stationid, a.pointid from qzdata.qz_abnormity_evalist a, qzdata.qz_abnormity_units b where a.unitcode = b.unit_code)  group by unitname");
            string sql各省局各ab_id套数 = @"select unitname, sum(decode(ab_id, '2', 1, 0)) ab2, 
sum(decode(ab_id, '3', 1, 0)) ab3,
sum(decode(ab_id, '4', 1, 0)) ab4,
sum(decode(ab_id, '5', 1, 0)) ab5,
sum(decode(ab_id, '6', 1, 0)) ab6,
sum(decode(ab_id, '7', 1, 0)) ab7
from(
select distinct a.stationid, a.pointid, c.unitname, a.ab_id from qzdata.qz_abnormity_log a, qzdata.qz_abnormity_evalist b, qzdata.qz_abnormity_units c  where b.unitcode = c.unit_code and _DATE and a.stationid = b.stationid and a.pointid = b.pointid and B.SCIENCE != '辅助' and a.ab_id >=2 and a.ab_id <= 7
) group by unitname order by unitname";
            DataTable 各省局各ab_id套数 = orahlper.GetDataTable(sql各省局各ab_id套数.Replace("_DATE", dsf.GetDateStr(beginyear, beginmonth, endyear, endmonth)));
            各省局各ab_id套数.Columns.Add("total", Type.GetType("System.Decimal"));
            for (int i = 0; i < 各省局各ab_id套数.Rows.Count; i++)
            {
                各省局各ab_id套数.Rows[i]["total"] = dthlper.ExtractRowByLeftFirstCol_Int(各省局运行总套数, 各省局各ab_id套数.Rows[i][0].ToString()) ?? 0;
            }
            DataView 表3_1_2view = new DataView(各省局各ab_id套数);
            表3_1_2view.RowFilter = "total > 0";
            表3_1_2view.Sort = "total desc";
            return 表3_1_2view.ToTable();
        }
        
        public DataTable Get_表3_1_2(int beginyear, int beginmonth, int endyear, int endmonth)
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
            DataTable 各bitem各ab_id套数 = orahlper.GetDataTable(sql各bitem各ab_id套数.Replace("_DATE", dsf.GetDateStr(beginyear, beginmonth, endyear, endmonth)));
            各bitem各ab_id套数.Columns.Add("total", Type.GetType("System.Decimal"));
            for (int i = 0; i < 各bitem各ab_id套数.Rows.Count; i++)
            {
                各bitem各ab_id套数.Rows[i]["total"] = dthlper.ExtractRowByLeftFirstCol_Int(各bitem运行总套数, 各bitem各ab_id套数.Rows[i][0].ToString()) ?? 0;
            }
            DataView 表3_1_2view = new DataView(各bitem各ab_id套数);
            表3_1_2view.RowFilter = "total > 0";
            表3_1_2view.Sort = "total desc";
            return 表3_1_2view.ToTable();
        }

        public DataTable Get_表3_1_3(int beginyear, int beginmonth, int endyear, int endmonth)
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
) group by name order by name".Replace("_DATE", dsf.GetDateStr(beginyear, beginmonth, endyear, endmonth)));

            各仪器类型各ab_id套数.Columns.Add("total", Type.GetType("System.Decimal"));
            for (int i = 0; i < 各仪器类型各ab_id套数.Rows.Count; i++)
            {
                object tmp = dthlper.ExtractRowByLeftFirstCol_SingleValue(各仪器类型运行总套数, 各仪器类型各ab_id套数.Rows[i][0].ToString());
                if (tmp == null)
                    tmp = 0;
                各仪器类型各ab_id套数.Rows[i]["total"] = Convert.ToInt32(tmp);
            }
            DataView 表3_1_3view = new DataView(各仪器类型各ab_id套数);
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
       //     a.第1_2工作质量();
       //    a.第3章();
            a.Year();
            return;
#endif
        
        }
    }
}
