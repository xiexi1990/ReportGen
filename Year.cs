using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using word = Microsoft.Office.Interop.Word;
using System.Data.OracleClient;
using System.Data;

namespace ReportGen
{
    partial class A
    {
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
                wordapp.Selection.TypeText(string.Format("3.{0}.2 {1}对各区域台网的影响", ab - 1, __abname2[ab - 2]) + Environment.NewLine);
                tmpstr = string.Format("2015年，全国前兆台网存在{0}较多的区域台网有", __abname2[ab - 2]);
                for (int i = 0; i < 10; i++)
                {
                    tmpstr += string.Format("{0}（{1}套）、", year_1view[i][0], year_1view[i][ab - 1]);
                }
                tmpstr = tmpstr.Remove(tmpstr.Length - 1) + string.Format("；全国前兆台网存在{0}比例较多的区域台网有", __abname2[ab - 2]);
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

                ta.AddTable(表3_1_2_year, (string[])null, (int[])null, string.Format("表3.{0}.2   2015年全国地震前兆台网{1}统计（分区域）", ab - 1, __abname2[ab - 2]));

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
    }
}