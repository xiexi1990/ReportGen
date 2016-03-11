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
        public void 第1_2工作质量()
        {
            wordapp.Selection.ParagraphFormat.set_Style("标题 2");
            wordapp.Selection.TypeText("1.2 工作质量" + Environment.NewLine);
            wordapp.Selection.ParagraphFormat.set_Style("标题 3");
            wordapp.Selection.TypeText("1.2.1 整体情况" + Environment.NewLine);
            wordapp.Selection.ParagraphFormat.set_Style("正文");
            strsql = @"select count(log_id) from( select distinct log_id from
 qzdata.qz_abnormity_log a, qzdata.qz_abnormity_evalist b where a.stationid = b.stationid and a.pointid = b.pointid and b.science != '辅助' and _DATE_ABID)";
            strsql = strsql.Replace("_DATE_ABID", datestr_abid);
            int 月总事件数 = orahlper.GetInt32(strsql);

            strsql = @"select count(log_id) from( select distinct log_id from
 qzdata.qz_abnormity_log a, qzdata.qz_abnormity_evalist b where a.stationid = b.stationid and a.pointid = b.pointid and b.science != '辅助' and a.class = '1' and _DATE_ABID)";
            strsql = strsql.Replace("_DATE_ABID", datestr_abid);
            int 月典型事件数 = orahlper.GetInt32(strsql);

            DataTable 图件_事件_分类 = orahlper.GetDataTable(string.Format(@"select sum(decode(is_agree, '0', 1, 0)) graphgood,
sum(decode(is_agree, '1', 1, 0)) graphmiddle,
sum(decode(is_agree, '2', 1, 0)) graphbad,
sum(decode(flag_2, '0', 1, 0)) analygood,
sum(decode(flag_2, '1', 1, 0)) analymiddle,
sum(decode(flag_2, '2', 1, 0)) analybad,
sum(decode(flag_5, '0', 1, 0)) cataggood,
sum(decode(flag_5, '1', 1, 0)) catagbad from(
select distinct a.log_id, B.IS_AGREE, B.FLAG_2, B.FLAG_5 from qzdata.qz_abnormity_log a, qzdata.qz_abnormity_check b, qzdata.qz_abnormity_evalist c where a.log_id = b.log_id and A.STATIONID = c.stationid and a.pointid = c.pointid and C.SCIENCE != '辅助' and {0}
)", datestr_abid));

            tmpstr = string.Format("（1）本月产出事件条目共{0}条，其中典型事件共{1}条；图件标注评价为好、中、差的条目数分别为：{2}条、{3}条和{4}条；事件分析评价为好、中、差的条目数分别为：{5}条、{6}条和{7}条；事件归类评价为对、错的条目数分别为{8}条、{9}条。", 月总事件数,
                月典型事件数, 图件_事件_分类.Rows[0][0], 图件_事件_分类.Rows[0][1], 图件_事件_分类.Rows[0][2], 图件_事件_分类.Rows[0][3], 图件_事件_分类.Rows[0][4], 图件_事件_分类.Rows[0][5], 图件_事件_分类.Rows[0][6], 图件_事件_分类.Rows[0][7] );

            wordapp.Selection.TypeText(tmpstr + Environment.NewLine);

            DataTable 各省应分析仪器套数 = orahlper.GetDataTable(@"select UNITNAME, count(instrid) from (
select distinct b.unitname, a.stationid||'xx'||a.pointid as instrid from
qzdata.qz_abnormity_evalist a, qzdata.qz_abnormity_units b where A.UNITCODE = b.unit_code and a.ab_flag = 'Y' and a.science != '辅助' )
 group by unitname order by unitname");
            DataTable 各省未分析仪器套数 = orahlper.GetDataTable(string.Format(@"select unitname, count(instrid) from(
select distinct b.unitname, a.stationid||'xx'||a.pointid as instrid from
qzdata.qz_abnormity_evalist a, qzdata.qz_abnormity_units b where A.UNITCODE = b.unit_code and a.ab_flag = 'Y' and a.science != '辅助'  and A.STATIONID||'xx'||A.POINTID not in (
select distinct stationid||'xx'||pointid from qzdata.qz_abnormity_log a where {0}
)) group by unitname order by unitname", datestr));
            DataTable 各省应分析仪器完整率 = orahlper.GetDataTable(@"select unitname, round(avg(integrality),2) integ from (select unitname, integrality  from qzdata.qz_abnormity_integrality a where a.subject != '辅助'  and a.flag='Y')  group by unitname order by unitname");
            decimal 总应分析仪器完整率 = orahlper.GetDecimal(@"select round(avg(integrality),2) integ from (select integrality  from qzdata.qz_abnormity_integrality a where a.subject != '辅助'  and a.flag='Y')");

            DataTable 各省事件总条数_典型事件条数 = orahlper.GetDataTable(string.Format(@"select unitname, count(log_id) cnt, sum(decode(class, '1', 1, 0)) typical from (select distinct a.log_id, c.unitname, A.CLASS from qzdata.qz_abnormity_log a, qzdata.qz_dict_stations b, qzdata.qz_abnormity_units c, qzdata.qz_abnormity_evalist d where a.stationid = b.stationid and b.unitcode = c.unit_code and {0} and a.stationid = d.stationid and a.pointid = D.POINTID and d.science != '辅助'
) group by unitname order by unitname", datestr_abid));

            DataTable 各省未审核事件数 = orahlper.GetDataTable(string.Format(@"select unitname, count(log_id) cnt from(
select distinct a.log_id, C.UNITNAME from qzdata.qz_abnormity_log a, qzdata.qz_abnormity_evalist b, qzdata.qz_abnormity_units c where {0} and a.stationid = b.stationid and a.pointid = b.pointid and B.SCIENCE != '辅助' and B.UNITCODE = C.UNIT_CODE
) where log_id not in (select distinct log_id from qzdata.qz_abnormity_check)
group by UNITNAME order by unitname", datestr_abid));

            DataTable 表1_2_1 = new DataTable();
            表1_2_1.Columns.Add("unitname");
            表1_2_1.Columns.Add("shouldanaly", Type.GetType("System.Decimal"));
            表1_2_1.Columns.Add("noanaly", Type.GetType("System.Decimal"));
            表1_2_1.Columns.Add("integ", Type.GetType("System.Decimal"));
            表1_2_1.Columns.Add("logcount", Type.GetType("System.Decimal"));
            表1_2_1.Columns.Add("典型事件（条）");
            表1_2_1.Columns.Add("checkrate", Type.GetType("System.Decimal"));
            for (int i = 0; i < __unitnamelist.GetLength(0); i++)
            {
                DataRow dr = 表1_2_1.NewRow();
                dr[0] = __unitnamelist[i];
                dr[1] = dthlper.ExtractRowByLeftFirstCol_SingleValue(各省应分析仪器套数, dr[0].ToString()) ?? 0;
                dr[2] = dthlper.ExtractRowByLeftFirstCol_SingleValue(各省未分析仪器套数, dr[0].ToString()) ?? 0;
                dr[3] = dthlper.ExtractRowByLeftFirstCol_SingleValue(各省应分析仪器完整率, dr[0].ToString()) ?? 0;
                object[] tmpr = dthlper.ExtractRowByLeftFirstCol_WithoutKey(各省事件总条数_典型事件条数, dr[0].ToString());
                dr[4] = tmpr == null ? 0 : tmpr[0];
                dr[5] = tmpr == null ? 0 : tmpr[1];
                int nocheck = dthlper.ExtractRowByLeftFirstCol_Int(各省未审核事件数, dr[0].ToString()) ?? 0;
                if (Convert.ToInt32(dr[4]) == 0)
                    dr[6] = 0;
                else
                {
                    dr[6] = (1 - nocheck / Convert.ToDecimal(dr[4])) * 100;
                }
                表1_2_1.Rows.Add(dr);
            }
            decimal 总事件审核率 = (1 - Convert.ToDecimal(各省未审核事件数.Compute("sum(cnt)", "true")) / 月总事件数) * 100;

            DataView tmpview = new DataView(表1_2_1);
            for (int typ = 0; typ <= 1; typ++)
            {
                if (typ == 0)
                {
                    tmpstr = string.Format("本月全台网的应分析仪器分析完整率为{0}%；除震防中心未开展相应工作外，", 总应分析仪器完整率);
                }
                else
                {
                    tmpstr = string.Format("本月全台网的事件审核率为{0}%；", Math.Round(总事件审核率,2));
                }
                
                if (typ == 0)
                    tmpview.RowFilter = "integ = 100";
                else
                    tmpview.RowFilter = "checkrate = 100";
                int tmpcount = tmpview.Count;
                if (tmpcount > 0)
                {
                    for (int i = 0; i < tmpcount; i++)
                    {
                        tmpstr += tmpview[i]["unitname"] + "、";
                    }
                    tmpstr = tmpstr.Remove(tmpstr.Length - 1) + "等" + tmpcount;
                    if(typ == 0)
                        tmpstr += "家单位的应分析仪器分析完整率为100%；";
                    else
                        tmpstr += "家单位的事件审核率为100%；";
                }
                for (int filt = 90; filt >= 70; filt -= 10)
                {
                    if(typ == 0)
                        tmpview.RowFilter = string.Format("integ < {0} and integ >= {1}", filt + 10, filt);
                    else
                        tmpview.RowFilter = string.Format("checkrate < {0} and checkrate >= {1}", filt + 10, filt);
                    tmpcount = tmpview.Count;
                    if (tmpcount > 0)
                    {
                        for (int i = 0; i < tmpcount; i++)
                        {
                            tmpstr += tmpview[i]["unitname"] + "、";
                        }
                        tmpstr = tmpstr.Remove(tmpstr.Length - 1) + "等" + tmpcount; 
                        if(typ == 0)
                            tmpstr += "家单位的应分析仪器分析完整率达到" + filt + "%以上；";
                        else
                            tmpstr += "家单位的事件审核率达到" + filt + "%以上；";
                    }
                }
                if(typ == 0)
                    tmpview.RowFilter = "integ < 70 and unitname <> '震防中心'";
                else
                    tmpview.RowFilter = "checkrate < 70 and unitname <> '震防中心'";
                if (tmpview.Count > 0)
                {
                    for (int i = 0; i < tmpview.Count; i++)
                    {
                        tmpstr += tmpview[i]["unitname"] + "、";
                    }
                    tmpstr = tmpstr.Remove(tmpstr.Length - 1) + "等" + tmpview.Count;
                    if(typ == 0)
                        tmpstr += "家单位的应分析仪器分析完整率未达到70%。" + Environment.NewLine;
                    else
                        tmpstr += "家单位的事件审核率未达到70%。" + Environment.NewLine;
                }
                else
                {
                    tmpstr = tmpstr.Remove(tmpstr.Length - 1) + "。" + Environment.NewLine;
                }
                wordapp.Selection.TypeText(tmpstr);
            }

            tmpview.RowFilter = null;  ///重要！否则接下来的语句中可能引发莫名错误
            表1_2_1.Columns.Remove("checkrate");  
            
            for (int i = 0; i < 表1_2_1.Rows.Count; i++)
            {
                for (int j = 1; j <= 5; j++)
                {
                    if (j == 3)
                        continue;
                    if (表1_2_1.Rows[i][j] != DBNull.Value && Convert.ToInt32(表1_2_1.Rows[i][j]) == 0)
                        表1_2_1.Rows[i][j] = DBNull.Value;
                 
                }
            }

            for (int i = 0; i < 表1_2_1.Rows.Count; i++)
            {
                for (int j = 0; j < __unitnamelist.GetLength(0); j++)
                {
                    if (表1_2_1.Rows[i]["unitname"].ToString() == __unitnamelist[j].ToString())
                    {
                        表1_2_1.Rows[i]["unitname"] = __abbrunitnamelist[j];
                        break;
                    }
                }
            }

            ta.AddDupFoldTable(2, 表1_2_1, new string[] { "单位名称", "应分析仪器(套)", "未分析仪器(套)", "分析完整率(%)", "事件记录(条)", "典型事件(条)" }, new int[]{50,30,30,40,30,30}, string.Format("表1.2.1.1 {0}年{1}月区域前兆台网观测数据跟踪分析工作情况统计", the_date.Year, the_date.Month));

            wordapp.Selection.ParagraphFormat.set_Style("标题 3");
            wordapp.Selection.TypeText("1.2.2 区域前兆台网数据跟踪分析月报质量" + Environment.NewLine);
            wordapp.Selection.ParagraphFormat.set_Style("正文");
            wordapp.Selection.TypeText(string.Format("{0}年{1}月，各区域前兆台网数据跟踪分析月报质量评价结果见表1.2.2.1。", the_date.Year, the_date.Month) + Environment.NewLine + Environment.NewLine);

            wordapp.Selection.ParagraphFormat.set_Style("标题 3");
            wordapp.Selection.TypeText("1.2.3 区域前兆台网数据跟踪分析事件记录抽查质量" + Environment.NewLine);
            wordapp.Selection.ParagraphFormat.set_Style("正文");
            wordapp.Selection.TypeText(string.Format("{0}年{1}月，各区域前兆台网数据跟踪分析事件记录抽查质量评价结果见表1.2.3.1。", the_date.Year, the_date.Month) + Environment.NewLine + Environment.NewLine);

            wordapp.Selection.ParagraphFormat.set_Style("标题 3");
            wordapp.Selection.TypeText("1.2.4 各省局月报编写存在问题" + Environment.NewLine);
            wordapp.Selection.ParagraphFormat.set_Style("正文");
            wordapp.Selection.TypeText(Environment.NewLine);

            wordapp.Selection.ParagraphFormat.set_Style("标题 3");
            wordapp.Selection.TypeText("1.2.5 各省局事件记录抽查存在问题" + Environment.NewLine);
            wordapp.Selection.ParagraphFormat.set_Style("正文");
            wordapp.Selection.TypeText(Environment.NewLine);

        }
    }
}