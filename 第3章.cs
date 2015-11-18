#define 第3章第0段
#define 第3章第1节
#define 第3章第2至6节
#define 第3章第7节
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
        public void 第3章()
        {
            const int ab_end = 6;
#if 第3章第0段
            #region 第3章第0段
            wordapp.Selection.ParagraphFormat.set_Style("标题 1");
            wordapp.Selection.TypeText("3 事件分析" + Environment.NewLine);
            wordapp.Selection.ParagraphFormat.set_Style("正文");
            int 月总仪器数含正常 = orahlper.GetInt32((@"select  count(*) from (select distinct a.stationid , a.pointid  from qzdata.qz_abnormity_log a, qzdata.qz_abnormity_evalist b where (_DATE_ABID) and A.STATIONID = b.stationid and a.pointid = b.pointid 
and b.science != '辅助') ").Replace("_DATE_ABID", datestr_abid));

            int 月总台站数不含正常 = orahlper.GetInt32((@"select count(stationid) from (select distinct a.stationid from qzdata.qz_abnormity_log a, qzdata.qz_abnormity_evalist b where (_DATE_ABID) and A.STATIONID = b.stationid and a.pointid = b.pointid 
and b.science != '辅助' and a.ab_id != 1)").Replace("_DATE_ABID", datestr_abid));
            int 月总仪器数不含正常 = orahlper.GetInt32((@"select count(*) from (select distinct a.stationid, a.pointid from qzdata.qz_abnormity_log a, qzdata.qz_abnormity_evalist b where (_DATE_ABID) and A.STATIONID = b.stationid and a.pointid = b.pointid 
and b.science != '辅助' and a.ab_id != 1)").Replace("_DATE_ABID", datestr_abid));
            int 月总事件数不含正常 = orahlper.GetInt32((@"select count(log_id) from (select distinct a.log_id from qzdata.qz_abnormity_log a, qzdata.qz_abnormity_evalist b where (_DATE_ABID) and A.STATIONID = b.stationid and a.pointid = b.pointid 
and b.science != '辅助' and a.ab_id != 1)").Replace("_DATE_ABID", datestr_abid));

            DataTable 月台站数 = orahlper.GetDataTable((@"select count(stationid) from (select distinct a.stationid, a.ab_id from qzdata.qz_abnormity_log a, qzdata.qz_abnormity_evalist b where (_DATE_ABID) and A.STATIONID = b.stationid and a.pointid = b.pointid 
and b.science != '辅助' and a.ab_id >=2 and a.ab_id <= 7
) group by ab_id order by ab_id").Replace("_DATE_ABID", datestr_abid));
            DataTable 月仪器数 = orahlper.GetDataTable((@"select count(instrid) from (select distinct a.stationid || a.pointid as instrid, a.ab_id from qzdata.qz_abnormity_log a, qzdata.qz_abnormity_evalist b where (_DATE_ABID) and A.STATIONID = b.stationid and a.pointid = b.pointid and b.science != '辅助' and a.ab_id >=2 and a.ab_id <= 7
) group by ab_id order by ab_id").Replace("_DATE_ABID", datestr_abid));
            DataTable 月事件数 = orahlper.GetDataTable((@"select count(log_id) from (select distinct a.log_id, a.ab_id from qzdata.qz_abnormity_log a, qzdata.qz_abnormity_evalist b where (_DATE_ABID) and A.STATIONID = b.stationid and a.pointid = b.pointid 
and b.science != '辅助' and a.ab_id >=2 and a.ab_id <= 7
) group by ab_id order by ab_id").Replace("_DATE_ABID", datestr_abid));
            tmpstr = string.Format(@"{0}年{1}月，全国前兆台网共计对{2}套观测仪器进行了数据跟踪分析，有{3}个台站{4}套仪器记录到各类事件（不含正常）共计{5}条，其中", the_date.Year, the_date.Month, 月总仪器数含正常, 月总台站数不含正常, 月总仪器数不含正常, 月总事件数不含正常);
            for (int i = 0; i <= 5; i++)
            {
                tmpstr += string.Format("{0}个台站{1}套仪器记录到{2}{3}条，", 月台站数.Rows[i][0], 月仪器数.Rows[i][0], __ablist[i], 月事件数.Rows[i][0]);
            }
            tmpstr = tmpstr.Remove(tmpstr.Length - 1) + "。" + Environment.NewLine;
            wordapp.Selection.TypeText(tmpstr);
            #endregion
#endif
#if 第3章第1节
            #region 第3章第1节
            wordapp.Selection.ParagraphFormat.set_Style("标题 2");
            wordapp.Selection.TypeText("3.1 总体情况" + Environment.NewLine);
            wordapp.Selection.ParagraphFormat.set_Style("正文");

            DataTable 学科各类型事件数 = orahlper.GetDataTable(@"select science, sum(decode(ab_id, '2',1, 0)) ab2,
sum(decode(ab_id, '3',1, 0)) ab3, 
sum(decode(ab_id, '4',1, 0)) ab4, 
sum(decode(ab_id, '5',1, 0)) ab5, 
sum(decode(ab_id, '6',1, 0)) ab6, 
sum(decode(ab_id, '7',1, 0)) ab7
from  ( 
select distinct a.log_id, a.ab_id, b.science from qzdata.qz_abnormity_log a, qzdata.qz_abnormity_evalist b where A.STATIONID = b.stationid and a.pointid = b.pointid and (_DATE_ABID) and a.ab_id >=2 and a.ab_id <= 7 and B.SCIENCE != '辅助'
) group by science order by decode(science, '形变', 1, '重力', 2, '地磁', 3, '地电', 4, '流体', 5, 'XB', 1, 'ZL', 2, 'DC', 3, 'DD', 4, 'LT', 5)".Replace("_DATE_ABID", datestr_abid));
            DataTable 学科各类型事件分析1 = new DataTable();
            学科各类型事件分析1.Columns.Add("science");
            学科各类型事件分析1.Columns.Add("sum", Type.GetType("System.Int32"));
            学科各类型事件分析1.Columns.Add("ab2r", Type.GetType("System.Decimal"));
            学科各类型事件分析1.Columns.Add("ab3r", Type.GetType("System.Decimal"));
            学科各类型事件分析1.Columns.Add("ab4r", Type.GetType("System.Decimal"));
            学科各类型事件分析1.Columns.Add("ab5r", Type.GetType("System.Decimal"));
            学科各类型事件分析1.Columns.Add("ab6r", Type.GetType("System.Decimal"));
            学科各类型事件分析1.Columns.Add("ab7r", Type.GetType("System.Decimal"));
            {
                int r = 学科各类型事件数.Rows.Count;
                int c = 学科各类型事件分析1.Columns.Count;
                for (int i = 0; i < r; i++)
                {
                    DataRow dr = 学科各类型事件分析1.NewRow();
                    dr[0] = 学科各类型事件数.Rows[i][0];
                    decimal sum = 0;
                    for (int j = 0; j < 6; j++)
                    {
                        sum += Convert.ToInt32(学科各类型事件数.Rows[i][j + 1]);
                    }
                    dr[1] = sum;
                    for (int j = 0; j < 6; j++)
                    {
                        dr[2 + j] = Convert.ToInt32(学科各类型事件数.Rows[i][j + 1]) / sum;
                    }
                    学科各类型事件分析1.Rows.Add(dr);
                }
            }
            //      ta.AddTable(学科各类型事件分析1);

            tmpstr = string.Format("{0}年{1}月，各学科台网产出的事件分析记录情况见表3.1.1。根据各学科台网记录各类型事件数与所有台台网记录对应类型事件总数的比例分析，", the_date.Year, the_date.Month);
            DataView 学科各类型事件分析1view = 学科各类型事件分析1.DefaultView;
            for (int i = 0; i < 6; i++)
            {
                学科各类型事件分析1view.Sort = string.Format("ab{0}r desc", 2 + i);
                string f;
                if (i == 4)
                    f = "{0}台网和{1}台网记录的{2}（大部分为同震响应）相对较高，分别占{3}%和{4}%；";
                else
                    f = "{0}台网和{1}台网记录的{2}相对较高，分别占{3}%和{4}%；";
                tmpstr += string.Format(f, 学科各类型事件分析1view[0]["science"], 学科各类型事件分析1view[1]["science"], __ablist[i], Math.Round(Convert.ToDecimal(学科各类型事件分析1view[0][string.Format("ab{0}r", 2 + i)]) * 100, 2), Math.Round(Convert.ToDecimal(学科各类型事件分析1view[1][string.Format("ab{0}r", 2 + i)]) * 100, 2));
            }
            tmpstr = tmpstr.Remove(tmpstr.Length - 1) + "。" + Environment.NewLine;
            wordapp.Selection.TypeText(tmpstr);



            //    return;
#if linq1
            var q = from x in 学科各类型事件数.AsEnumerable()
                    select new
                        {
                            ab2 = x.Field<decimal>(0),
                            ab3 = x.Field<decimal>(1),
                            ab4 = x.Field<decimal>(2),
                            ab5 = x.Field<decimal>(3),
                            ab6 = x.Field<decimal>(4),
                            ab7 = x.Field<decimal>(5),
                            sum = x.Field<decimal>(0) + x.Field<decimal>(1) + x.Field<decimal>(2) + x.Field<decimal>(3) + x.Field<decimal>(4) + x.Field<decimal>(5)
                        }
                        into m
                        select new 
                            {
                                m.ab2,
                                ab2r = m.ab2 / m.sum,
                                m.ab3,
                                ab3r = m.ab3 / m.sum,
                                m.ab4,
                                ab4r = m.ab4 / m.sum,
                                m.ab5,
                                ab5r = m.ab5 / m.sum,
                                m.ab6,
                                ab6r = m.ab6 / m.sum,
                                m.ab7,
                                ab7r = m.ab7 / m.sum,
                                m.sum
                            };
            object[,] tmptab1 = new object[q.Count<object>(), 13];
            {
                int i = 0;
            foreach (var r in q)
            {
                tmptab1[i,0] = r.ab2;
                tmptab1[i, 1] = r.ab2r;
                tmptab1[i, 2] = r.ab3;
                tmptab1[i, 3] = r.ab3r;
                tmptab1[i, 4] = r.ab4;
                tmptab1[i, 5] = r.ab4r;
                tmptab1[i, 6] = r.ab5;
                tmptab1[i, 7] = r.ab5r;
                tmptab1[i, 8] = r.ab6;
                tmptab1[i, 9] = r.ab6r;
                tmptab1[i, 10] = r.ab7;
                tmptab1[i, 11] = r.ab7r;
                tmptab1[i, 12] = r.sum;
                i++;
            }
            ta.AddTable(wordapp, worddoc, tmptab1);
            }
#endif ////linq1


            //     var q = from x in 学科各类型事件数.AsEnumerable()  select new { ab3 = x.Field<decimal>("ab3"), ab4 = x.Field<decimal>("ab4"), sum = x.Field<decimal>(0) + x.Field<decimal>(1) + x.Field<decimal>(2) + x.Field<decimal>(3) } into m orderby  m.sum descending , m.ab3 select m;
            //    var qarr = q.ToArray();

            DataTable 学科实际分析仪器数 = orahlper.GetDataTable(@"select science, count(instrid) from ( select distinct a.stationid||a.pointid as instrid, b.science from qzdata.qz_abnormity_log a, qzdata.qz_abnormity_evalist b where a.stationid = b.stationid and a.pointid = b.pointid and (_DATE_ABID) and B.SCIENCE != '辅助') group by science order by decode(science, '形变', 1, '重力', 2, '地磁', 3, '地电', 4, '流体', 5, 'XB', 1, 'ZL', 2, 'DC', 3, 'DD', 4, 'LT', 5)".Replace("_DATE_ABID", datestr_abid));
            DataTable 学科完整性 = orahlper.GetDataTable(@"select subject, round(avg(integrality),2) from (select subject, integrality  from qzdata.qz_abnormity_integrality a where a.subject<>'辅助'  and a.flag='Y' )  group by subject order by decode(subject, '形变', 1, '重力', 2, '地磁', 3, '地电', 4, '流体', 5, 'XB', 1, 'ZL', 2, 'DC', 3, 'DD', 4, 'LT', 5)");
            DataTable 学科典型事件数 = orahlper.GetDataTable(@"select science, count(log_id) from (
 select distinct a.log_id, b.science from qzdata.qz_abnormity_log a, qzdata.qz_abnormity_evalist b where a.stationid = b.stationid and a.pointid = b.pointid and (_DATE_ABID) and B.SCIENCE != '辅助' and a.class = '1') group by science order by decode(science, '形变', 1, '重力', 2, '地磁', 3, '地电', 4, '流体', 5, 'XB', 1, 'ZL', 2, 'DC', 3, 'DD', 4, 'LT', 5)".Replace("_DATE_ABID", datestr_abid));
            object[,] 表3_1_1 = new object[7, 10];

            表3_1_1[0, 0] = "学科名称"; 表3_1_1[0, 1] = "实际分析仪器（套）"; 表3_1_1[0, 2] = "完整性（%）"; 表3_1_1[0, 3] = "典型事件（条）"; 表3_1_1[0, 4] = "事件类别（条）"; 表3_1_1[0, 5] = null; 表3_1_1[0, 6] = null; 表3_1_1[0, 7] = null; 表3_1_1[0, 8] = null; 表3_1_1[0, 9] = new MERGEINTO(0, 4);
            表3_1_1[1, 0] = new MERGEINTO(0, 0); 表3_1_1[1, 1] = new MERGEINTO(0, 1); 表3_1_1[1, 2] = new MERGEINTO(0, 2); 表3_1_1[1, 3] = new MERGEINTO(0, 3); 表3_1_1[1, 4] = "观测系统"; 表3_1_1[1, 5] = "自然环境"; 表3_1_1[1, 6] = "场地环境"; 表3_1_1[1, 7] = "人为干扰"; 表3_1_1[1, 8] = "地球物理事件"; 表3_1_1[1, 9] = "不明原因";
            表3_1_1[2, 0] = "形变";
            表3_1_1[3, 0] = "重力";
            表3_1_1[4, 0] = "地磁";
            表3_1_1[5, 0] = "地电";
            表3_1_1[6, 0] = "流体";

            for (int i = 0; i < 5; i++)
            {
                表3_1_1[i + 2, 1] = dthlper.ExtractRowByLeftFirstCol_SingleValue(学科实际分析仪器数, __scilist[i]) ?? 0;
                表3_1_1[i + 2, 2] = dthlper.ExtractRowByLeftFirstCol_SingleValue(学科完整性, __scilist[i]) ?? 0;
                表3_1_1[i + 2, 3] = dthlper.ExtractRowByLeftFirstCol_SingleValue(学科典型事件数, __scilist[i]) ?? 0;
                object[] tmpr = dthlper.ExtractRowByLeftFirstCol_WithoutKey(学科各类型事件数, __scilist[i]);
                for (int j = 0; j < 6; j++)
                {
                    表3_1_1[i + 2, 4 + j] = tmpr == null ? 0 : tmpr[j];
                }
            }
            ta.AddTable(wordapp, worddoc, 表3_1_1, (int[])null, string.Format("表3.1.1 {0}年{1}月学科事件统计", the_date.Year, the_date.Month));
            wordapp.Selection.TypeParagraph();

            DataTable 表3_1_2 = Get_表3_1_2(the_date);
            {
                string[] newcolname = { "测项名称", "观测系统(套)", "自然环境(套)", "场地环境(套)", "人为干扰(套)", "地球物理事件(套)", "不明原因(套)", "运行总数(套)" };

                wordapp.Selection.TypeText(string.Format("{0}年{1}月，前兆台网观测项目事件记录统计见表3.1.2。", the_date.Year, the_date.Month) + Environment.NewLine);

                ta.AddTable(表3_1_2, newcolname, (int[])null, string.Format("表3.1.2 {0}年{1}月前兆台网观测项目事件记录统计", the_date.Year, the_date.Month));
            }
            wordapp.Selection.TypeParagraph();

            DataTable 表3_1_3 = Get_表3_1_3(the_date);
            {
                string[] newcolname = { "仪器名称", "观测系统(套)", "自然环境(套)", "场地环境(套)", "人为干扰(套)", "地球物理事件(套)", "不明原因(套)", "运行总数(套)" };
                wordapp.Selection.TypeText(string.Format("{0}年{1}月，前兆台网仪器事件记录统计见表3.1.3。", the_date.Year, the_date.Month) + Environment.NewLine);
                ta.AddTable(表3_1_3, newcolname, new int[]{90, 53, 53, 53, 53, 53, 53, 53}, string.Format("表3.1.3 {0}年{1}月前兆台网仪器事件记录统计", the_date.Year, the_date.Month));
            }
            wordapp.Selection.TypeParagraph();
            #endregion
#endif
            //        ta.enable = true;
#if 第3章第2至6节
            #region 第3章第2至6节
            string[] __abanaly = { "观测系统故障分析", "自然环境干扰分析", "场地环境影响分析", "人为干扰分析", "地球物理事件分析", "不明原因事件分析" };
            string[] __abname2 = { "观测系统故障", "自然环境干扰", "场地环境影响", "人为干扰", "地球物理事件", "不明原因事件" };
            for (int ab = 2; ab <= ab_end; ab++)
            {
                wordapp.Selection.ParagraphFormat.set_Style("标题 2");
                wordapp.Selection.TypeText(string.Format("3.{0} ", ab) + __abanaly[ab - 2] + Environment.NewLine);
                wordapp.Selection.ParagraphFormat.set_Style("正文");
                wordapp.Selection.TypeText(string.Format(@"{0}年{1}月，全国前兆台网共有{2}个台站{3}套仪器记录到{4}{5}条。记录到{6}的台站分布图见图3.{7}.1。", the_date.Year, the_date.Month, 月台站数.Rows[ab - 2][0], 月仪器数.Rows[ab - 2][0], __ablist[ab - 2], 月事件数.Rows[ab - 2][0], __ablist[ab - 2], ab) + Environment.NewLine);
                wordapp.Selection.TypeText(Environment.NewLine);
                wordapp.Selection.ParagraphFormat.set_Style("图例表例");
                wordapp.Selection.TypeText(string.Format("图3.{0}.1 {1}年{2}月记录到{3}的台站分布图", ab, the_date.Year, the_date.Month, __ablist[ab - 2]) + Environment.NewLine);

            #region 第3_x_1节
                wordapp.Selection.ParagraphFormat.set_Style("标题 3");
                wordapp.Selection.TypeText(string.Format("3.{0}.1 不同测项记录到{1}的主要因素", ab, __ablist[ab - 2]) + Environment.NewLine);
                wordapp.Selection.ParagraphFormat.set_Style("正文");

                string str3 = string.Format("各测项的影响因素统计见表3.{0}.1.1。其中", ab);
                DataTable 表3_x_1_1 = orahlper.GetDataTable(string.Format(@"select bitem, type2_name, count(instrid) from(
select distinct B.BITEM, C.TYPE2_NAME, B.STATIONID||'xx'||B.POINTID instrid from qzdata.qz_abnormity_itemlog a, qzdata.qz_abnormity_log b, QZDATA.QZ_ABNORMITY_TYPE2 c, QZDATA.QZ_ABNORMITY_EVALIST d where A.LOG_ID = b.log_id and A.TYPE2_ID = C.TYPE2_ID and _DATE and b.stationid = d.stationid and b.pointid = D.POINTID and D.SCIENCE != '辅助' and B.AB_ID = {0} 
) where _BITEMFILTER and _TYPE2FILTER group by type2_name, bitem order by bitem, count(instrid) desc", ab).Replace("_DATE", datestr).Replace("_BITEMFILTER", bitemfilterstr).Replace("_TYPE2FILTER", type2filterstr));
                DataView 表3_x_1_1view = 表3_x_1_1.DefaultView;
                DataTable bitemlist = 表3_x_1_1view.ToTable(true, "bitem");
                for (int bitemi = 0; bitemi < bitemlist.Rows.Count; bitemi++)
                {
                    str3 += bitemlist.Rows[bitemi][0].ToString() + "主要受";
                    表3_x_1_1view.RowFilter = string.Format("bitem = '{0}'", bitemlist.Rows[bitemi][0].ToString());
                    int factorcount = 表3_x_1_1view.Count;
                    if (factorcount > 3)
                        factorcount = 3;
                    for (int i = 0; i < factorcount; i++)
                    {
                        str3 += 表3_x_1_1view[i][1].ToString() + "、";
                    }
                    str3 = str3.Remove(str3.Length - 1) + "等影响；";
                }
                str3 = str3.Remove(str3.Length - 1) + "。" + Environment.NewLine;
                wordapp.Selection.TypeText(str3);

                ta.AddFolioTable(表3_x_1_1, new string[] { "测项名称", "影响因素", "仪器套数" }, (int[])null, string.Format("表3.{0}.1.1 {1}年{2}月记录到{3}测项统计", ab, the_date.Year, the_date.Month, __ablist[ab - 2]));

                wordapp.Selection.TypeParagraph();
                #endregion

            #region 第3_x_2节
                DataTable 表3_1_2比率 = 表3_1_2.Copy();
                for (int i = 0; i < 表3_1_2.Rows.Count; i++)
                {
                    for (int j = 1; j <= 6; j++)
                    {
                        表3_1_2比率.Rows[i][j] = Convert.ToDecimal(表3_1_2.Rows[i][j]) / Convert.ToDecimal(表3_1_2.Rows[i][7]);
                    }
                }
                DataTable 上月表3_1_2 = Get_表3_1_2(the_date.AddMonths(-1));
                DataTable 上月表3_1_2比率 = 上月表3_1_2.Copy();
                for (int i = 0; i < 上月表3_1_2.Rows.Count; i++)
                {
                    for (int j = 1; j <= 6; j++)
                    {
                        上月表3_1_2比率.Rows[i][j] = Convert.ToDecimal(上月表3_1_2.Rows[i][j]) / Convert.ToDecimal(上月表3_1_2.Rows[i][7]);
                    }
                }

                DataView 表3_1_2view = 表3_1_2.DefaultView,
                    上月表3_1_2view = 上月表3_1_2.DefaultView,
                    表3_1_2比率view = 表3_1_2比率.DefaultView,
                    上月表3_1_2比率view = 上月表3_1_2比率.DefaultView;
                表3_1_2view.Sort =
                    上月表3_1_2view.Sort =
                    表3_1_2比率view.Sort =
                    上月表3_1_2比率view.Sort = 表3_1_2.Columns[ab - 1].ColumnName + " desc";
                表3_1_2比率view.RowFilter =
                    上月表3_1_2比率view.RowFilter = "total >= 50";

                wordapp.Selection.ParagraphFormat.set_Style("标题 3");
                wordapp.Selection.TypeText(string.Format("3.{0}.2 {1}多发的测项与仪器", ab, __abname2[ab - 2]) + Environment.NewLine);
                tmpstr = string.Format("根据测项统计，{0}从数量角度多发测项为", __abname2[ab - 2]);
                for (int i = 0; i < 4; i++)
                {
                    tmpstr += string.Format("{0}（{1}套）、", 表3_1_2view[i][0], 表3_1_2view[i][ab - 1]);
                }
                tmpstr = tmpstr.Remove(tmpstr.Length - 1) + "（上月多发测项为";
                for (int i = 0; i < 4; i++)
                {
                    tmpstr += string.Format("{0}（{1}套）、", 上月表3_1_2view[i][0], 上月表3_1_2view[i][ab - 1]);
                }
                tmpstr = tmpstr.Remove(tmpstr.Length - 1) + "）；从比率角度多发测项（考虑运行总数在50套以上的测项）为";
                for (int i = 0; i < 4; i++)
                {
                    tmpstr += string.Format("{0}（{1}%）、", 表3_1_2比率view[i][0], Math.Round(Convert.ToDecimal(表3_1_2比率view[i][ab - 1]) * 100, 1));
                }
                tmpstr = tmpstr.Remove(tmpstr.Length - 1) + "（上月从比率角度多发的测项为";
                for (int i = 0; i < 4; i++)
                {
                    tmpstr += string.Format("{0}（{1}%）、", 上月表3_1_2比率view[i][0], Math.Round(Convert.ToDecimal(上月表3_1_2比率view[i][ab - 1]) * 100, 1));
                }
                tmpstr = tmpstr.Remove(tmpstr.Length - 1) + "）。" + Environment.NewLine;

                decimal 月仪器ab_id总套数 = orahlper.GetDecimal(string.Format(@"select count(instrid) from (
select distinct a.stationid||'xx'||a.pointid as instrid from qzdata.qz_abnormity_log a, qzdata.qz_abnormity_evalist b where a.stationid = b.stationid and a.pointid = b.pointid and B.SCIENCE != '辅助' and _DATE and a.ab_id = {0})", ab).Replace("_DATE", datestr));
                decimal 上月仪器ab_id总套数 = orahlper.GetDecimal(string.Format(@"select count(instrid) from (
select distinct a.stationid||'xx'||a.pointid as instrid from qzdata.qz_abnormity_log a, qzdata.qz_abnormity_evalist b where a.stationid = b.stationid and a.pointid = b.pointid and B.SCIENCE != '辅助' and _DATE and a.ab_id = {0})", ab).Replace("_DATE", dsf.GetDateStr(the_date.AddMonths(-1))));
                decimal 总运行仪器数 = orahlper.GetDecimal(@"select count(instrid) from (
select distinct stationid||'xx'||pointid as instrid from  qzdata.qz_abnormity_evalist  where  SCIENCE != '辅助' )");
                decimal growth = (月仪器ab_id总套数 - 上月仪器ab_id总套数) / 上月仪器ab_id总套数 * 100;
                tmpstr += string.Format("从仪器角度，本月出现过{0}的仪器总体比率为{1}%，环比{2}{3}%，在运行数量较多（30套及以上）的仪器中，出现{4}的比率较高的仪器有", __abname2[ab - 2], Math.Round(月仪器ab_id总套数 / 总运行仪器数 * 100, 1), growth < 0 ? "下降" : "上升", Math.Abs(Math.Round(growth, 1)), __abname2[ab - 2]);

                DataTable 表3_1_3比率 = 表3_1_3.Copy();
                for (int i = 0; i < 表3_1_3.Rows.Count; i++)
                {
                    for (int j = 1; j <= 6; j++)
                    {
                        表3_1_3比率.Rows[i][j] = Convert.ToDecimal(表3_1_3.Rows[i][j]) / Convert.ToDecimal(表3_1_3.Rows[i][7]);
                    }
                }
                DataTable 上月表3_1_3 = Get_表3_1_3(the_date.AddMonths(-1));
                DataTable 上月表3_1_3比率 = 上月表3_1_3.Copy();
                for (int i = 0; i < 上月表3_1_3.Rows.Count; i++)
                {
                    for (int j = 1; j <= 6; j++)
                    {
                        上月表3_1_3比率.Rows[i][j] = Convert.ToDecimal(上月表3_1_3.Rows[i][j]) / Convert.ToDecimal(上月表3_1_3.Rows[i][7]);
                    }
                }

                DataView 表3_1_3view = 表3_1_3.DefaultView,
                    上月表3_1_3view = 上月表3_1_3.DefaultView,
                    表3_1_3比率view = 表3_1_3比率.DefaultView,
                    上月表3_1_3比率view = 上月表3_1_3比率.DefaultView;
                表3_1_3view.Sort =
                    上月表3_1_3view.Sort =
                    表3_1_3比率view.Sort =
                    上月表3_1_3比率view.Sort = 表3_1_3.Columns[ab - 1].ColumnName + " desc";

                //      decimal 仪器比率filter = 0.2M;
                表3_1_3比率view.RowFilter = "total >= 30";
                上月表3_1_3比率view.RowFilter = "total >= 30";

                for (int i = 0; i < 4; i++)
                {
                    tmpstr += 表3_1_3比率view[i]["name"] + string.Format("（{0}%）、", Math.Round(Convert.ToDecimal(表3_1_3比率view[i][ab - 1]) * 100, 1));
                }
                tmpstr = tmpstr.Remove(tmpstr.Length - 1) + "（上月出现" + __abname2[ab - 2] + "的比率较高的仪器有";

                for (int i = 0; i < 4; i++)
                {
                    tmpstr += 上月表3_1_3比率view[i]["name"] + string.Format("（{0}%）、", Math.Round(Convert.ToDecimal(上月表3_1_3比率view[i][ab - 1]) * 100, 1));
                }
                tmpstr = tmpstr.Remove(tmpstr.Length - 1) + "）。" + Environment.NewLine;

                wordapp.Selection.ParagraphFormat.set_Style("正文");
                wordapp.Selection.TypeText(tmpstr);
                #endregion

            #region 第3_x_3节
                DataTable 影响因素_仪器套数 = orahlper.GetDataTable(string.Format(@"select type2_name, count(instrid) from(
select distinct d.TYPE2_NAME, b.stationid||'xx'||B.POINTID as instrid from qzdata.qz_abnormity_itemlog a, qzdata.qz_abnormity_log b, qzdata.qz_abnormity_evalist c, QZDATA.QZ_ABNORMITY_TYPE2 d where a.log_id = b.log_id and B.STATIONID = c.stationid and b.pointid = c.pointid and C.SCIENCE != '辅助' and _DATE and a.type2_id = D.TYPE2_ID and B.AB_ID = {0}
) where _TYPE2FILTER group by type2_name order by count(instrid) desc", ab).Replace("_DATE", datestr).Replace("_TYPE2FILTER", type2filterstr));
                DataTable 影响因素_台站数 = orahlper.GetDataTable(string.Format(@"select type2_name, count(stationid) from(
select distinct d.TYPE2_NAME, b.stationid from qzdata.qz_abnormity_itemlog a, qzdata.qz_abnormity_log b, qzdata.qz_abnormity_evalist c, QZDATA.QZ_ABNORMITY_TYPE2 d where a.log_id = b.log_id and B.STATIONID = c.stationid and b.pointid = c.pointid and C.SCIENCE != '辅助' and _DATE and a.type2_id = D.TYPE2_ID and B.AB_ID = {0}
) where _TYPE2FILTER group by type2_name order by count(stationid) desc", ab).Replace("_DATE", datestr).Replace("_TYPE2FILTER", type2filterstr));

                DataTable 影响因素_bitem_仪器套数 = orahlper.GetDataTable(string.Format(@"select type2_name, bitem, count(instrid) from(
select distinct d.TYPE2_NAME, b.stationid||'xx'||B.POINTID as instrid, b.bitem from qzdata.qz_abnormity_itemlog a, qzdata.qz_abnormity_log b, qzdata.qz_abnormity_evalist c, QZDATA.QZ_ABNORMITY_TYPE2 d where a.log_id = b.log_id and B.STATIONID = c.stationid and b.pointid = c.pointid and C.SCIENCE != '辅助' and _DATE and a.type2_id = D.TYPE2_ID and B.AB_ID = {0}
) where _BITEMFILTER and _TYPE2FILTER group by type2_name, bitem order by type2_name, count(instrid) desc", ab).Replace("_DATE", datestr).Replace("_BITEMFILTER", bitemfilterstr).Replace("_TYPE2FILTER", type2filterstr));

                DataTable 影响因素_仪器名称_仪器套数 = orahlper.GetDataTable(string.Format(@"select type2_name, name, count(instrid) from(
select distinct d.TYPE2_NAME, b.stationid||'xx'||B.POINTID as instrid, E.INSTRTYPE||E.INSTRNAME as name from qzdata.qz_abnormity_itemlog a, qzdata.qz_abnormity_log b, qzdata.qz_abnormity_evalist c, QZDATA.QZ_ABNORMITY_TYPE2 d, qzdata.qz_abnormity_instrinfo e where a.log_id = b.log_id and B.STATIONID = c.stationid and b.pointid = c.pointid and C.SCIENCE != '辅助' and {0} and a.type2_id = D.TYPE2_ID and B.AB_ID = {1} and B.INSTRCODE = E.INSTRCODE
) where _TYPE2FILTER group by type2_name, name order by type2_name, count(instrid) desc".Replace("_TYPE2FILTER", type2filterstr), datestr, ab));

                DataView 影响因素_bitem_仪器套数view = 影响因素_bitem_仪器套数.DefaultView;
                DataView 影响因素_仪器名称_仪器套数view = 影响因素_仪器名称_仪器套数.DefaultView;

                DataTable 表3_x_3_1tmp = 影响因素_仪器套数.Copy();
                表3_x_3_1tmp.Columns[0].ColumnName = "影响因素";
                表3_x_3_1tmp.Columns[1].ColumnName = "仪器套数";
                表3_x_3_1tmp.Columns.Add("影响台站数");
                表3_x_3_1tmp.Columns.Add("多发测项");
                表3_x_3_1tmp.Columns.Add("多发仪器");

                for (int i = 0; i < 表3_x_3_1tmp.Rows.Count; i++)
                {
                    表3_x_3_1tmp.Rows[i]["影响台站数"] = dthlper.ExtractRowByLeftFirstCol_SingleValue(影响因素_台站数, 表3_x_3_1tmp.Rows[i][0].ToString());
                    影响因素_bitem_仪器套数view.RowFilter = 影响因素_仪器名称_仪器套数view.RowFilter = string.Format("type2_name = '{0}'", 表3_x_3_1tmp.Rows[i][0].ToString());
                    string tstr = "";
                    int tmpcount = 影响因素_bitem_仪器套数view.Count;
                    if (tmpcount == 0)
                    {
                        tstr = null;
                    }
                    else
                    {
                        if (tmpcount > 4)
                            tmpcount = 4;
                        for (int tmpi = 0; tmpi < tmpcount; tmpi++)
                        {
                            tstr += 影响因素_bitem_仪器套数view[tmpi][1] + "、";
                        }
                        tstr = tstr.Remove(tstr.Length - 1);
                    }
                    表3_x_3_1tmp.Rows[i]["多发测项"] = tstr;
                    tstr = "";
                    tmpcount = 影响因素_仪器名称_仪器套数view.Count;
                    if (tmpcount == 0)
                    {
                        tstr = null;
                    }
                    else
                    {
                        if (tmpcount > 4)
                            tmpcount = 4;
                        for (int tmpi = 0; tmpi < tmpcount; tmpi++)
                        {
                            tstr += 影响因素_仪器名称_仪器套数view[tmpi][1] + "、";
                        }
                        tstr = tstr.Remove(tstr.Length - 1);
                    }
                    表3_x_3_1tmp.Rows[i]["多发仪器"] = tstr;
                }
                DataView 表3_x_3_1tmpview = 表3_x_3_1tmp.DefaultView;
                表3_x_3_1tmpview.RowFilter = "影响台站数 is not null and 多发测项 is not null and 多发仪器 is not null";
                DataTable 表3_x_3_1 = 表3_x_3_1tmpview.ToTable();

                wordapp.Selection.ParagraphFormat.set_Style("标题 3");
                wordapp.Selection.TypeText(string.Format("3.{0}.3 {1}中主要影响因素", ab, __abname2[ab - 2]) + Environment.NewLine);
                wordapp.Selection.ParagraphFormat.set_Style("正文");
                tmpstr = string.Format(@"{0}中各影响因素分类讨论如表3.{1}.3.1所示。从影响因素看，本月{2}突出影响因素有", __abname2[ab - 2], ab, __abname2[ab - 2]);
                for (int i = 0; i < (表3_x_3_1.Rows.Count > 4 ? 4 : 表3_x_3_1.Rows.Count); i++)
                {
                    tmpstr += 表3_x_3_1.Rows[i][0] + "、";
                }
                tmpstr = tmpstr.Remove(tmpstr.Length - 1) + "等。" + Environment.NewLine;
                wordapp.Selection.TypeText(tmpstr);
                ta.AddTable(表3_x_3_1, (string[])null, new int[]{60,40,40,120,200}, string.Format("表3.{0}.3.1 {1}年{2}月{3}影响因素统计", ab, the_date.Year, the_date.Month, __abname2[ab - 2]));
                wordapp.Selection.TypeParagraph();
                #endregion

            #region 第3_x_4节
                DataTable 表3_x_4_1 = orahlper.GetDataTable(string.Format(@"select style, sum(decode(dura, 1, 1, 0)) short, sum(decode(dura, 2, 1, 0)) middle, sum(decode(dura, 3, 1, 0)) lon, sum(decode(dura, 0, 1, 0)) ongoing, count(dura) from (
select distinct decode(b.style, null, '其它', b.style) as style, a.log_id, case when a.END_DATE is null then 0 when a.end_date - a.START_DATE >= 7 then 3 when a.end_date - a.start_date < 3 then 1 else 2 end as dura from qzdata.qz_abnormity_log a, qzdata.qz_abnormity_itemlog b, qzdata.qz_abnormity_evalist c where a.log_id = b.log_id and a.stationid = c.stationid and a.pointid = c.pointid and c.science != '辅助' and a.ab_id = {0} and {1}
) group by style order by count(dura) desc", ab, datestr));
                wordapp.Selection.ParagraphFormat.set_Style("标题 3");
                wordapp.Selection.TypeText(string.Format("3.{0}.4 {1}的持续时间与表现形态", ab, __ablist[ab - 2]) + Environment.NewLine);
                wordapp.Selection.ParagraphFormat.set_Style("正文");
                {
                    string[] newcolname = { "表现形态", "持续时间短（<3天）", "持续时间中（3-7天）", "持续时间长（>7天）", "尚未结束", "合计" };
                    ta.AddTable(表3_x_4_1, newcolname, (int[])null, string.Format("表3.{0}.4.1 {1}年{2}月{3}的持续时间与表现形态统计", ab, the_date.Year, the_date.Month, __ablist[ab - 2]));
                }
                wordapp.Selection.TypeText(string.Format("从表3.{0}.4.1中可看出，", ab));
                if (ab == 2)
                {
                    tmpstr = "本月观测系统故障事件的持续时间以小于7天的中短期故障为主。从表现形态上看，观测系统故障事件对观测曲线造成的影响主要是缺数和错误数据（阶变、突跳、固体潮畸变等也算作错误数据的一种），上升、下降、趋势转折等趋势性变化比较少见。";
                }
                else if (ab == 3)
                {
                    tmpstr = "台站记录到自然环境干扰事件的持续时间多为小于7天的中短期干扰，通常与该自然环境事件（风、雨）等的实际发生时间相一致，也有少量呈现滞后效应。自然环境干扰对观测曲线的影响表现形态主要有上升、畸变、噪声大、下降等。其中上升形态的数量较多，这与降水导致水位上升有关。风力作用则通常导致观测曲线噪声大、畸变等。";
                }
                else if (ab == 4)
                {
                    tmpstr = "场地环境影响中大于7天的长时间持续事件的比例相对其它事件类型要大。场地环境影响对观测曲线的影响表现形态主要有两类：第一类是突跳、阶变、噪声大等表现形态，造成此种表现形态的场地环境影响因素多是振动干扰、基建工程等因素；第二类是上升或下降等表现形态，造成此种表现形态的场地环境影响因素多是抽水或排水、蓄水等因素。";
                }
                else if (ab == 5)
                {
                    tmpstr = "人为干扰事件的持续时间通常较短，持续时间短的人为干扰事件数量占总数的很高比例。由于人为干扰的主要内容是标定、调零等，对观测曲线造成的影响主要有缺数、阶变等。";
                }
                else if (ab == 6)
                {
                    tmpstr = "地球物理事件对观测曲线的影响的主要表现形态分为两类：第一类是突跳和阶变等表现形态，此类变化主要由地震事件引起；第二类是噪声大、固体潮畸变等表现形态，此类变化主要是由地磁、地电学科仪器记录到地磁暴、地电暴的体现。从持续时间来看，占前兆台网记录到的地球物理事件的大部分的地震事件持续时间较短；而地磁暴、地电暴事件的持续时间通常与其实际发生、持续时间相一致，具有部分持续时间中长的记录。";
                }
                else
                {
                    tmpstr = "";
                }
                tmpstr += "" + Environment.NewLine;
                wordapp.Selection.TypeText(tmpstr);
                #endregion
            }
            #endregion
#endif
#if 第3章第7节
            #region 第3章第7节

            wordapp.Selection.ParagraphFormat.set_Style("标题 2");
            wordapp.Selection.TypeText("3.7 不明原因事件分析" + Environment.NewLine);

            int 月疑似前兆异常台站数 = orahlper.GetInt32(string.Format(@"select count(stationid) from(
select distinct b.stationid from qzdata.qz_abnormity_itemlog a, QZDATA.QZ_ABNORMITY_LOG b, QZDATA.QZ_ABNORMITY_TYPE2 c, qzdata.qz_abnormity_evalist d where a.log_id = b.log_id and b.stationid = d.stationid and b.pointid = d.pointid and A.TYPE2_ID = C.TYPE2_ID and D.SCIENCE != '辅助' and b.ab_id = 7 and {0} and C.TYPE2_NAME = '疑似前兆异常')", datestr));
            int 月疑似前兆异常仪器数 = orahlper.GetInt32(string.Format(@"select count(*) from(
select distinct b.stationid, b.pointid from qzdata.qz_abnormity_itemlog a, QZDATA.QZ_ABNORMITY_LOG b, QZDATA.QZ_ABNORMITY_TYPE2 c, qzdata.qz_abnormity_evalist d where a.log_id = b.log_id and b.stationid = d.stationid and b.pointid = d.pointid and A.TYPE2_ID = C.TYPE2_ID and D.SCIENCE != '辅助' and b.ab_id = 7 and {0} and C.TYPE2_NAME = '疑似前兆异常')", datestr));
            int 月疑似前兆异常事件数 = orahlper.GetInt32(string.Format(@"select count(*) from(
select distinct b.log_id from qzdata.qz_abnormity_itemlog a, QZDATA.QZ_ABNORMITY_LOG b, QZDATA.QZ_ABNORMITY_TYPE2 c, qzdata.qz_abnormity_evalist d where a.log_id = b.log_id and b.stationid = d.stationid and b.pointid = d.pointid and A.TYPE2_ID = C.TYPE2_ID and D.SCIENCE != '辅助' and b.ab_id = 7 and {0} and C.TYPE2_NAME = '疑似前兆异常')", datestr));
            decimal 上月不明原因事件数 = orahlper.GetDecimal(string.Format(@"select count(*) from(
select distinct a.LOG_ID from qzdata.qz_abnormity_log a, qzdata.qz_abnormity_evalist d where a.stationid = d.stationid and a.pointid = d.pointid  and D.SCIENCE != '辅助' and a.ab_id = 7 and {0} )", dsf.GetDateStr(the_date.AddMonths(-1))));
            decimal 不明原因事件数环比 = (Convert.ToDecimal(月事件数.Rows[5][0]) - 上月不明原因事件数) / 上月不明原因事件数 * 100;

            tmpstr = string.Format("{0}年{1}月，全国地震前兆台网共有{2}个台站{3}套仪器记录到不明原因事件{4}条，其中{5}个台站{6}套仪器记录到疑似前兆异常事件{7}条。本月不明原因事件环比{8}{9}%。表3.7.1和图3.7.1对本月记录到的不明原因事件按区域进行了划分统计；图3.7.2绘制了不明原因事件的分布图。", the_date.Year, the_date.Month, 月台站数.Rows[5][0], 月仪器数.Rows[5][0], 月事件数.Rows[5][0], 月疑似前兆异常台站数, 月疑似前兆异常仪器数, 月疑似前兆异常事件数, 不明原因事件数环比 < 0 ? "下降" : "增加", Math.Abs(Math.Round(不明原因事件数环比, 1))) + Environment.NewLine;
            wordapp.Selection.ParagraphFormat.set_Style("正文");
            wordapp.Selection.TypeText(tmpstr);

            DataTable 表3_7_1tmp = orahlper.GetDataTable(string.Format(@"select unitname, count(log_id), sum(decode(dura, 1, 1, 0)) short, sum(decode(dura, 2, 1, 0)) middle, sum(decode(dura, 3, 1, 0)) lon, sum(decode(dura, 0, 1, 0)) ongoing from (
select distinct D.UNITNAME, a.log_id, case when a.END_DATE is null then 0 when a.end_date - a.START_DATE >= 7 then 3 when a.end_date - a.start_date < 3 then 1 else 2 end as dura  from qzdata.qz_abnormity_log a, qzdata.qz_abnormity_evalist b, qzdata.qz_dict_stations c, qzdata.qz_abnormity_units d where {0} and a.stationid = b.stationid and a.pointid = b.pointid and B.SCIENCE != '辅助' and a.ab_id = 7 and A.STATIONID = C.STATIONID and C.UNITCODE = D.UNIT_CODE
) group by unitname order by count(log_id) desc", datestr));

            DataTable 各省疑似前兆异常事件数 = orahlper.GetDataTable(string.Format(@"select unitname, count(log_id) from (
select distinct B.LOG_ID, D.UNITNAME from qzdata.qz_abnormity_itemlog a, qzdata.qz_abnormity_log b, QZDATA.QZ_DICT_STATIONS c, qzdata.qz_abnormity_units d, qzdata.qz_abnormity_evalist e, qzdata.qz_abnormity_type2 f where A.LOG_ID = b.log_id and b.stationid = e.stationid and B.POINTID = e.pointid and E.SCIENCE != '辅助' and B.STATIONID = C.STATIONID and C.UNITCODE = D.UNIT_CODE and A.TYPE2_ID = F.TYPE2_ID and F.TYPE2_NAME = '疑似前兆异常' and B.AB_ID = 7 and {0}
) group by unitname", datestr));
            表3_7_1tmp.Columns.Add("疑似前兆异常", Type.GetType("System.Decimal"));
            for (int i = 0; i < 表3_7_1tmp.Rows.Count; i++)
            {
                object tmp = dthlper.ExtractRowByLeftFirstCol_SingleValue(各省疑似前兆异常事件数, 表3_7_1tmp.Rows[i][0].ToString());
                if (tmp == null)
                    表3_7_1tmp.Rows[i]["疑似前兆异常"] = 0;
                else
                    表3_7_1tmp.Rows[i]["疑似前兆异常"] = tmp;
            }
            {
                DataRow dr = 表3_7_1tmp.NewRow();
                dr[0] = "合计";
                for (int j = 1; j < 表3_7_1tmp.Columns.Count; j++)
                {
                    dr[j] = 表3_7_1tmp.Compute(string.Format("sum([{0}])", 表3_7_1tmp.Columns[j].ColumnName), "true");
                }
                表3_7_1tmp.Rows.Add(dr);
            }
            ta.AddTable(表3_7_1tmp, new string[] { "单位名称", "不明原因事件", "持续时间短（<3天）", "持续时间中（3-7天）", "持续时间长（>7天）", "尚未结束", "疑似异常事件" }, (int[])null, string.Format("表3.7.1 {0}年{1}月区域前兆台网不明原因事件统计", the_date.Year, the_date.Month));
            wordapp.Selection.TypeText(Environment.NewLine);
            ta.AddNullTable("图3.7.1  各区域台网记录到不明原因事件数量统计");
            wordapp.Selection.TypeText(Environment.NewLine);
            ta.AddNullTable(string.Format("图3.7.2  {0}年{1}月记录到不明原因事件台站（左）和疑似异常事件台站（右）分布图", the_date.Year, the_date.Month));
            wordapp.Selection.ParagraphFormat.set_Style("标题 3");
            wordapp.Selection.TypeText("3.7.1 不明原因事件的区域性和时域性" + Environment.NewLine);
            wordapp.Selection.ParagraphFormat.set_Style("正文");
            tmpstr = @"从表3.7.1、图3.7.1和图3.7.2中可看出，本月不明原因事件集中分布于以云南、四川为主的西南区域（分布图见3.7.1.1），以辽宁、河北、山东、山西为主的东北华北区域（分布图见3.7.1.2）和以甘肃、宁夏为主的中西部区域（分布图见3.7.1.3）。其中以云南区域发生不明原因事件与疑似异常事件最为频繁集中，与近期川滇一带地震活动较为活跃存在对应性。" + Environment.NewLine;

            wordapp.Selection.TypeText(tmpstr);
            wordapp.Selection.TypeText(Environment.NewLine);
            ta.AddNullTable(string.Format("图3.7.1.1  {0}年{1}月西南区域记录到不明原因事件（黄色三角）、疑似异常事件（红色三角）台站分布图", the_date.Year, the_date.Month));
            wordapp.Selection.TypeText(Environment.NewLine);
            ta.AddNullTable(string.Format("图3.7.1.2  {0}年{1}月东北华北区域记录到不明原因事件（黄色三角）、疑似异常事件（红色三角）台站分布图", the_date.Year, the_date.Month));
            wordapp.Selection.TypeText(Environment.NewLine);
            ta.AddNullTable(string.Format("图3.7.1.3  {0}年{1}月中西部区域记录到不明原因事件（黄色三角）、疑似异常事件（红色三角）台站分布图", the_date.Year, the_date.Month));
            wordapp.Selection.ParagraphFormat.set_Style("标题 3");
            wordapp.Selection.TypeText("3.7.2 持续时间中、短的不明原因事件的形态特点" + Environment.NewLine);
            wordapp.Selection.ParagraphFormat.set_Style("正文");

            DataTable 表3_7_2_1 = orahlper.GetDataTable(string.Format(@"select style, sum(decode(science,'形变',1,0)) xb,
sum(decode(science,'重力',1,0)) zl,
sum(decode(science,'地磁',1,0)) dc,
sum(decode(science,'地电',1,0)) dd,
sum(decode(science,'流体',1,0)) lt, 
count(log_id) total from(
select distinct decode(b.style, null, '其它', b.style) as style, a.log_id, c.science from qzdata.qz_abnormity_log a, qzdata.qz_abnormity_itemlog b, qzdata.qz_abnormity_evalist c where a.log_id = b.log_id and a.stationid = c.stationid and a.pointid = c.pointid and C.SCIENCE != '辅助' and a.ab_id = 7 and {0} and a.end_date - a.start_date < 7
) group by style order by total desc", datestr));
            {
                DataRow dr = 表3_7_2_1.NewRow();
                dr[0] = "合计";
                for (int j = 1; j < 表3_7_2_1.Columns.Count; j++)
                {
                    dr[j] = 表3_7_2_1.Compute(string.Format("sum([{0}])", 表3_7_2_1.Columns[j].ColumnName), "true");
                }
                表3_7_2_1.Rows.Add(dr);
            }
            ta.AddTable(表3_7_2_1, new string[] { "变化形态", "形变", "重力", "地磁", "地电", "流体", "合计" }, (int[])null, "表3.7.2.1 持续时间中、短的不明原因事件学科、变化形态分布");
            wordapp.Selection.TypeText("从表3.7.2.1中可看出，形变、地磁、地电和流体学科存在持续时间中、短的不明原因事件，其中以流体、形变学科的事件数量居多，变化形态多为突跳和阶变，也有部分为上升、下降、固体潮畸变等原因。持续时间中、短的不明原因事件可能与观测仪器、场地环境、自然环境的短时不稳定或原因未明的变动有关，也有可能是地壳应力瞬时变化的反映。" + Environment.NewLine);
            wordapp.Selection.ParagraphFormat.set_Style("标题 3");
            wordapp.Selection.TypeText("3.7.3 持续时间长的不明原因事件的形态特点" + Environment.NewLine);
            wordapp.Selection.ParagraphFormat.set_Style("正文");
            DataTable 表3_7_3_1 = orahlper.GetDataTable(string.Format(@"select style, sum(decode(science,'形变',1,0)) xb,
sum(decode(science,'重力',1,0)) zl,
sum(decode(science,'地磁',1,0)) dc,
sum(decode(science,'地电',1,0)) dd,
sum(decode(science,'流体',1,0)) lt, 
count(log_id) total from(
select distinct decode(b.style, null, '其它', b.style) as style, a.log_id, c.science from qzdata.qz_abnormity_log a, qzdata.qz_abnormity_itemlog b, qzdata.qz_abnormity_evalist c where a.log_id = b.log_id and a.stationid = c.stationid and a.pointid = c.pointid and C.SCIENCE != '辅助' and a.ab_id = 7 and {0} and a.end_date - a.start_date >= 7
) group by style order by total desc", datestr));
            {
                DataRow dr = 表3_7_3_1.NewRow();
                dr[0] = "合计";
                for (int j = 1; j < 表3_7_3_1.Columns.Count; j++)
                {
                    dr[j] = 表3_7_3_1.Compute(string.Format("sum([{0}])", 表3_7_3_1.Columns[j].ColumnName), "true");
                }
                表3_7_3_1.Rows.Add(dr);
            }
            ta.AddTable(表3_7_3_1, new string[] { "变化形态", "形变", "重力", "地磁", "地电", "流体", "合计" }, (int[])null, "表3.7.3.1 持续时间长的不明原因事件学科、变化形态分布");
            wordapp.Selection.TypeText("从表3.7.3.1中可看出，持续时间长的不明原因事件在变化形态上多为趋势转折、上升、下降、破年变等趋势性变化，此类不明原因事件往往与地质活动和应力场积累转移的中长期作用有关，对地质构造运动和地壳运动有一定程度的反映。" + Environment.NewLine + "持续时间长的不明原因事件时间进程图见图3.7.3.1。" + Environment.NewLine);
            wordapp.Selection.TypeText(Environment.NewLine);

            DataTable 不明进程图用 = orahlper.GetDataTable(string.Format(@"select unitname, stationname, decode(type2_name, '长期不明原因', '不明原因', '中短期不明原因', '不明原因', '疑似前兆异常', '疑似前兆异常') type2_name, start_date, end_date, len from(select distinct a.log_id, D.UNITNAME, E.STATIONNAME, F.TYPE2_NAME, A.START_DATE, a.end_date, round(a.end_date - a.start_date, 4) as len from qzdata.qz_abnormity_log a, qzdata.qz_abnormity_itemlog b, qzdata.qz_abnormity_evalist c, qzdata.qz_abnormity_units d, qzdata.qz_dict_stations e, qzdata.qz_abnormity_type2 f where B.TYPE2_ID = F.TYPE2_ID and a.log_id = b.log_id and A.STATIONID = c.stationid and a.pointid = c.pointid and C.SCIENCE != '辅助' and C.UNITCODE = D.UNIT_CODE and a.stationid = E.STATIONID and {0} and a.ab_id = 7 and (F.TYPE2_NAME = '长期不明原因' or f.type2_name = '中短期不明原因' or F.TYPE2_NAME = '疑似前兆异常') and a.end_date is not null and a.end_date - a.start_date >= 7 ) order by unitname", datestr));
            不明进程图用.Columns.Add("len2", Type.GetType("System.Decimal"));

            for (int i = 0; i < 不明进程图用.Rows.Count; i++)
            {
                double tmp;
                if ((tmp = (Convert.ToDateTime(不明进程图用.Rows[i]["start_date"]) - the_month_begin).TotalDays) < 0)
                {
                    不明进程图用.Rows[i]["len"] = Convert.ToDouble(不明进程图用.Rows[i]["len"]) + tmp;
                    不明进程图用.Rows[i]["start_date"] = the_month_begin;
                }
                if ((tmp = (Convert.ToDateTime(不明进程图用.Rows[i]["end_date"]) - the_month_end).TotalDays) > 0)
                {
                    不明进程图用.Rows[i]["len"] = Convert.ToDouble(不明进程图用.Rows[i]["len"]) - tmp;
                    不明进程图用.Rows[i]["end_date"] = the_month_end;
                }
                if (不明进程图用.Rows[i]["type2_name"].ToString() == "不明原因")
                    不明进程图用.Rows[i]["len2"] = 0;
                else
                {
                    不明进程图用.Rows[i]["len2"] = 不明进程图用.Rows[i]["len"];
                    不明进程图用.Rows[i]["len"] = 0;
                }
                
            }
            不明进程图用.Columns.Remove("type2_name");
            不明进程图用.Columns.Remove("end_date");
            不明进程图用.Columns["len"].ColumnName = "不明原因";
            不明进程图用.Columns["len2"].ColumnName = "疑似前兆异常";
            object[,] 不明进程图用tmp = dthlper.DataTableTo2DTable(不明进程图用);
            for (int i = 1; i < 不明进程图用tmp.GetLength(0); i++)
            {
                不明进程图用tmp[i, 2] = Convert.ToDateTime(不明进程图用tmp[i, 2]).ToOADate();
            }
            for (int i = 1, k; i < 不明进程图用tmp.GetLength(0); i++)
            {
                for (k = i; k < 不明进程图用tmp.GetLength(0) - 1; k++)
                {
                    if (不明进程图用tmp[k + 1, 0].ToString() != 不明进程图用tmp[i, 0].ToString())
                    {
                        break;
                    }
                    不明进程图用tmp[k + 1, 0] = null;
                }
                if (k == i)
                    continue;
                else
                {
                    不明进程图用tmp[k, 0] = new MERGEINTO(i, 0);
                    i = k;
                    continue;
                }
            }
            ta.AddTable(不明进程图用tmp, (int[])null);


            ta.AddNullTable("图3.7.3.1  持续时间长的不明原因事件时间进程图");
            wordapp.Selection.ParagraphFormat.set_Style("标题 3");
            wordapp.Selection.TypeText("3.7.4 与国内地震有关联的不明原因事件" + Environment.NewLine);
            wordapp.Selection.ParagraphFormat.set_Style("正文");
            wordapp.Selection.TypeText("详见第二章“地震事件回顾分析”。" + Environment.NewLine);


            #endregion
#endif
        }
    }
}