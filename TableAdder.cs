using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using word = Microsoft.Office.Interop.Word;
using System.Data.OracleClient;
using System.Data;
using Audit;


namespace ReportGen
{
    class TableAdder
    {
        public bool enable = true;
        public bool consoleoutput = true;
        public word.Application wordapp = null;
        public word.Document worddoc = null;
        private bool addnull = false;
        public int fonttype = 2;

        public void AddDupFoldTable(int dup, DataTable dt, string[] newcolname, int[] colwidth, string title = null)
        {
            DataTableHelper dth = new DataTableHelper();
            object[,] table = dth.DataTableTo2DTable(dt);
            if (newcolname != null)
            {
                for (int j = 0; j < newcolname.GetLength(0); j++)
                {
                    table[0, j] = newcolname[j];
                }
            }
            table = dth.DupFold2DTable_HasColHeader(dup, table);
            if (colwidth == null)
            {
                AddTable(table, null, title);
            }
            else
            {
                int cn = colwidth.GetLength(0);
                int[] _wid = new int[cn * 2];
                for (int i = 0; i < cn; i++)
                {
                    _wid[i] = _wid[cn + i] = colwidth[i];
                }
                AddTable(table, _wid, title);
            }
        }

        public void AddDupFoldTable(int dup, DataTable dt, int[] colwidth, string title = null)
        {
            AddDupFoldTable(dup, dt, null, colwidth, title);
        }

        public void AddNullTable(string title)
        {
            addnull = true;
            AddTable(wordapp, worddoc, (object[,])null, null, title);
            addnull = false;
        }
        public void AddTable(object[,] t, int[] colwidth, string title = null)
        {
            AddTable(this.wordapp, this.worddoc, t, colwidth, title);
        }
        public void AddTable(DataTable dt, string[] newcolname, int[] colwidth, string title = null)
        {
            AddTable(this.wordapp, this.worddoc, dt, newcolname, colwidth, title);
        }
        
        public void AddTable(word.Application wdapp, word.Document wddoc, OracleConnection oracon, string strsql, string title = null)
        {
            OracleDataAdapter oda = new OracleDataAdapter(strsql, oracon);
            DataTable dt = new DataTable();
            oda.Fill(dt);
            this.AddTable(wdapp, wddoc, dt, null, null, title);
        }
        public void AddTable(word.Application wdapp, word.Document wddoc, DataTable dt, string[] newcolname, int[] colwidth, string title = null)
        {
            DataTableHelper dth = new DataTableHelper();
            object[,] table = dth.DataTableTo2DTable(dt);
            if (newcolname != null)
            {
                for (int j = 0; j < newcolname.GetLength(0); j++)
                {
                    table[0, j] = newcolname[j];
                }
            }
            this.AddTable(wdapp, wddoc, table, colwidth, title);
        }

        public void AddTable(word.Application wdapp, word.Document wddoc, object[,] t, int[] colwidth, string title = null)
        {
            if (!enable)
                return;
            if (title != null)
            {
                wordapp.Selection.ParagraphFormat.set_Style("图例表例");
                wordapp.Selection.TypeText(title);
                wordapp.Selection.TypeParagraph();
                wordapp.Selection.ParagraphFormat.set_Style("正文");
            }
            if (addnull)
                return;

            if (fonttype == 1)
            {
                wordapp.Selection.ParagraphFormat.set_Style("表格字体");
            }
            else
            {
            }
            int rowcount = t.GetLength(0), colcount = t.GetLength(1);
            word.Table table = wddoc.Tables.Add(wdapp.Selection.Range, rowcount, colcount);
            table.Borders[word.WdBorderType.wdBorderHorizontal].Visible = true;
            table.Borders[word.WdBorderType.wdBorderVertical].Visible = true;
            table.Borders.OutsideLineStyle = word.WdLineStyle.wdLineStyleSingle;
            if (colwidth != null)
            {
                for (int j = 0; j < colcount; j++)
                {
                    table.Columns[j+1].Width = colwidth[j];
                }
            }
            for (int i = 0; i < rowcount; i++)
            {
                for (int j = 0; j < colcount; j++)
                {
                    if (fonttype == 2)
                    {
                        if (i == 0)
                        {
                            table.Cell(i + 1, j + 1).Range.ParagraphFormat.set_Style("表格列首");
                        }
                        else
                        {
                            table.Cell(i + 1, j + 1).Range.ParagraphFormat.set_Style("表格内容");
                        }
                    }
                    if (t[i, j] == null)
                    {
                    }
                    else if (t[i, j] is MERGEINTO)
                    {
                        //         table.Cell(i + 1, j + 1).Range.Shading.BackgroundPatternColor = word.WdColor.wdColorBlue;
                        table.Cell(i + 1, j + 1).Merge(table.Cell(((MERGEINTO)t[i, j]).row + 1, ((MERGEINTO)t[i, j]).col + 1));
                    }
                    else
                    {
                        table.Cell(i + 1, j + 1).Range.Text = t[i, j].ToString();
                    }
                }
            }
            //    table.Cell(1, 5).Range.Shading.BackgroundPatternColor = word.WdColor.wdColorBlue;

            table.Rows.HeightRule = word.WdRowHeightRule.wdRowHeightAuto;
            //      table.Columns.AutoFit();
            table.Range.ParagraphFormat.Alignment = word.WdParagraphAlignment.wdAlignParagraphCenter;
            table.Range.Cells.VerticalAlignment = word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            table.Rows.Alignment = word.WdRowAlignment.wdAlignRowCenter;
            wdapp.Selection.GoTo(word.WdGoToItem.wdGoToLine, word.WdGoToDirection.wdGoToLast);
      
            wordapp.Selection.ParagraphFormat.set_Style("正文");
      
            if (consoleoutput)
            {
                Console.WriteLine("added table '{0}'", title);
            }
        }
    }
}
