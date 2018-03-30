using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;

namespace school_scale
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string filepath = @"C:\word_test\123.docx";
            string filepath2 = @"C:\word_test\1234.docx";

            using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.CreateFromTemplate(filepath))
            {
                
                Table tb = wordprocessingDocument.MainDocumentPart.Document.Body.Elements<Table>().First();
                DataTable dt = new DataTable();
                dt.Columns.Add("1");
                dt.Columns.Add("2");
                dt.Rows.Add(new string[] { "3", "4" });
                TableRowInsert(tb, dt);

                //TableCellMerge(tb, 3,2,1,1);
                //wordprocessingDocument.SaveAs(filepath2).Close();

            }

             using (WordprocessingDocument doc = WordprocessingDocument.CreateFromTemplate(filepath))
            {
                Dictionary<string, string> dic = new Dictionary<string, string>();
                dic.Add("aaa", "haha");
                this.replaceTags(doc.MainDocumentPart.Document.Body, dic);
                doc.SaveAs(filepath2).Close();

            }
        }

        private void TableColumnInsert(Table tb, DataTable data)
        {
            for (int j = 0; j < data.Rows.Count; j++)
            {
                DataRow dr = data.Rows[j];

                for (int i = 0; i < data.Columns.Count; i++)
                {
                    tb.Elements<TableRow>().ElementAt(i).AppendChild<TableCell>(new TableCell(new TableCellProperties(new Paragraph(new Run(new Text(dr[i].ToString()))))));

                }
            }
        }

        private void TableRowInsert(Table tb, DataTable data)
        {


            //if (tb.Elements<GridColumn>().Count() != data.Columns.Count)
            //{
            //    throw new Exception("目標表格與dataTable不符");
            //}
            foreach (DataRow dr in data.Rows)
            {
                TableRow tr = new TableRow();

                for (int i = 0; i < data.Columns.Count; i++)
                {
                    tr.AppendChild<TableCell>(new TableCell(new TableCellProperties(new Paragraph(new Run(new Text(dr[i].ToString()))))));
                }
                tb.AppendChild<TableRow>(tr);
            }

        }
        private void TableCellMerge(Table tb, int x1, int y1, int x2, int y2)
        {
            if (x1 == x2 && y1 == y2)
            {
                return;

            }
            int minX = Math.Min(x1, x2);
            int maxX = Math.Max(x1, x2);
            int minY = Math.Min(y1, y2);
            int maxY = Math.Max(y1, y2);
            int mergeColumnCount = maxX - minX + 1;
            TableRow tr;
            TableCell tc;
            TableCellProperties tcpr;
            GridSpan gs;
            VerticalMerge vm;
            //horizontal merge
            for (int j = minY; j <= maxY; j++)
            {
                tr = tb.Elements<TableRow>().ElementAt<TableRow>(j);
                tc = tr.Elements<TableCell>().ElementAt<TableCell>(minX);
                tcpr = tc.Elements<TableCellProperties>().FirstOrDefault();
                gs = new GridSpan() { Val = mergeColumnCount };
                if (tcpr != null)
                {
                    tcpr.AppendChild<GridSpan>(gs);

                }
                else
                {
                    tc.AppendChild<TableCellProperties>(tcpr);
                    tcpr.AppendChild<GridSpan>(gs);
                }

                for (int i = minX + 1; i <= maxX; i++)
                {
                    tr.Elements<TableCell>().ElementAt<TableCell>(minX + 1).Remove();
                }

            }
            //vertical merge
            if (maxY != minY)
            {
                tr = tb.Elements<TableRow>().ElementAt<TableRow>(minY);
                tc = tr.Elements<TableCell>().ElementAt<TableCell>(minX);
                tcpr = tc.Elements<TableCellProperties>().FirstOrDefault();
                vm = new VerticalMerge() { Val = MergedCellValues.Restart };
                if (tcpr != null)
                {
                    tcpr.AppendChild<VerticalMerge>(vm);
                }
                else
                {
                    tcpr.AppendChild<VerticalMerge>(vm);
                    tc.AppendChild<TableCellProperties>(tcpr);
                }

                for (int j = minY + 1; j <= maxY; j++)
                {
                    tr = tb.Elements<TableRow>().ElementAt<TableRow>(j);
                    tc = tr.Elements<TableCell>().ElementAt<TableCell>(minX);
                    tcpr = tc.Elements<TableCellProperties>().FirstOrDefault();

                    vm = new VerticalMerge() { Val = MergedCellValues.Continue };
                    if (tcpr != null)
                    {
                        tcpr.AppendChild<VerticalMerge>(vm);

                    }
                    else
                    {
                        tcpr.AppendChild<VerticalMerge>(vm);
                        tc.AppendChild<TableCellProperties>(tcpr);
                    }
                }
            }
        }
        private void replaceTags(Body body, Dictionary<string, string> dic)
        {
            var tables = body.Elements<Table>();
            var paragraphs = body.Elements<Paragraph>();


            foreach (var pair in dic)
            {
                foreach (Paragraph p in paragraphs)
                {
                    Regex regex = new Regex(string.Format("{{{0}}}", pair.Key));
                    if (regex.IsMatch(p.InnerText))
                    {
                        string newText = regex.Replace(p.InnerText, pair.Value);

                        Run r = (Run)p.Elements<Run>().First().CloneNode(true);
                        Text text = r.Elements<Text>().First();
                        text.Text = (newText);


                        p.RemoveAllChildren();
                        p.AppendChild<Run>(r);

                    }
                }


                foreach (Table t in tables)
                {
                    foreach (TableRow tr in t.Elements<TableRow>())
                    {
                        foreach (TableCell tc in tr.Elements<TableCell>())
                        {
                            foreach (Paragraph p in tc.Elements<Paragraph>())
                            {
                                Regex regex = new Regex(string.Format("{{{0}}}", pair.Key));
                                if (regex.IsMatch(p.InnerText))
                                {
                                    string newText = regex.Replace(p.InnerText, pair.Value);

                                    Run r = (Run)p.Elements<Run>().First().CloneNode(true);
                                    Text text = r.Elements<Text>().First();
                                    text.Text = (newText);


                                    p.RemoveAllChildren();
                                    p.AppendChild<Run>(r);

                                }

                            }
                        }
                    }
                }


            }
        }
    }
}
