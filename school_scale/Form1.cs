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
using System.Data.SqlClient;


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
            string filepath = @"C:\word_test\template.docx";
            string filepath2 = @"C:\word_test\1234.docx";
            DataTable dt_f06 = new DataTable();
            DataTable dt_f06s = new DataTable();
            DataTable dt_f06d = new DataTable();
            


            using (WordprocessingDocument doc = WordprocessingDocument.CreateFromTemplate(filepath))
            {
                Table table6 = doc.MainDocumentPart.Document.Body.Elements<Table>().ElementAt<Table>(3);
                foreach (DataRow dr_f06 in dt_f06.Rows)
                {
                    string groupID = dr_f06.Field<object>("GroupID").ToString();
                    DataRow[] a = dt_f06s.Select(string.Format(("GroupID = '{0}'"), groupID));

                    DataTable dt = new DataTable();
                    dt.Columns.AddRange(new DataColumn[] { new DataColumn("major"), new DataColumn("countType"), new DataColumn("firstGrade1")
                    , new DataColumn("firstGrade2"), new DataColumn("secondGrade1"), new DataColumn("secondGrade2"), new DataColumn("thirdGrade1"), new DataColumn("thirdGrade2"), new DataColumn("sum") });

                    //mock up 
                    dt = new DataTable();
                    dt.Columns.AddRange(new DataColumn[] { new DataColumn("major"), new DataColumn("countType") });
                    DataRow dr = dt.NewRow();
                    dr["major"] = "普通科";
                    dr["countType"] = "班級數";
                    dt.Rows.Add(dr);
                     dr = dt.NewRow();
                    dr["major"] = "普通科";
                    dr["countType"] = "班級數";
                    dt.Rows.Add(dr);
                     dr = dt.NewRow();
                    dr["major"] = "普通科";
                    dr["countType"] = "班級數";
                    dt.Rows.Add(dr);
                    //DocProcessor.TableColumnInsert(table6, dt);
                    //DocProcessor.TableRowInsert(table6, dt);
                    //DocProcessor.DeleteTableColumn(table6, 1);
                    //DocProcessor.DeleteTableRow(table6, 5);
                    DocProcessor.DuplicateElement(table6, new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                    doc.SaveAs(filepath2).Close();

                }
            }
        }


    }
}
