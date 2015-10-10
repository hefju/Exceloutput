using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System.IO;

namespace Exceloutput
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            output(dgv,"人事资料");
        }

        private void output(DataGridView dgv,string Title)
        {
            using (var p = new ExcelPackage())
            {
                var sheet = p.Workbook.Worksheets.Add("My Sheet");
                //合并单元格
                sheet.Cells[1,1,1,dgv.Columns.Count].Merge = true; 
                //标题列
                sheet.Cells[1,1].Value = Title;
                //列
                for (int i = 0; i < dgv.Columns.Count; i++)
                {
                    //列名
                    sheet.Cells[2, i+1].Value = dgv.Columns[i].Name.ToString();
                    //行
                    for (int j = 0; j < dgv.Rows.Count-1; j++)
                    {
                        sheet.Cells[j+3, i+1].Value = dgv.Rows[j].Cells[i].Value.ToString();
                    }

                }
                //样式
                //自动对齐
                sheet.Cells[1, 1, 1, dgv.Columns.Count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //自动条列宽
                sheet.Cells.AutoFitColumns();
                //底色
                sheet.Cells[2, 1, 2, dgv.Columns.Count].Style.Fill.PatternType = ExcelFillStyle.Solid; 
                sheet.Cells[2,1,2,dgv.Columns.Count].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                //边框
                sheet.Cells[1, 1, dgv.Rows.Count+1, dgv.Columns.Count].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                sheet.Cells[1, 1, dgv.Rows.Count+1, dgv.Columns.Count].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                sheet.Cells[1, 1, dgv.Rows.Count+1, dgv.Columns.Count].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                sheet.Cells[1, 1, dgv.Rows.Count+1, dgv.Columns.Count].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                p.SaveAs(new FileInfo(@"D:\output.xlsx"));
                
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            DataTable table = GetTable();
            dgv.DataSource = table;
        }
        private DataTable GetTable()
        {
            // Here we create a DataTable with four columns.
            DataTable table = new DataTable();
            table.Columns.Add("Dosage", typeof(int));
            table.Columns.Add("Drug", typeof(string));
            table.Columns.Add("Patient", typeof(string));
            table.Columns.Add("Date", typeof(DateTime));

            // Here we add five DataRows.
            table.Rows.Add(25, "Indocin", "David", DateTime.Now);
            table.Rows.Add(50, "Enebrel", "Sam", DateTime.Now);
            table.Rows.Add(10, "Hydralazine", "Christoff", DateTime.Now);
            table.Rows.Add(21, "Combivent", "Janet", DateTime.Now);
            table.Rows.Add(100, "Dilantin", "Melanie", DateTime.Now);
            return table;
        }
        
    }
}
