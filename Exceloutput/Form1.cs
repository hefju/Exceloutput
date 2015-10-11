using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
//using OfficeOpenXml;
//using OfficeOpenXml.Drawing;
//using OfficeOpenXml.Drawing.Chart;
//using OfficeOpenXml.Style;
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
            using (var p = new OfficeOpenXml.ExcelPackage())
            {
                var sheet = p.Workbook.Worksheets.Add("Sheet1");
                #region 设置标题
                var title_width = 8; //默认纵向8个单元格 //dgv.Columns.Count > 8 ? dgv.Columns.Count : 8;
                //合并标题的单元格
                sheet.Cells[1, 1, 1, title_width].Merge = true;
                //居中对题
                sheet.Cells[1, 1, 1, dgv.Columns.Count].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[1, 1, 1, dgv.Columns.Count].Style.Font.Bold = true;
                //设置标题列
                sheet.Cells[1, 1].Value = Title;//第一行标题
                #endregion

                #region 内容
                var realColumn = 0;//列数(不包含隐藏列)
                for (int i = 0; i < dgv.Columns.Count; i++)
                {
                    if (!dgv.Columns[i].Visible) continue;//不输出隐藏列
                    realColumn++;
                    //表头
                    sheet.Cells[2, realColumn].Value = dgv.Columns[realColumn].HeaderText.ToString();//第二行表头
                    //数据行
                    for (int j = 0; j < dgv.Rows.Count - 1; j++)
                    {
                        object obj = dgv.Rows[j].Cells[realColumn].Value;
                        if (obj != null)
                            sheet.Cells[j + 3, realColumn].Value = obj.ToString();//第三行数据行
                    }
                }
                #endregion

                #region 边框和表头样式
                //自动条列宽
                sheet.Cells.AutoFitColumns();
                //表头底色
                sheet.Cells[2, 1, 2, realColumn].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                sheet.Cells[2, 1, 2, realColumn].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                //边框
                sheet.Cells[1, 1, dgv.Rows.Count + 1, realColumn].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                sheet.Cells[1, 1, dgv.Rows.Count + 1, realColumn].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                sheet.Cells[1, 1, dgv.Rows.Count + 1, realColumn].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                sheet.Cells[1, 1, dgv.Rows.Count + 1, realColumn].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                #endregion

                #region 建立临时文件并打开文件
                string mikecat_filename = Path.GetTempFileName().Replace(".tmp", "") + ".xlsx";
                p.SaveAs(new FileInfo(mikecat_filename));
                System.Diagnostics.Process.Start(mikecat_filename);
                #endregion
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            DataTable table = GetTable();
            dgv.DataSource = table;
            dgv.Columns["ID"].Visible = false;
        }
        private DataTable GetTable()
        {
            // Here we create a DataTable with four columns.
            DataTable table = new DataTable();
            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("Dosage", typeof(int));
            table.Columns.Add("Drug", typeof(string));
            table.Columns.Add("Patient", typeof(string));
            table.Columns.Add("Date", typeof(DateTime));

            // Here we add five DataRows.
            table.Rows.Add(10,25, "Indocin", "David", DateTime.Now);
            table.Rows.Add(10,50, "Enebrel", "Sam", DateTime.Now);
            table.Rows.Add(10,10, "Hydralazine", "Christoff", DateTime.Now);
            table.Rows.Add(10,21, "Combivent", "Janet", DateTime.Now);
            table.Rows.Add(10,100, "Dilantin", "Melanie", DateTime.Now);
            return table;
        }
        
    }
}
