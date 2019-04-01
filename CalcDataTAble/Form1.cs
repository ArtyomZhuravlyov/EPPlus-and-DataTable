using ClosedXML.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CalcDataTAble
{
    public partial class Form1 : Form
    {
        DataTable dt = new DataTable();

        DataTable dtBig = new DataTable();
        int RowsDtSmall = 100;
        int timeCount = 10;
        public Form1()
        {
            InitializeComponent();
            dt.Columns.Add("Дата");
            dt.Columns.Add("Время");
            dt.Columns.Add("Тип события");
            dt.Columns.Add("Описание");
            dt.Columns.Add("Описание2");
            //dt.Columns.Add("Описание2");
            for (int i = 0; i <= RowsDtSmall; i++)
            {
                dt.Rows.Add(i, i, i, i, "");
            }

            string n1 = "DD";
            string n2 = "Dd";
            string n3 = "Aafff";

            int a1 = String.Compare(n1, n2);
            int a2 = String.Compare(n1, n2, true);
            int a3 = String.Compare(n1, n2, false);
            int a4 = String.Compare(n1, n3, false);
            int a5 = String.Compare(n1, n3);
            int a6 = String.Compare(n3, n1);

            dtBig.Columns.Add("Дата");
            dtBig.Columns.Add("Время");
            dtBig.Columns.Add("Тип события");
            dtBig.Columns.Add("Описание");
            //dtBig.Columns.Add("Описание1");
            //dtBig.Columns.Add("Описание2");
            //dtBig.Columns.Add("Описание3");
            //dtBig.Columns.Add("Описание4");
            //dtBig.Columns.Add("Описание5");
            //dtBig.Columns.Add("Описание6");
            //dtBig.Columns.Add("Описание7");
            //dtBig.Columns.Add("Описание8");
            //dtBig.Columns.Add("Описание9");
            //dtBig.Columns.Add("Описание0");
            //dtBig.Columns.Add("Описание12");
            //dtBig.Columns.Add("Описание11");
            for (int i = 0; i <= 1000; i++)
            {
                 dtBig.Rows.Add("ghbdtnnnnnn", i, "isdfsd", "idsdddddddddddddddddd");
                //dtBig.Rows.Add(i, i, "", "", "", "", "", i, i, "", "", "", "","", "", "");
            }
            toolStripStatusLabel1.Text = "Базы готовы к работе";

        }

        private void button1_Click(object sender, EventArgs e)
        {
          //  ClearDataTable(dt); 
            dataGridView1.DataSource = dtBig;
        }

        //ClosedXML
        private void button2_Click(object sender, EventArgs e)
        {
            DateTime now1 = DateTime.Now;
            //Console.WriteLine("T: {0:T}", now);
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dtBig, "Customers");
                DateTime now2 = DateTime.Now;
                TimeSpan timeDateTime = now2 - now1;

                // toolStripStatusLabel3.Text = timeDateTime.ToString(); //"T: {0:T}"
                // (TimeSpan.TryParseExact(value, "ss\\.fff", null, out interval))
                toolStripStatusLabel2.Text = timeDateTime.ToString("ss\\.fff"); //"T: {0:T}"
               // MessageBox.Show("закончил обработку большой базы");

                now1 = DateTime.Now;
                wb.SaveAs(@"C:\Users\artem zhuravlev\Desktop\" + "DataGridViewExport.xlsx");
                now2 = DateTime.Now;
                timeDateTime = now2 - now1;
                label1.Text = timeDateTime.ToString("mm\\:ss\\.ff"); //"T: {0:T}"
            }


        }

        //EpPLUS BigTable
        async private void button3_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                timer1.Enabled = true;
                timer1.Start();
            }
            await Task.Run(() =>
            {
                File.Delete(@"C:\Users\artem zhuravlev\Desktop\" + "EpPlusBIG.xlsx");
                using (ExcelPackage pck = new ExcelPackage(new System.IO.FileInfo(@"C:\Users\artem zhuravlev\Desktop\" + "EpPlusBIG.xlsx")))
                {
                    int pos = 1;
                    ExcelWorksheet ws;
                    DateTime now1 = DateTime.Now;

                        ws = pck.Workbook.Worksheets.Add("Accounts");

                        ws.Cells["A" + pos].LoadFromDataTable(dtBig, true);
                    //GC.Collect();
                    //GC.WaitForPendingFinalizers();

                    //GC.Collect();
                    //GC.WaitForPendingFinalizers();

                    string modelRange = "A1:D1000";

                    var modelTable = ws.Cells[modelRange];
                    modelTable.AutoFitColumns();

                    // Assign borders
                    modelTable.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    modelTable.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    modelTable.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    //modelTable.Style.Border.Bottom.Color.SetColor(Color.AliceBlue);
                    //modelTable.Style.Border.Top.Color.SetColor(Color.Red);
                    //modelTable.Style.Border.Bottom.Color.SetColor(Color.Green);
                    //modelTable.Style.Border.Left.Color.SetColor(Color.Blue);
                    //modelTable.Style.Border.Right.Color.SetColor(Color.Yellow);

                    
                    modelTable.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                     modelTable = ws.Cells["A1:D1"];
                    modelTable.Style.Font.Bold = true;
                    // modelTable.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    modelTable.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    modelTable.Style.Fill.BackgroundColor.SetColor(Color.Gray);

                    // modelTable.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    // modelTable.Style.Border.
                    //cells.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
                    //cells.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                    //cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    //cells.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    //cells.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    //cells.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;

                    //cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlMedium;
                    //cells.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlMedium;
                    //cells.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;
                    //cells.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlMedium;

                    pck.Save();
                    Process.Start(@"C:\Users\artem zhuravlev\Desktop\" + "EpPlusBIG.xlsx");
                    DateTime now2 = DateTime.Now;
                    TimeSpan timeDateTime = now2 - now1;
                    toolStripStatusLabel6.Text = timeDateTime.ToString("mm\\:ss\\.ff"); //"T: {0:T}"

                    dtBig.Clear();
                   // dataGridView1.DataSource = dtBig;
                    // timer1.Enabled = false;
                    timer1.Stop();
                    //MessageBox.Show("закончил ");
                }
              });


            //using (ExcelPackage pck = new ExcelPackage(new System.IO.FileInfo(@"C:\Users\artem zhuravlev\Desktop\" + "EpPLUS.xlsx")))
            //{
            //    int pos = 1;
            //    ExcelWorksheet ws;
            //    if (pck.Workbook.Worksheets.Count == 0)
            //        ws = pck.Workbook.Worksheets.Add("Accounts");
            //    else
            //        ws = pck.Workbook.Worksheets[1];
            //    ws.Cells["A" + pos].LoadFromDataTable(dt, true);

            //    pos += dt.Rows.Count;
            //    ws.Cells["A" + pos].LoadFromDataTable(dt, true);
            //    pck.Save();
            //}

        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            //if (wb2.Worksheets.Count==0)
            //wb2.Worksheets.Add(dt, "asd");
            //else wb2.Worksheets.Worksheet. // = pck.Workbook.Worksheets[1];
            //if (wb2.Worksheets.Count == 0)
            //    wb2.Worksheets.Add(dt, "asd");
            //else
            //{
            //    wb2.Worksheets.First();
            //    wb2.Worksheets.Add(dt);
            //}
            // wb2.Worksheets.Add(dt, "asd");


            //var worksheet = workbook.Worksheets.Add("Лист1");
            //var workbook = new XLWorkbook();
            //IXLWorksheet workSheet = workbook.Worksheet(1);

            //dt.TableName = "sdfsdf";
            //if() 


            //var workSheet = xlWorkBook.Worksheets.First(ws => ws.Name == dataTable.TableName);
            // wb2.Worksheets.First().Cell(wb2.Worksheets.First().RowCount().ToString()).InsertTable(dt);

            //IXLRangeRow row = sheet.Range(rowIdx, 1, rowIdx, sheet.ColumnsUsed().Count()).Rows().First();
            //IXLRangeRows newRows = row.InsertRowsBelow(detailTable.Rows.Count + 1);
            //newRows.First().Cell(1).InsertTable(dt);
            //wb2.Worksheets.First().Cell(RowCount, 1).InsertTable(dt);

            //closedXML
            DateTime now1 = DateTime.Now;
            int RowCount = 1;
            //closedXML
            using (XLWorkbook wb2 = new XLWorkbook())
            {
                for (int i = 0; i < timeCount; i++)
                {
                    if (wb2.Worksheets.Count == 0)
                    {
                        wb2.Worksheets.Add(dt, "sdfsdf");
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        wb2.SaveAs(@"C:\Users\artem zhuravlev\Desktop\" + "Export.xlsx");
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        RowCount += (RowsDtSmall + 2);
                    }
                    else
                    {
                        wb2.Worksheets.First().Cell(RowCount, 1).InsertTable(dt);
                        RowCount += (RowsDtSmall + 2);
                        GC.Collect();
                        GC.WaitForPendingFinalizers();

                        wb2.Save();
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
                }
                wb2.SaveAs(@"C:\Users\artem zhuravlev\Desktop\" + "Export.xlsx");
                DateTime now2 = DateTime.Now;
                TimeSpan timeDateTime = now2 - now1;
                toolStripStatusLabel3.Text = timeDateTime.ToString("mm\\:ss\\.ff"); //"T: {0:T}"
            }


            /*
            DateTime now1 = DateTime.Now;
            for (int i = 0; i < 50; i++)
            {
                if (wb2.Worksheets.Count == 0)
                    wb2.Worksheets.Add(dt, "sdfsdf");
                else wb2.Worksheets.First().Cell(RowCount, 1).InsertTable(dt);
                RowCount += 2001;
            }
            
            // toolStripStatusLabel3.Text = timeDateTime.ToString(); //"T: {0:T}"
            // (TimeSpan.TryParseExact(value, "ss\\.fff", null, out interval))

            DateTime now2 = DateTime.Now;
            TimeSpan timeDateTime = now2 - now1;
            toolStripStatusLabel3.Text = timeDateTime.ToString("mm\\:ss\\.ff"); //"T: {0:T}"
            //MessageBox.Show("Закончилось");

            now1 = DateTime.Now;
            wb2.SaveAs(@"C:\Users\artem zhuravlev\Desktop\" + "Export.xlsx");
            now2 = DateTime.Now;
             timeDateTime = now2 - now1;
            label2.Text = timeDateTime.ToString("mm\\:ss\\.ff"); //"T: {0:T}"
            */

            //  wb2.Worksheets.First().Cell(1).InsertTable(dt);
            // wb2.Worksheets.Add(dt, "sdfsdf");
            //wb2.Worksheets.Add
            //ws.Name = "Enter_a_Name_same_as_Excelsheetname";


            //https://stackoverflow.com/questions/43458813/outofmemory-errors-using-large-datatables-with-closedxml?rq=1
            //https://stackoverflow.com/questions/36549264/closedxml-excel-document-reported-as-corrupted
        }

        private void button7_Click(object sender, EventArgs e)
        {
            File.Delete(@"C:\Users\artem zhuravlev\Desktop\" + "DataGridViewExport.xlsx");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            File.Delete(@"C:\Users\artem zhuravlev\Desktop\" + "Export.xlsx");
        }


        //EpPLUS Add Parths
        async private void button8_Click(object sender, EventArgs e)
        {

            //ws.Column(1).Style.Numberformat.Format = "$"; //задание формата для столбца
            //sheet.Cells[row, col].Value = "Наименование";

            if (checkBox1.Checked == true)
            {
                timer1.Enabled = true;
                timer1.Start();
            }
            await Task.Run(() =>
                {
                    using (ExcelPackage pck = new ExcelPackage(new System.IO.FileInfo(@"C:\Users\artem zhuravlev\Desktop\" + "EpPlusAddPaths.xlsx")))
                    {
                        int pos = 1;
                        ExcelWorksheet ws;
                        DateTime now1 = DateTime.Now;
                        ws = pck.Workbook.Worksheets.Add("Accounts");
                       //ws = pck.Workbook.Worksheets[1];
                        for (int i = 0; i < timeCount; i++)
                        {
                            ws.Cells["A" + pos].LoadFromDataTable(dt, false);
                            pos += RowsDtSmall + 1;
                            //ws.Cells["A" + pos].LoadFromDataTable(dt, true);
                            //pck.Save();
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                        }
                        pck.Save();


                        DateTime now2 = DateTime.Now;
                        TimeSpan timeDateTime = now2 - now1;
                        toolStripStatusLabel6.Text = timeDateTime.ToString("mm\\:ss\\.ff"); //"T: {0:T}"
                       
                        // timer1.Enabled = false;
                        timer1.Stop();
                        MessageBox.Show("закончил ");
                    }
                });
        }

        private void button9_Click(object sender, EventArgs e)
        {
            File.Delete(@"C:\Users\artem zhuravlev\Desktop\" + "EpPlusAddPaths.xlsx");
            File.Delete(@"C:\Users\artem zhuravlev\Desktop\" + "EpPlusBIG.xlsx");
        }

        long memory = 0;
        long countTimer = 0;
        private void timer1_Tick(object sender, EventArgs e)
        {
            countTimer++;
            memory +=  GC.GetTotalMemory(true)/ 1048576;
            label3.Text = (memory / countTimer).ToString();

        }

        /// <summary>
        /// Clear Empty Columns DataTable
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public DataTable ClearDataTable(DataTable dt)
        {
            bool IsColumnEmpty;

            for (int i = dt.Columns.Count - 1; i >= 0; i--)
            {
                IsColumnEmpty = dt.AsEnumerable().All(dr => string.IsNullOrEmpty(dr[i].ToString()));
                if (IsColumnEmpty)
                    dt.Columns.RemoveAt(i);
            }

            Task.Run(() =>
            {
                int i = 3;
                Label label5 = new Label();
                Form aa = new Form();
                aa.Show();
               // label5.Location = 
                aa.Controls.Add(label5);
                while (true)
                    label5.Text = i.ToString();
            });

            return dt;
        }
    }
}
