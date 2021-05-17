using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WinFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();//測資表格
            DataTable dtDraw = new DataTable();//畫圖表格
            OpenFileDialog path = new OpenFileDialog();

            path.Filter = "Excel 工作表|*.xlsx|所有檔案|*.*";//設定檔案型別
            path.FileName = "";//設定預設檔名
            path.DefaultExt = "xlsx";//設定預設格式（可以不設）
            path.AddExtension = true;//設定自動在檔名中新增副檔名

            path.ShowDialog();

            if (path.FileName != "")
            {
                filetext.Text = System.IO.Path.GetFullPath(path.FileName);
                //string[] lines = System.IO.File.ReadAllLines(filetext.Text);

                XSSFWorkbook workbook = new XSSFWorkbook(path.FileName);
                ISheet sheet = workbook.GetSheetAt(3);//第四標籤頁(Df10um)

                //DataTable table = new DataTable();
                //由第一列取標題做為欄位名稱
                IRow headerRow = sheet.GetRow(0);
                int cellCount = headerRow.LastCellNum;
                //for (int i = headerRow.FirstCellNum; i < cellCount; i++)
                //    //以欄位文字為名新增欄位，此處全視為字串型別以求簡化

                //    dt.Columns.Add(
                //      new DataColumn(headerRow.GetCell(i).StringCellValue));

                //略過第零列(標題列)，一直處理至最後一列
                List<string> titles = new List<string>();
                List<Test> tests = new List<Test>();

                for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);

                    titles.Add(row.GetCell(0).ToString());
                    //if (row == null) continue;
                    //DataRow dataRow = table.NewRow();
                    //依先前取得的欄位數逐一設定欄位內容
                    //for (int j = row.FirstCellNum; j < cellCount; j++)
                    //{
                    //    if (row.GetCell(j) != null)
                    //        //如要針對不同型別做個別處理，可善用.CellType判斷型別
                    //        //再用.StringCellValue, .DateCellValue, .NumericCellValue...取值

                    //        //此處只簡單轉成字串

                    //        dataRow[j] = row.GetCell(j).ToString();
                    //}

                    //table.Rows.Add(dataRow);
                }

                foreach (var title in titles)//處理A欄標題
                {
                    var temp = title.Split('_');


                    tests.Add(new Test
                    {
                        Dend = Regex.Replace(temp[4], "[^0-9]", ""),
                        Ltp = Regex.Replace(temp[5], "[^0-9]", "")

                    });

                }

                for (int i = 0; i < tests.Count; i++)
                {
                    IRow colb = sheet.GetRow(i + 1);//處理B欄 +1跳過標題 
                    tests[i].Pow_eff = colb.GetCell(1).ToString();//1~20列的B欄位
                }

                //gridview顯示標題
                List<string> dataTabletitle = new List<string> { "Dend", "Ltp", "Pow_eff" };
                List<string> drawTabletitle = new List<string> { "1", "5", "10", "15", "20" };

                //    dt.Columns.Add(
                //      new DataColumn(headerRow.GetCell(i).StringCellValue));

                foreach (var title in dataTabletitle)
                {
                    dt.Columns.Add(new DataColumn(title));
                }

                dtDraw.Columns.Add(new DataColumn());//(0,0)是空白

                foreach (var title in drawTabletitle)
                {
                    dtDraw.Columns.Add(new DataColumn(title));
                }

                int count = tests.Count / 10;//10筆一輪，有幾輪
                for (int i = 0; i < count; i++)
                {
                    DataRow space = dt.NewRow();//換行區別群組用
                    DataRow dr = dt.NewRow();//顯示測資用
                    DataRow drDraw;//顯示畫圖用
                    DataRow drawspace = dtDraw.NewRow();//換行區別群組用

                    List<Test> temps = new List<Test>();

                    for (int j = 0; j < 10; j++)
                    {
                        temps.Add(new Test
                        {
                            Dend = tests[j + 10 * i].Dend,
                            Ltp = tests[j + 10 * i].Ltp,
                            Pow_eff = tests[j + 10 * i].Pow_eff
                        });

                        dr["Dend"] = tests[j + 10 * i].Dend;
                        dr["Ltp"] = tests[j + 10 * i].Ltp;
                        dr["Pow_eff"] = tests[j + 10 * i].Pow_eff;
                        dt.Rows.Add(dr.ItemArray);
                    }
                    var groupAns = temps.OrderBy(c => int.Parse(c.Dend)).ThenBy(c => int.Parse(c.Ltp)).ToList();

                    for (int j = 0; j < 10; j++)
                    {
                        drDraw = dtDraw.NewRow();
                        drDraw[0] = groupAns[j].Ltp;
                        drDraw[groupAns[j].Dend] = groupAns[j].Pow_eff;
                        dtDraw.Rows.Add(drDraw.ItemArray);
                    }

                    dt.Rows.Add(space.ItemArray);
                    dtDraw.Rows.Add(drawspace.ItemArray);
                }

                ans.DataSource = dt;//放到girdview
                draw.DataSource = dtDraw;
            }
        }

        private void export_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel 工作表|*.xlsx|所有檔案|*.*";//設定檔案型別
            sfd.FileName = "";//設定預設檔名
            sfd.DefaultExt = "xlsx";//設定預設格式（可以不設）
            sfd.AddExtension = true;//設定自動在檔名中新增副檔名
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                Microsoft.Office.Interop.Excel.Application xls = null;

                xls = new Microsoft.Office.Interop.Excel.Application();
                // Excel WorkBook
                Microsoft.Office.Interop.Excel.Workbook book = xls.Workbooks.Add();
                // Excel WorkBook，預設會產生一個 WorkSheet，索引從 1 開始，而非 0
                //// 寫法1
                //Excel.Worksheet Sheet = (Excel.Worksheet)book.Worksheets.Item[1];
                //// 寫法2
                //Excel.Worksheet Sheet = (Excel.Worksheet)book.Worksheets[1];
                // 寫法3
                Microsoft.Office.Interop.Excel.Worksheet Sheet = xls.ActiveSheet;

                // 把 DataGridView 資料塞進 Excel 內
                DataGridView2Excel(Sheet,draw);
                // 儲存檔案
                book.SaveAs(sfd.FileName);
            }
        }

        private void DataGridView2Excel(Microsoft.Office.Interop.Excel.Worksheet worksheet,DataGridView myDGV)
        {
            //寫入標題
            for (int i = 0; i < myDGV.ColumnCount; i++)
            {
                worksheet.Cells[1, i + 1] = myDGV.Columns[i].HeaderText;
            }
            //寫入數值
            for (int r = 0; r < myDGV.Rows.Count; r++)
            {
                for (int i = 0; i < myDGV.ColumnCount; i++)
                {
                    worksheet.Cells[r + 2, i + 1] = myDGV.Rows[r].Cells[i].Value;
                }
                System.Windows.Forms.Application.DoEvents();
            }
            worksheet.Columns.EntireColumn.AutoFit();//列寬自適應
        }
    }
}
