using RealGood.Module;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Spire.Xls;

namespace EXCEL_Export
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            GC.Collect();
        }

        #region Func-匯出:查詢時間範圍內全部資料_BIND_RG_LOG_ALL()
        private DataTable _BIND_RG_LOG_ALL()
        {
            //取得生產資訊明細 DataTable
            String _sqlconn = ConfigurationManager.AppSettings["SQL_CONN"].ToString();
            String sQuery = @"SELECT * FROM LOG_DATA " +
                "WHERE LOG_MakeTime BETWEEN '" + dtpAfter.Value.ToString("yyyyMMdd") + "'" +
                " AND '" + DateTime.Parse(dtpBefore.Value.ToString()).AddDays(1).ToString("yyyyMMdd") + "'" +
                " ORDER BY LOG_MakeTime ASC";
            DataTable dtMRK_DATA = modGlobal.GetDataTableFromSQL_Adapter(_sqlconn, sQuery, "LOG_DATA");
            return dtMRK_DATA;

        }
        #endregion

        #region Button Func-匯出資料成EXCEL檔
        private void btnExport_Click(object sender, EventArgs e)
        {
            DataTable _dtLog = _BIND_RG_LOG_ALL();
            if(_dtLog.Rows.Count == 0)
            {
                MessageBox.Show("此時間範圍查無紀錄!");
                return;
            }
            else
            {
                SaveFileDialog sfdExport = new SaveFileDialog();
                sfdExport.Filter = "|*.xlsx";
                sfdExport.Title = "記錄存檔";
                DialogResult result = sfdExport.ShowDialog();
                if (result == DialogResult.OK)
                {
                    if(TableToExcel(_dtLog, sfdExport.FileName))
                    {
                        MessageBox.Show("匯出完成!");
                    }
                }
            }            
        }
        #endregion

        #region Datable匯出成Excel
        /// <summary>
        /// Datable匯出成Excel
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="file"></param>
        private bool TableToExcel(DataTable dt, string file)
        {
            IWorkbook workbook;
            string sFileExt = Path.GetExtension(file).ToLower();
            if (sFileExt == ".xlsx")
            {
                workbook = new XSSFWorkbook();
            }
            else if (sFileExt == ".xls")
            {
                workbook = new HSSFWorkbook();
            }
            else
            {
                workbook = null;
                MessageBox.Show("請點選EXCEL檔!");
                return false;
            }

            try
            {
                string sDate = DateTime.Parse(dt.Rows[0]["LOG_MakeTime"].ToString()).ToString("yyyy.MM.dd");
                ISheet sheet = string.IsNullOrEmpty(dt.TableName) ? workbook.CreateSheet("Sheet1") : workbook.CreateSheet(dt.TableName + "_" + sDate);

                //表頭  
                IRow row = sheet.CreateRow(0);
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    ICell cell = row.CreateCell(i);
                    cell.SetCellValue(dt.Columns[i].ColumnName);
                }

                //資料  
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    IRow row1 = sheet.CreateRow(i + 1);
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        ICell cell = row1.CreateCell(j);
                        cell.SetCellValue(dt.Rows[i][j].ToString());
                    }
                }               

                //轉為位元組陣列  
                MemoryStream stream = new MemoryStream();
                workbook.Write(stream);
                byte[] buf = stream.ToArray();

                //儲存為Excel檔案  
                using (FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write))
                {
                    fs.Write(buf, 0, buf.Length);
                    fs.Flush();
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("匯出處理失敗:" + ex.Message);
                return false;
            }
            return true;
        }
        #endregion

        #region Button Func-選擇合併檔1
        private void btnSelect1_Click(object sender, EventArgs e)
        {
            dlg1.Title = "開啟xlsx合併檔";
            dlg1.Filter = "|*.xlsx";
            dlg1.FilterIndex = 1;
            dlg1.RestoreDirectory = true;
            dlg1.Multiselect = false;
            if (dlg1.ShowDialog() == DialogResult.OK)
            {
                txtPath1.Text = dlg1.FileName;
                dlg1.Dispose();
            }
        }
        #endregion

        #region Button Func-選擇合併檔2
        private void btnSelect2_Click(object sender, EventArgs e)
        {
            dlg2.Title = "開啟xlsx合併檔";
            dlg2.Filter = "|*.xlsx";
            dlg2.FilterIndex = 1;
            dlg2.RestoreDirectory = true;
            dlg2.Multiselect = false;
            if (dlg2.ShowDialog() == DialogResult.OK)
            {
                txtPath2.Text = dlg2.FileName;
                dlg2.Dispose();
            }
        }
        #endregion

        #region Button Func-多個工作簿合併
        private void btnMerge_Click(object sender, EventArgs e)
        {
            if(txtPath1.Text == "" || txtPath2.Text == "")
            {
                MessageBox.Show("請選擇兩個合併檔!");
            }

            //創建一個Workbook類對象
            Workbook newbook = new Workbook();
            newbook.Version = ExcelVersion.Version2013;
            //刪除文檔中的工作表（新創建的文檔默認包含3張工作表）
            newbook.Worksheets.Clear();

            //創建一個臨時的workbook，用於加載需要合併的Excel文檔
            Workbook tempbook = new Workbook();

            //加載需要合併的Excel文檔
            string[] excelFiles = new string[] { txtPath1.Text.Trim() , txtPath2.Text.Trim() };

            //遍歷數組
            for (int i = 0; i < excelFiles.Length; i++)
            {
                //載入Excel文檔
                tempbook.LoadFromFile(excelFiles[i]);
                //調用AddCopy方法，將文檔中所有工作表添加到新的workbook
                foreach (Worksheet sheet in tempbook.Worksheets)
                {
                    newbook.Worksheets.AddCopy(sheet);
                }
            }

            //保存文檔
            newbook.SaveToFile(@"C:\Users\harry\OneDrive\桌面\result.xlsx", ExcelVersion.Version2013);

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            ExcelPackage epA = new ExcelPackage(new System.IO.FileInfo(@"C:\Users\harry\OneDrive\桌面\result.xlsx"));
            for (int i = 0; i < epA.Workbook.Worksheets.Count; i++)
            {
                string workSheetName = epA.Workbook.Worksheets[i].Name;
                if (workSheetName.Contains("Evaluation Warning"))
                {
                    epA.Workbook.Worksheets.Delete(i);
                }
            }
            //保存文檔
            epA.SaveAs(new FileInfo(@"C:\Users\harry\OneDrive\桌面\result.xlsx"));
            txtPath1.Text = "";
            txtPath2.Text = "";
            MessageBox.Show("合併完成!");

            //ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            //ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            //try
            //{
            //    var resultFile = @"C:\Users\harry\OneDrive\桌面\result.xlsx";

            //    using (ExcelPackage epA = new ExcelPackage(new System.IO.FileInfo(txtPath1.Text.Trim())))
            //    {
            //        using (ExcelPackage epB = new ExcelPackage(new System.IO.FileInfo(txtPath2.Text.Trim())))
            //        {
            //            foreach (var sheetB in epB.Workbook.Worksheets)
            //            {
            //                //check name of worksheet, in case that worksheet with same name already exist exception will be thrown by EPPlus
            //                string workSheetName = sheetB.Name;
            //                foreach (var sheetA in epA.Workbook.Worksheets)
            //                {
            //                    if (sheetB.Name == sheetA.Name)
            //                    {
            //                        //如果檢測到相同名稱資料表，直接跳過
            //                        continue;
            //                    }
            //                }

            //                //add new sheet
            //                epA.Workbook.Worksheets.Add(workSheetName, sheetB);
            //            }
            //        }
            //        epA.SaveAs(new FileInfo(resultFile));
            //    }

            //    txtPath1.Text = "";
            //    txtPath2.Text = "";
            //    MessageBox.Show("合併完成!");
            //}
            //catch(Exception ex)
            //{
            //    MessageBox.Show("合併失敗:" + ex.Message);
            //}
        }
        #endregion

    }
}
