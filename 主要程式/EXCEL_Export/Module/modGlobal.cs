using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Drawing;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Configuration;

namespace RealGood.Module
{
    public class modGlobal
    {

        #region Func.: 設定DataGridView顯示格式 Layouts(....)
        public static void Layouts(DataGridView dgSelect, Color BackgroundColor, Color RowsBackColor, Color AlternatebackColor, Boolean AutoGenerateColumns, Color HeaderColor, Boolean HeaderVisual, Boolean RowHeadersVisible, Boolean AllowUserToAddRows, Color HeaderForeColor, int headerHeight)
        {
            //設定欄位寬度調整=自動延展整個顯示區域
            dgSelect.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            #region 欄位寬度調整說明
            //DataGridViewAutoSizeColumnsMode 定義值來指定要如何調整資料行的寬度。 
            //成員名稱                       說明 
            //AllCells                       資料行寬度調整以適合資料行中的所有儲存格的內容 (包括標題儲存格)。   
            //AllCellsExceptHeader           資料行寬度調整以適合資料行中的所有儲存格的內容 (不包括標題儲存格)。   
            //ColumnHeader                   資料行寬度調整以適合資料行行首儲存格的內容。   
            //DisplayedCells                 資料行寬度調整以適合資料行中的所有儲存格的內容 (位在目前顯示在螢幕上的資料列中)，包括標題儲存格。   
            //DisplayedCellsExceptHeader     資料行寬度調整以適合資料行中的所有儲存格的內容 (位在目前顯示在螢幕上的資料列中)，不包括標題儲存格。   
            //Fill                           資料行寬度調整使得所有資料行的寬度可以剛好填滿控制項的顯示區，且必須要使用水平捲動方式，才能讓資料行寬度維持在  DataGridViewColumn.MinimumWidth 屬性值之上。相對的資料行寬度是由相對的  DataGridViewColumn.FillWeight 屬性值所決定。  
            //None                           資料行寬度不會自動調整。   
            #endregion

            //Grid Background Color 
            dgSelect.BackgroundColor = BackgroundColor;

            //Grid Back Color 
            dgSelect.RowsDefaultCellStyle.BackColor = RowsBackColor;

            //GridColumnStylesCollection Alternate Rows Backcolr
            dgSelect.AlternatingRowsDefaultCellStyle.BackColor = AlternatebackColor;

            //Auto generated here set to true or false.
            dgSelect.AutoGenerateColumns = AutoGenerateColumns;
            //dgSelect.DefaultCellStyle.Font = new Font("Verdana", 10.25f, FontStyle.Regular);
            //dgSelect.ColumnHeadersDefaultCellStyle.Font = new Font("Calibri", 11, FontStyle.Regular);

            //Column Header Back Color
            dgSelect.ColumnHeadersDefaultCellStyle.BackColor = HeaderColor;
            dgSelect.ColumnHeadersDefaultCellStyle.ForeColor = HeaderForeColor;
            dgSelect.ColumnHeadersHeight = headerHeight;
            //Header Visisble
            dgSelect.EnableHeadersVisualStyles = HeaderVisual;

            //Row Header Back Color
            dgSelect.RowHeadersDefaultCellStyle.BackColor = HeaderColor;

            //Row Header Close
            //dgSelect.RowHeadersVisible = RowHeadersVisible;

            //To Hide the Last Empty row here we use false.
            dgSelect.AllowUserToAddRows = AllowUserToAddRows;

            //設定選取時的顏色
            dgSelect.DefaultCellStyle.SelectionBackColor = Color.Green;

            //設定資料行,整合標示
            dgSelect.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }
        #endregion

        #region Func.: 透過ODBC Conn取得DataSet GetDataSetFromODBC_Adapter(DataSet dsReturn, string sCONN, string sQueryCMD)
        public static DataSet GetDataSetFromODBC_Adapter(DataSet dsReturn, string sCONN, string sQueryCMD)
        {
            using (OdbcConnection connection = new OdbcConnection(sCONN))
            {
                try //資料庫連線
                {
                    connection.Open();
                }
                catch (Exception e1)
                {
                    MessageBox.Show("Oracle 連線失敗:" + e1.Message);
                }
                OdbcCommand command = new OdbcCommand();
                command.Connection = connection;
                command.CommandText = sQueryCMD;

                OdbcDataAdapter adapter = new OdbcDataAdapter();
                adapter.SelectCommand = command;

                dsReturn.Tables.Clear();    //清除DataSet
                dsReturn.AcceptChanges();   //DataSet異動確認
                try //資料庫綁定
                {
                    adapter.Fill(dsReturn);
                }
                catch (Exception e2)
                {
                    MessageBox.Show("Oracle 載入失敗:" + e2.Message);
                }
                // The connection is automatically closed when the
                // code exits the using block.
            }
            return dsReturn;
        }
        #endregion

        #region Func.: 透過SQL Conn取得DataSet GetDataSetFromSQL_Adapter(DataSet _dsReturn, string sCONN, string sQueryCMD)
        public static DataSet GetDataSetFromSQL_Adapter(DataSet _dsReturn, string sCONN, string sQueryCMD)
        {
            using (SqlConnection connection = new SqlConnection(sCONN))
            {
                try //資料庫連線
                {
                    connection.Open();
                }
                catch (Exception e1)
                {
                    MessageBox.Show("資料庫連線失敗:" + e1.Message);
                }
                SqlCommand command = new SqlCommand();
                command.Connection = connection;
                command.CommandText = sQueryCMD;

                SqlDataAdapter adapter = new SqlDataAdapter();
                adapter.SelectCommand = command;

                _dsReturn.Tables.Clear();    //清除DataSet
                _dsReturn.AcceptChanges();   //DataSet異動確認
                try //資料庫綁定
                {
                    adapter.Fill(_dsReturn);
                }
                catch (Exception e2)
                {
                    MessageBox.Show("資料庫載入失敗:" + e2.Message);
                }
                // The connection is automatically closed when the
                // code exits the using block.
            }
            return _dsReturn;
        }
        #endregion

        #region Func.: 透過SQL Conn取得DataTable GetDataTableFromSQL_Adapter(string sCONN, string sQueryCMD, string sDataMember)
        public static DataTable GetDataTableFromSQL_Adapter(string sCONN, string sQueryCMD, string sDataMember)
        {
            DataSet dsData = new DataSet();
            DataTable dtReturn = new DataTable();

            using (SqlConnection connection = new SqlConnection(sCONN))
            {
                try //資料庫連線
                {
                    connection.Open();
                }
                catch (Exception e1)
                {
                    MessageBox.Show("資料庫連線失敗:" + e1.Message);
                }
                SqlCommand command = new SqlCommand();
                command.Connection = connection;
                command.CommandText = sQueryCMD;

                SqlDataAdapter adapter = new SqlDataAdapter();
                adapter.SelectCommand = command;

                dsData.Tables.Clear();    //清除DataSet
                dsData.AcceptChanges();   //DataSet異動確認
                if (dsData.Tables[sDataMember] != null) //清除DataTable's DataMember
                {
                    DataTable _dt = dsData.Tables[sDataMember];
                    _dt.Rows.Clear();
                }
                try //資料庫綁定
                {
                    adapter.Fill(dsData, sDataMember);
                    dtReturn = dsData.Tables[sDataMember];
                }
                catch (Exception e2)
                {
                    MessageBox.Show("資料庫載入失敗:" + e2.Message);
                }
                // The connection is automatically closed when the
                // code exits the using block.
            }
            return dtReturn;
        }
        #endregion

        #region Func.: 取得目前選取的資料列 GetCurrentDataRow(BindingSource _bs)
        public static DataRow GetCurrentDataRow(BindingSource _bs)
        {
            if (_bs.Current == null)
                return null;
            else
                return ((DataRowView)_bs.Current).Row;
        }
        #endregion

        #region Func.: 字串自訂累加跳號 STR_INC(string sSN, string sFormat)
        public static string STR_INC(string sSN, string sFormat)
        {
            string sReturn = "";

            for (int i = sSN.Length - 1; i >= 0; i--)
            {
                string sGet = sSN.Substring(i, 1);
                int iPos = sFormat.IndexOf(sGet, 0);
                if (sSN.Substring(i, 1) == sFormat.Substring(sFormat.Length - 1, 1))    //判斷是否與跳號格式的最後碼相同,如相同代表需要進位
                {
                    sReturn = sFormat.Substring(0, 1) + sReturn;                        //取跳號格式的第一碼
                }
                else
                {
                    sReturn = sFormat.Substring(iPos + 1, 1) + sReturn;                 //取跳號格式的下一碼
                    sReturn = sSN.Substring(0, sSN.Length - sReturn.Length) + sReturn;  //組合剩餘的前置字串
                    break;
                }
            }
            return sReturn;
        }
        #endregion

        #region Func.: 十六進制轉十進制(H/L) Hex_HL_Dec(string sHexString)
        public static string Hex_HL_Dec(string sHex)  
        {
            sHex = sHex.Substring(4, 4) + sHex.Substring(0, 4);               //High Byte 與 Low Bye 互換
            int iNum = Int32.Parse(sHex, System.Globalization.NumberStyles.HexNumber);    //將16進位字串轉成10進制數值
            return iNum.ToString();
        }
        #endregion

        #region Func.: 字串轉ASCII String2ASCII(string sBuf)
        public static string String2ASCII(string sBuf)
        {
            string sASCII = "";
            int iValue = 0;
            try
            {
                foreach (char c in sBuf)
                {
                    iValue = System.Convert.ToInt32(c); //字元轉ASCII Code
                    sASCII += String.Format("{0:X}", iValue);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("字串轉換錯誤，錯誤訊息:\n" + ex.ToString(), "管理系統", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return sASCII;
        }
        #endregion

        #region Func.: 資料表匯出Excel報表 DataTable2Excel(string sTitle, DataTable _dtData, DataGridView _dgvData)
        private void DataTable2Excel(string sTitle, DataTable _dtData, DataGridView _dgvData)
        {
            //建立Excel檔案
            SaveFileDialog _sfdExcel = new SaveFileDialog();
            _sfdExcel.FileName = DateTime.Now.ToString("yyyy-MM-dd") + "-" + sTitle; //建立檔案名稱
            _sfdExcel.Filter = "Excel檔(*.xls) |*.xls ";    //存檔類型(格式)

            if (_sfdExcel.ShowDialog() == DialogResult.OK)
            {
                string sFilePath = _sfdExcel.FileName;      //取得檔案儲存路徑
                System.IO.StreamWriter _sw = new System.IO.StreamWriter(sFilePath, false, System.Text.Encoding.Unicode);
                _sw.WriteLine("<table border=1 cellspacing=0 cellpadding=0>");

                //設定EXCEL表格主標題
                _sw.WriteLine("<tr><td colspan=" + _dtData.Columns.Count.ToString() + " align=center>" + sTitle + "表格--" + "報表輸出日期:" + DateTime.Now.ToShortDateString() + "</td></tr>");

                #region 設定EXCEL表格欄位標題
                string sTagCommand = "<tr>";
                for (int i = 0; i < _dtData.Columns.Count; i++)
                {
                    sTagCommand += "<td align=center bgcolor=#48d1cc>" + _dgvData.Columns[i].HeaderText + "</td>";
                }
                sTagCommand += "</tr>";
                _sw.WriteLine(sTagCommand);
                #endregion

                #region 產生EXCEL表格資料內容
                sTagCommand = "";
                for (int i = 0; i < _dtData.Rows.Count; i++)
                {
                    sTagCommand += "<tr>";
                    for (int j = 0; j < _dtData.Columns.Count; j++)
                    {
                        sTagCommand += "<td align=center>" + _dtData.Rows[i][j].ToString() + "</td>";
                    }
                    sTagCommand += "</tr>";
                }
                _sw.WriteLine(sTagCommand);
                #endregion

                _sw.WriteLine("</table>");
                _sw.Close();

                MessageBox.Show(sTitle + "表格匯出完成!!");
            }
        }
        #endregion
    }
}
