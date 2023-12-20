using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;
using ExcelData = Microsoft.Office.Interop.Excel;

namespace MaterialAreaCode
{
    public partial class Form1 : Form
    {
        string sQuery = "";
        SqlCommon dc = new SqlCommon();


        DataTable dataCopy = new DataTable();
        DataTable dataCopy1 = new DataTable();
        DataTable dataCopy2 = new DataTable();

        public Form1()
        {
            InitializeComponent();
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            dataCopy = new DataTable();

            gridView1.Columns.Clear();
            gridControl1.DataSource = null;

            string connectionString = string.Empty;
            string sheetName = string.Empty;
            DataTable dtExecl = new DataTable();


            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "excel Files|*.xlsx;*.xls";
            ofd.Title = "Select a excel File";

            Application.DoEvents();

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                txtTab.Text = ofd.FileName;

                dataCopy = ReadData(txtTab.Text, new string[] { "Code", "Line" });

                if (dataCopy.Rows.Count != 0)
                {
                    gridControl1.DataSource = dataCopy;
                    gridControl1.Update();

                    this.gridView1.OptionsView.ColumnAutoWidth = false;

                    this.gridView1.Columns[0].Caption = "현품표";
                    this.gridView1.Columns[1].Caption = "라인";

                    this.gridView1.Columns[0].Width = 190;
                    this.gridView1.Columns[1].Width = 50;
                }
            }
        }

        public DataTable ReadData(string aFilePath, string[] aColumns, int aHeaderRow = 1, int aHeaderColumn = 1, int aNullableCheckMaxColumn = 1, int aNullableRowCount = 1, bool aCheckColumnName = false)
        {
            System.Data.DataTable dtExcel = new System.Data.DataTable();
            string path = aFilePath;
            string excelConnStr = string.Empty;
            ExcelData.Application ExcelObj = new ExcelData.Application();
            ExcelData.Workbook theWorkBook = ExcelObj.Workbooks.Open(aFilePath);
            try
            {
                ExcelData.Sheets sheets = theWorkBook.Worksheets;
                ExcelData.Worksheet worksheet = (ExcelData.Worksheet)sheets.get_Item(1);

                foreach (var column in aColumns)
                {
                    dtExcel.Columns.Add(column);
                }

                int currentNullRowCount = 0;
                for (int i = 1; i <= worksheet.Rows.Count; i++)
                {
                    ExcelData.Range range = worksheet.get_Range(GetExcelColumnNameUsingIndex(aHeaderColumn) + i.ToString(),
                                                                GetExcelColumnNameUsingIndex(aHeaderColumn + aColumns.Length - 1) + i.ToString());
                    System.Array myvalues = (System.Array)range.Cells.Value;
                    if (i <= aHeaderRow)
                    {
                        if (aCheckColumnName == true && i == aHeaderRow)
                        {
                            if (myvalues.GetValue(1, 1) == null ||
                                dtExcel.Columns[0].ColumnName != myvalues.GetValue(1, 1).ToString())
                            {
                                break;
                            }
                        }
                        continue;
                    }
                    if (myvalues.GetValue(1, aNullableCheckMaxColumn) == null)
                    {
                        currentNullRowCount++;
                        if (currentNullRowCount == aNullableRowCount)
                        {
                            break;
                        }
                    }
                    else
                    {
                        currentNullRowCount = 0;
                    }
                    dtExcel.Rows.Add(ConvertToDataRow(dtExcel.NewRow(), myvalues));

                }
            }
            finally
            {
                if (theWorkBook != null)
                {
                    theWorkBook.Close();
                }

                ExcelObj.Quit();
            }

            return dtExcel;
        }

        private string GetExcelColumnNameUsingIndex(int aIndex)
        {
            int asciiPrefix = 64;
            int convertValue = aIndex + asciiPrefix;
            string returnValue = string.Empty;

            if (convertValue > 90)
            {
                returnValue = "A";
                convertValue = asciiPrefix + (convertValue - 90);
            }

            returnValue += ((char)(convertValue)).ToString();

            return returnValue;
        }

        private DataRow ConvertToDataRow(DataRow dr_Temp, System.Array myValue)
        {
            int ColumnIndex = 0;
            foreach (object iTem in myValue)
            {
                dr_Temp[ColumnIndex] = iTem;
                ColumnIndex++;
            }
            return dr_Temp;
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            if (dataCopy.Rows.Count == 0)
            {
                MessageBox.Show("등록 데이터가 존재하지 않습니다.");
                return;
            }

            else
            {
                int cnt = 1;

                for (int i = 0; i < dataCopy.Rows.Count; i++)
                {
                    string sCode = dataCopy.Rows[i]["Code"].ToString();//
                    string sLine = dataCopy.Rows[i]["Line"].ToString();

                    sQuery = "";
                    sQuery += "  SELECT COUNT(*) FROM WMS.MaterialID WITH(NOLOCK) \n";
                    sQuery += "  WHERE MaterialID = '" + sCode.Trim() + "'               \n";
                    dc = new SqlCommon();
                    string sMaterialCnt = dc.getSimpleScalar(sQuery).ToString();

                    if (sMaterialCnt != "0")
                    {
                        sQuery = "";
                        sQuery += "  SELECT COUNT(*) FROM baLineMaster WITH(NOLOCK) \n";
                        sQuery += "  WHERE LineCode = '" + sLine + "'               \n";
                        dc = new SqlCommon();
                        string sLineCnt = dc.getSimpleScalar(sQuery).ToString();

                        if (sLineCnt != "0")
                        {
                            sQuery = "";
                            sQuery += "  UPDATE WMS.MaterialID SET                               \n";
                            sQuery += "  AreaCode = '" + sLine + "'                               \n";
                            sQuery += "  WHERE MaterialID = '" + sCode.Trim() + "'               \n";
                            dc = new SqlCommon();
                            int iReCnt = dc.execNonQuery(sQuery);

                            listBox1.Items.Add(cnt.ToString() + ". " + sCode + " 저장 완료");
                        }
                        else
                        {
                            listBox1.Items.Add(cnt.ToString() + ". " + sLine + " 라인 정보를 알 수 없음");
                        }
                    }
                    else
                    {
                        listBox1.Items.Add(cnt.ToString() + ". " + sCode + " 현품표 정보를 알 수 없음");
                    }

                    cnt++;
                }

                MessageBox.Show("처리가 완료되었습니다.");
            }
        }

        private void simpleButton3_Click(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            dataCopy1 = new DataTable();

            gridView2.Columns.Clear();
            gridControl2.DataSource = null;

            string connectionString = string.Empty;
            string sheetName = string.Empty;
            DataTable dtExecl = new DataTable();


            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "excel Files|*.xlsx;*.xls";
            ofd.Title = "Select a excel File";

            Application.DoEvents();

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                txtTab1.Text = ofd.FileName;

                dataCopy1 = ReadData(txtTab1.Text, new string[] { "Code", "Qty" });

                if (dataCopy1.Rows.Count != 0)
                {
                    gridControl2.DataSource = dataCopy1;
                    gridControl2.Update();

                    this.gridView2.OptionsView.ColumnAutoWidth = false;

                    this.gridView2.Columns[0].Caption = "현품표";
                    this.gridView2.Columns[1].Caption = "수량";

                    this.gridView2.Columns[0].Width = 190;
                    this.gridView2.Columns[1].Width = 50;
                }
            }
        }

        private void simpleButton5_Click(object sender, EventArgs e)
        {
            listBox3.Items.Clear();
            dataCopy2 = new DataTable();

            gridView3.Columns.Clear();
            gridControl3.DataSource = null;

            string connectionString = string.Empty;
            string sheetName = string.Empty;
            DataTable dtExecl = new DataTable();


            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "excel Files|*.xlsx;*.xls";
            ofd.Title = "Select a excel File";

            Application.DoEvents();

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                txtTab2.Text = ofd.FileName;

                dataCopy2 = ReadData(txtTab2.Text, new string[] { "Code", "LotNo" });

                if (dataCopy2.Rows.Count != 0)
                {
                    gridControl3.DataSource = dataCopy2;
                    gridControl3.Update();

                    this.gridView3.OptionsView.ColumnAutoWidth = false;

                    this.gridView3.Columns[0].Caption = "현품표";
                    this.gridView3.Columns[1].Caption = "LotNo";

                    this.gridView3.Columns[0].Width = 190;
                    this.gridView3.Columns[1].Width = 100;
                }
            }
        }

        private void simpleButton6_Click(object sender, EventArgs e)
        {
            if (dataCopy1.Rows.Count == 0)
            {
                MessageBox.Show("등록 데이터가 존재하지 않습니다.");
                return;
            }

            else
            {
                int cnt = 1;

                for (int i = 0; i < dataCopy1.Rows.Count; i++)
                {
                    string sCode = dataCopy1.Rows[i]["Code"].ToString();//
                    string sQty = dataCopy1.Rows[i]["Qty"].ToString();

                    sQuery = "";
                    sQuery += "  SELECT COUNT(*) FROM WMS.MaterialID WITH(NOLOCK) \n";
                    sQuery += "  WHERE MaterialID = '" + sCode.Trim() + "'               \n";
                    dc = new SqlCommon();
                    string sMaterialCnt = dc.getSimpleScalar(sQuery).ToString();

                    if (sMaterialCnt != "0")
                    {

                        sQuery = "";
                        sQuery += "  UPDATE WMS.MaterialID SET                               \n";
                        sQuery += "  StockQty = '" + sQty + "'                               \n";
                        sQuery += "  WHERE MaterialID = '" + sCode.Trim() + "'               \n";
                        dc = new SqlCommon();
                        int iReCnt = dc.execNonQuery(sQuery);

                        listBox2.Items.Add(cnt.ToString() + ". " + sCode + " 저장 완료");

                    }
                    else
                    {
                        listBox2.Items.Add(cnt.ToString() + ". " + sCode + " 현품표 정보를 알 수 없음");
                    }

                    cnt++;
                }

                MessageBox.Show("처리가 완료되었습니다.");
            }
        }

        private void simpleButton8_Click(object sender, EventArgs e)
        {
            if (dataCopy2.Rows.Count == 0)
            {
                MessageBox.Show("등록 데이터가 존재하지 않습니다.");
                return;
            }

            else
            {
                int cnt = 1;

                for (int i = 0; i < dataCopy2.Rows.Count; i++)
                {
                    string sCode = dataCopy2.Rows[i]["Code"].ToString();//
                   // string sQty = dataCopy2.Rows[i]["Qty"].ToString();

                    sQuery = "";
                    sQuery += "  SELECT COUNT(*) FROM WMS.MaterialID WITH(NOLOCK) \n";
                    sQuery += "  WHERE MaterialID = '" + sCode.Trim() + "'               \n";
                    dc = new SqlCommon();
                    string sMaterialCnt = dc.getSimpleScalar(sQuery).ToString();

                    if (sMaterialCnt != "0")
                    {

                        sQuery = "";
                        sQuery += "  UPDATE WMS.MaterialID SET                               \n";
                        sQuery += "  Status = 'C'                                            \n";
                        sQuery += "  WHERE MaterialID = '" + sCode.Trim() + "'               \n";
                        dc = new SqlCommon();
                        int iReCnt = dc.execNonQuery(sQuery);

                        listBox3.Items.Add(cnt.ToString() + ". " + sCode + " 저장 완료");

                    }
                    else
                    {
                        listBox3.Items.Add(cnt.ToString() + ". " + sCode + " 현품표 정보를 알 수 없음");
                    }

                    cnt++;
                }

                MessageBox.Show("처리가 완료되었습니다.");
            }
        }

        //맨마지막 줄 주석 추가2222
        //브랜치 추가함
        //마스터 수정 해봄 
    }
}
