using System;
using System.Collections.Generic;
using System.Data;
using System.Windows;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using Microsoft.Win32;
using System.Text;
using System.Drawing;
using System.Drawing.Imaging;
using ZXing;
using System.IO;
using ZXing.Common;

namespace BusinessCardQR
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 
    public partial class MainWindow : Window
    {
        #region GlobalVariables
        string excelFilePath = string.Empty;
        DataTable dtExcel = null;
        DataTable dtFiltered = null;
        StringBuilder sbQRContent = null;
        //Pager
        int pageSize = 0;
        int currentPage = 1;
        int lastPage = 0;
        //List for combo box
        List<int> lstItemsPerPage = new List<int>();
        #endregion

        #region C'tor
        public MainWindow()
        {
            InitializeComponent();
            SetContolsVisibility(false);

            //DataView dv = GetEmpyView();
            //dgEmployees.ItemsSource = dv;
        }
        #endregion

        #region Methods
        private string GetUploadedExcelPath()
        {
            string path = string.Empty;
            OpenFileDialog fdlg = new OpenFileDialog();

            fdlg.Title = "Select excel template";
            fdlg.Filter = "Excel files|*.xls;*.xlsx;*.xlsm";
            fdlg.FilterIndex = 3;
            fdlg.RestoreDirectory = true;

            if (fdlg.ShowDialog() == true)
            {
                path = fdlg.FileName;
            }
            else
            {
                return string.Empty;
            }

            return path;
        }
        private DataTable GetDataTableFromUploadedExcel(string path)
        {
            using (DataTable dt = ReadExcelFileDOM(path))
            {
                if (dt.Rows.Count >= 1 && dt.Columns.Count == 11)
                {
                    dt.Rows[0].Delete();

                    dt.Columns[0].ColumnName = "FirstName";
                    dt.Columns[1].ColumnName = "MiddleName";
                    dt.Columns[2].ColumnName = "LastName";
                    dt.Columns[3].ColumnName = "Organization";
                    dt.Columns[4].ColumnName = "Title";
                    dt.Columns[5].ColumnName = "Mobile";
                    dt.Columns[6].ColumnName = "Landline";
                    dt.Columns[7].ColumnName = "Fax";
                    dt.Columns[8].ColumnName = "URL";
                    dt.Columns[9].ColumnName = "Email";
                    dt.Columns[10].ColumnName = "Address";

                    dt.AcceptChanges();
                }
                else
                {
                    //If invalid excel is uploaded, remove all columns from data table
                    dt.Clear();
                }

                return dt;
            }
        }
        private void GenerateQRCode(object sender, RoutedEventArgs e)
        {
            if (!Equals(dtExcel, null) && dtExcel.Rows.Count > 0)
            {
                var row = ((DataRowView)dgEmployees.SelectedItem).Row.ItemArray;
                Console.WriteLine("Data Row");

                string fName = row[0].ToString();
                string mName = row[1].ToString();
                string lName = row[2].ToString();
                string org = row[3].ToString();
                string title = row[4].ToString();
                string mobile = row[5].ToString();
                string landline = row[6].ToString();
                string fax = row[7].ToString();
                string url = row[8].ToString();
                string email = row[9].ToString();
                string address = row[10].ToString();

                SaveFileDialog sv = new SaveFileDialog() { Filter = "PNG|.png", ValidateNames = true };

                sv.FileName = fName + lName;

                sbQRContent = null;
                sbQRContent = new StringBuilder();
                sbQRContent.Append("BEGIN:VCARD\r\n");
                sbQRContent.Append("VERSION:3.0\r\n");
                if (string.IsNullOrEmpty(mName) || string.IsNullOrWhiteSpace(mName))
                {
                    sbQRContent.AppendFormat("N:{0} {1} \r\n", fName, lName);
                }
                else
                {
                    sbQRContent.AppendFormat("N:{0} {1} {2}\r\n", fName, mName, lName);
                }
                sbQRContent.AppendFormat("ORG:{0}\r\n", org);
                sbQRContent.AppendFormat("TITLE:{0}\r\n", title);
                sbQRContent.AppendFormat("TEL:{0}\r\n", mobile);
                sbQRContent.AppendFormat("TEL;type=WORK:{0}\r\n", landline);
                sbQRContent.AppendFormat("TEL;type=FAX:{0}\r\n", fax);
                sbQRContent.AppendFormat("URL:{0}\r\n", url);
                sbQRContent.AppendFormat("EMAIL:{0}\r\n", email);
                sbQRContent.AppendFormat("ADR:{0}\r\n", address);
                sbQRContent.Append("END:VCARD\r\n");

                if (sv.ShowDialog() == true)
                {
                    using (Bitmap bmp = GenerateQRCodeImage(sbQRContent.ToString()))
                    {
                        bmp.Save(sv.FileName, ImageFormat.Png);
                    }
                }
            }
            else
            {
                MessageBox.Show("Please upload Employee Data excel sheet with valid data.");
            }
        }
        public static Bitmap GenerateQRCodeImage(string content, string alt = "QR code", int height = 300, int width = 300, int margin = 0)
        {
            var qrWriter = new BarcodeWriter()
            {
                Format = BarcodeFormat.QR_CODE,
                Options = new EncodingOptions() { Height = height, Width = width, Margin = margin }
            };

            Bitmap q = qrWriter.Write(content);
            q.Save(new MemoryStream(), ImageFormat.Png);
            return q;
        }
        private DataTable GetDataTable()
        {
            if (dtFiltered != null && dtFiltered.Rows.Count > 0)
            {
                return dtFiltered;
            }
            else
            {
                return dtExcel;
            }
        }
        private DataView ShowData(int pageNumber, DataTable dataSource)
        {
            if (string.IsNullOrEmpty(txtSearch.Text.Trim()) || string.IsNullOrEmpty(txtSearch.Text.Trim()))
            {
                dataSource = dtExcel;
            }
            else
            {
                dataSource = dtFiltered;
            }

            if (cmbRecordCount.Items.Count == 0)
            {
                BindItemsPerPageDropDown();
            }

            DataTable dt = new DataTable();

            pageSize = GetPageSize();

            int startIndex = pageSize * (pageNumber - 1);
            var result = dataSource.AsEnumerable().Where((s, k) => (k >= startIndex && k < (startIndex + pageSize)));

            foreach (DataColumn colunm in dataSource.Columns)
            {
                dt.Columns.Add(colunm.ColumnName);
            }

            foreach (var item in result)
            {
                dt.ImportRow(item);
            }

            dgEmployees.ItemsSource = dt.DefaultView;

            int rowCount = dataSource.Rows.Count;

            if (rowCount == 0)
            {
                lblPager.Text = "No records found!";
            }
            else if (rowCount >= 1 && rowCount < pageSize)
            {
                lblPager.Text = string.Format("Page {0} Of {1}", pageNumber, 1);
            }
            else
            {
                if ((rowCount % pageSize) == 0)
                {
                    lblPager.Text = string.Format("Page {0} Of {1}", pageNumber, (rowCount / pageSize));
                }
                else
                {
                    lblPager.Text = string.Format("Page {0} Of {1}", pageNumber, (rowCount / pageSize) + 1);
                }
            }

            //Adjust Last Page
            if (rowCount > 0 && rowCount < pageSize)
            {
                lastPage = 1;
            }
            else
            {
                if (rowCount % pageSize == 0)
                {
                    lastPage = (rowCount / pageSize);
                }
                else
                {
                    lastPage = (rowCount / pageSize) + 1;
                }
            }

            EnableDisablePagerButtons();

            return dt.DefaultView;
        }
        private void BindItemsPerPageDropDown()
        {
            lstItemsPerPage = ItemsPerPageList();

            cmbRecordCount.ItemsSource = lstItemsPerPage;

            //cmbRecordCount.SelectedIndex = 0;
            cmbRecordCount.SelectedIndex = lstItemsPerPage.Count - 1;
        }
        private List<int> ItemsPerPageList()
        {
            const int lower_value = 5, step = 5;

            int upper_value = dtExcel.Rows.Count;

            int numPages = (upper_value % 5) == 0 ? upper_value : (upper_value / step) + 1;

            return Enumerable.Range(lower_value, upper_value).Where(x => x % step == 0).ToList();
        }
        private int GetPageSize()
        {
            return int.Parse(cmbRecordCount.SelectedValue.ToString());
        }
        private void SetContolsVisibility(bool state)
        {
            if (state)
            {
                //imgStartScreen.Visibility = Visibility.Hidden;
                dgEmployees.Visibility = Visibility.Visible;
                txtSearch.Visibility = Visibility.Visible;
                lblRecordToShow.Visibility = Visibility.Visible;
                cmbRecordCount.Visibility = Visibility.Visible;
                btnPrevious.Visibility = Visibility.Visible;
                btnNext.Visibility = Visibility.Visible;
                btnFirst.Visibility = Visibility.Visible;
                btnLast.Visibility = Visibility.Visible;
                btnClearSort.Visibility = Visibility.Visible;
                lblPager.Visibility = Visibility.Visible;
                lblSearch.Visibility = Visibility.Visible;
            }
            else
            {
                //imgStartScreen.Visibility = Visibility.Visible;
                dgEmployees.Visibility = Visibility.Hidden;
                txtSearch.Visibility = Visibility.Hidden;
                cmbRecordCount.Visibility = Visibility.Hidden;
                lblRecordToShow.Visibility = Visibility.Hidden;
                btnPrevious.Visibility = Visibility.Hidden;
                btnNext.Visibility = Visibility.Hidden;
                btnFirst.Visibility = Visibility.Hidden;
                btnLast.Visibility = Visibility.Hidden;
                btnClearSort.Visibility = Visibility.Hidden;
                lblPager.Visibility = Visibility.Hidden;
                lblSearch.Visibility = Visibility.Hidden;
            }
        }
        private void EnableDisablePagerButtons()
        {
            if (currentPage == 1)
            {
                btnFirst.IsEnabled = false;
                btnPrevious.IsEnabled = false;
            }
            else if (currentPage > 1)
            {
                btnFirst.IsEnabled = true;
                btnPrevious.IsEnabled = true;
            }
            else
            {
                btnFirst.IsEnabled = false;
                btnPrevious.IsEnabled = false;
            }

            if (currentPage == lastPage)
            {
                btnNext.IsEnabled = false;
                btnLast.IsEnabled = false;
            }
            else
            {
                btnNext.IsEnabled = true;
                btnLast.IsEnabled = true;
            }
        }
        private DataView GetEmpyView()
        {
            DataTable dt = new DataTable();

            DataColumn col1 = new DataColumn("FirstName", typeof(String));
            DataColumn col2 = new DataColumn("MiddleName", typeof(DateTime));
            DataColumn col3 = new DataColumn("LastName", typeof(String));
            DataColumn col4 = new DataColumn("Organization", typeof(DateTime));
            DataColumn col5 = new DataColumn("Title", typeof(String));
            DataColumn col6 = new DataColumn("Mobile", typeof(DateTime));
            DataColumn col7 = new DataColumn("Landline", typeof(String));
            DataColumn col8 = new DataColumn("Fax", typeof(DateTime));
            DataColumn col9 = new DataColumn("Url", typeof(String));
            DataColumn col10 = new DataColumn("Email", typeof(DateTime));
            DataColumn col11 = new DataColumn("Address", typeof(String));

            dt.Columns.Add(col1);
            dt.Columns.Add(col2);
            dt.Columns.Add(col3);
            dt.Columns.Add(col4);
            dt.Columns.Add(col5);
            dt.Columns.Add(col6);
            dt.Columns.Add(col7);
            dt.Columns.Add(col8);
            dt.Columns.Add(col9);
            dt.Columns.Add(col10);

            return dt.DefaultView;
        }
        #region OpenXML
        private DataTable ReadExcelFileDOM(string filename)
        {
            DataTable table;

            using (SpreadsheetDocument myDoc = SpreadsheetDocument.Open(filename, true))
            {
                WorkbookPart workbookPart = myDoc.WorkbookPart;
                Sheet worksheet = workbookPart.Workbook.Descendants<Sheet>().First();
                WorksheetPart worksheetPart =
                 (WorksheetPart)(workbookPart.GetPartById(worksheet.Id));
                SheetData sheetData =
                    worksheetPart.Worksheet.Elements<SheetData>().First();
                List<List<string>> totalRows = new List<List<string>>();
                int maxCol = 0;

                foreach (Row r in sheetData.Elements<Row>())//Skip header row
                {
                    // Add the empty row.
                    string value = null;
                    while (totalRows.Count < r.RowIndex - 1)
                    {
                        List<string> emptyRowValues = new List<string>();
                        for (int i = 0; i < maxCol; i++)
                        {
                            emptyRowValues.Add("");
                        }
                        totalRows.Add(emptyRowValues);
                    }


                    List<string> tempRowValues = new List<string>();
                    foreach (Cell c in r.Elements<Cell>())
                    {
                        #region get the cell value of c.
                        if (c != null)
                        {
                            value = c.InnerText;

                            // If the cell represents a numeric value, you are done. 
                            // For dates, this code returns the serialized value that 
                            // represents the date. The code handles strings and Booleans
                            // individually. For shared strings, the code looks up the 
                            // corresponding value in the shared string table. For Booleans, 
                            // the code converts the value into the words TRUE or FALSE.
                            if (c.DataType != null)
                            {
                                switch (c.DataType.Value)
                                {
                                    case CellValues.SharedString:
                                        // For shared strings, look up the value in the shared 
                                        // strings table.
                                        var stringTable = workbookPart.
                                            GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                                        // If the shared string table is missing, something is 
                                        // wrong. Return the index that you found in the cell.
                                        // Otherwise, look up the correct text in the table.
                                        if (stringTable != null)
                                        {
                                            value = stringTable.SharedStringTable.
                                                ElementAt(int.Parse(value)).InnerText;
                                        }
                                        break;

                                    case CellValues.Boolean:
                                        switch (value)
                                        {
                                            case "0":
                                                value = "FALSE";
                                                break;
                                            default:
                                                value = "TRUE";
                                                break;
                                        }
                                        break;
                                }
                            }

                            Console.Write(value + "  ");
                        }
                        #endregion

                        // Add the cell to the row list.
                        int i = Convert.ToInt32(c.CellReference.ToString().ToCharArray().First() - 'A');

                        // Add the blank cell in the row.
                        while (tempRowValues.Count < i)
                        {
                            tempRowValues.Add("");
                        }
                        tempRowValues.Add(value);
                    }

                    // add the row to the totalRows.
                    maxCol = processList(tempRowValues, totalRows, maxCol);

                    Console.WriteLine();
                }

                table = ConvertListListStringToDataTable(totalRows, maxCol);
            }
            return table;
        }
        private int processList(List<string> tempRows, List<List<string>> totalRows, int MaxCol)
        {
            if (tempRows.Count > MaxCol)
            {
                MaxCol = tempRows.Count;
            }

            totalRows.Add(tempRows);
            return MaxCol;
        }
        private DataTable ConvertListListStringToDataTable(List<List<string>> totalRows, int maxCol)
        {
            DataTable table = new DataTable();
            for (int i = 0; i < maxCol; i++)
            {
                table.Columns.Add();
            }
            foreach (List<string> row in totalRows)
            {
                while (row.Count < maxCol)
                {
                    row.Add("");
                }
                table.Rows.Add(row.ToArray());
            }
            return table;
        }

        #endregion
        #endregion

        #region Events
        private void btnUploadExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string excelFilePath = GetUploadedExcelPath();

                if (!string.IsNullOrEmpty(excelFilePath))
                {
                    lstItemsPerPage.Clear();

                    dtExcel = GetDataTableFromUploadedExcel(excelFilePath);

                    if (!Equals(dtExcel, null) && dtExcel.Rows.Count > 0)
                    {
                        SetContolsVisibility(true);
                    }
                }

                if (!Equals(dtExcel, null) && dtExcel.Rows.Count > 0)
                {
                    dgEmployees.ItemsSource = ShowData(currentPage, dtExcel);
                }
                else
                {
                    if (excelFilePath.Length > 0)
                    {
                        MessageBox.Show("Please upload employee data excel sheet with valid data");
                    }
                    else
                    {
                        MessageBox.Show("Please upload employee data excel sheet.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(new StringBuilder().AppendFormat("Error occured while processing your request.\r\nError detail are : {0}", ex.Message)
                    .ToString(), "Error");
            }
        }
        private void btnClearSort_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = GetDataTable();
            dgEmployees.ItemsSource = ShowData(currentPage, dt);
        }
        private void txtSearch_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            currentPage = 1;

            EnumerableRowCollection<DataRow> filtered = null;

            if (!Equals(dtExcel, null))
            {
                filtered = dtExcel.AsEnumerable()
                            .Where(
                                r => r.Field<String>("FirstName").Contains(txtSearch.Text.Trim())
                                || r.Field<String>("MiddleName").Contains(txtSearch.Text.Trim())
                                || r.Field<String>("LastName").Contains(txtSearch.Text.Trim())
                                || r.Field<String>("Organization").Contains(txtSearch.Text.Trim())
                                || r.Field<String>("Title").Contains(txtSearch.Text.Trim())
                                || r.Field<String>("Mobile").Contains(txtSearch.Text.Trim())
                                || r.Field<String>("Landline").Contains(txtSearch.Text.Trim())
                                || r.Field<String>("Fax").Contains(txtSearch.Text.Trim())
                                || r.Field<String>("URL").Contains(txtSearch.Text.Trim())
                                || r.Field<String>("Email").Contains(txtSearch.Text.Trim())
                                || r.Field<String>("Address").Contains(txtSearch.Text.Trim())
                            );
            }

            if (!Equals(filtered, null) && filtered.Count() > 0)
            {
                dtFiltered = filtered.CopyToDataTable();

                ShowData(currentPage, dtFiltered);
            }
            else
            {
                if (!Equals(dtFiltered, null))
                {
                    dtFiltered.Clear();

                    ShowData(currentPage, dtFiltered);
                }
            }
        }
        private void btnFirst_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = GetDataTable();

            if (currentPage == 1)
            {
                MessageBox.Show("You are already on First Page.");
            }
            else
            {
                currentPage = 1;
                dgEmployees.ItemsSource = ShowData(currentPage, dt);
            }
        }
        private void btnPrevious_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = GetDataTable();

            if (currentPage == 1)
            {
                //btnPrevious.Enabled = false;
                MessageBox.Show("You are already on First page, you can not go to previous of First page.");
            }
            else
            {
                btnPrevious.IsEnabled = true;
                currentPage -= 1;
                dgEmployees.ItemsSource = ShowData(currentPage, dt);
            }
        }
        private void btnNext_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = GetDataTable();

            int rowCount = dt.Rows.Count;

            if ((rowCount % pageSize) == 0)
            {
                lastPage = (rowCount / pageSize);
            }
            else
            {
                lastPage = (rowCount / pageSize) + 1;
            }

            //int lastPage = (rowCount / pageSize) + 1;

            if (currentPage == lastPage)
            {
                MessageBox.Show("You are already on Last page, you can not go to next page of Last page.");
            }
            else
            {
                currentPage += 1;
                dgEmployees.ItemsSource = ShowData(currentPage, dt);
            }
        }
        private void btnLast_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = GetDataTable();

            int rowCount = dt.Rows.Count;

            int previousPage = currentPage;

            if ((rowCount % pageSize) == 0)
            {
                currentPage = (rowCount / pageSize);
            }
            else
            {
                currentPage = (rowCount / pageSize) + 1;
            }

            if (previousPage == currentPage)
            {
                MessageBox.Show("You are already on Last Page.");
            }
            else
            {
                dgEmployees.ItemsSource = ShowData(currentPage, dt);
            }
        }
        private void cmbRecordCount_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (cmbRecordCount.SelectedValue != null)
            {
                //Rebind grid
                currentPage = 1;
                ShowData(currentPage, dtExcel);
            }
        }
        private void Window_StateChanged(object sender, EventArgs e)
        {
            if (this.WindowState == WindowState.Normal)
            {
                this.WindowState = WindowState.Maximized;
            }
        }
        #endregion
    }
}
