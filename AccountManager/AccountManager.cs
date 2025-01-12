using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using NickBuhro.NumToWords;
using NickBuhro.NumToWords.Russian;


namespace AccountManager
{
    public partial class AccountManager : Form
    {
        private string headerFilePath;
        private string accnum;
        private string accdata;
        private string tableFilePath;
        private string priceListFilePath;
        private string contractNumbersFilePath;
        private string currentProductName;

        public AccountManager()
        {
            InitializeComponent();

            button7.Click += button7_Click;
            button1.Click += button1_Click;
            button8.Click += button8_Click;
            button2.Click += button2_Click;
            button9.Click += button9_Click;
            button3.Click += button3_Click;
            button5.Click += button5_Click;
            button6.Click += button6_Click;
            button10.Click += button10_Click;
            dataGridView1.CellMouseClick += dataGridView1_CellMouseClick;
        }



        private void UpdateAccountDetails()
        {
            accnum = textBox1.Text;
            accdata = textBox2.Text;
        }
        public void LoadExcelToDataGridView(string filePath, DataGridView targetGrid)
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                DataTable dataTable = new DataTable();

                // ��������� ��������� ��� ������� ������
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    dataTable.Columns.Add($"������� {col}"); // �������� ��������� �����
                }

                // ���������� ������, ������� ���������
                for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                {
                    DataRow dataRow = dataTable.NewRow();
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        dataRow[col - 1] = worksheet.Cells[row, col].Text;
                    }
                    dataTable.Rows.Add(dataRow);
                }

                targetGrid.DataSource = dataTable;

                // ��������� �������������� ������������� ������ ������ ��� ���������
                targetGrid.ColumnHeadersVisible = false;
                targetGrid.AllowUserToAddRows = false; // ������� ������ ������ �����
                targetGrid.RowHeadersVisible = false; // ������� ������� � ����������� �����
            }
        }

        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex < 3 || e.RowIndex == dataGridView1.Rows.Count - 1) // ���������� ������ 1-3 � ��������� ������
                return;

            int clickedRow = e.RowIndex;

            int orderCount = 0;

            // ��� ��������� �������
            var tempData = new List<dynamic>();
            var visitedRows = new HashSet<int>();

            // ����� ������ ������� ����
            int startRow = -1;
            for (int row = clickedRow; row >= 3; row--)
            {
                string numberCell = dataGridView1.Rows[row].Cells[0].Value?.ToString();
                string nameCell = dataGridView1.Rows[row].Cells[1].Value?.ToString();

                if (!string.IsNullOrWhiteSpace(numberCell) && !string.IsNullOrWhiteSpace(nameCell))
                {
                    string trimmedNameCell = nameCell.Trim();

                    // �������� �� "�� ������� ��� ��������"
                    if (trimmedNameCell.StartsWith("�� ") && trimmedNameCell.Split(' ').Length == 4)
                    {
                        string[] nameParts = trimmedNameCell.Split(' ');
                        // ��������� ����� ��������: "�� ������� �.�."
                        currentProductName = $"{nameParts[0]} {nameParts[1]} {nameParts[2][0]}.{nameParts[3][0]}.";
                    }
                    else
                    {
                        // ���� ��� "�� ������� �.�.", ���������� ��� ����
                        currentProductName = trimmedNameCell;
                    }
                    startRow = row;
                    break;
                }
            }

            if (startRow == -1)
            {
                MessageBox.Show("�� ������� ���������� ����.", "������", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // ����� ����� ������� ����
            int endRow = dataGridView1.Rows.Count - 2; // ���������, ��� ��������� ������ �� ��������������
            for (int row = startRow + 1; row <= endRow; row++)
            {
                string numberCell = dataGridView1.Rows[row].Cells[0].Value?.ToString();
                string nameCell = dataGridView1.Rows[row].Cells[1].Value?.ToString();

                if (!string.IsNullOrWhiteSpace(numberCell) && !string.IsNullOrWhiteSpace(nameCell))
                {
                    endRow = row - 1;
                    break;
                }
            }

            // ������������ ������ � �������� ����
            for (int row = startRow; row <= endRow; row++)
            {
                if (!visitedRows.Contains(row))
                {
                    string shippingDate = dataGridView1.Rows[row].Cells[4].Value?.ToString();  // ���� ��������
                    string unloadPoint = dataGridView1.Rows[row].Cells[13].Value?.ToString(); // ����� ���������

                    // ������� ���������� unloadPoint ������ ����
                    int matchingRowCount = 1; // ������� ������ ��� ���������
                    List<string> shippingDates = new List<string> { shippingDate }; // ��� ����� ���� ���

                    for (int checkRow = row + 1; checkRow <= endRow; checkRow++)
                    {
                        string checkUnloadPoint = dataGridView1.Rows[checkRow].Cells[13].Value?.ToString();
                        string checkShippingDate = dataGridView1.Rows[checkRow].Cells[4].Value?.ToString();

                        if (!string.IsNullOrWhiteSpace(checkUnloadPoint) &&
                            checkUnloadPoint == unloadPoint &&
                            !visitedRows.Contains(checkRow))
                        {
                            matchingRowCount++;
                            shippingDates.Add(checkShippingDate);
                            visitedRows.Add(checkRow); // ��������, ��� ������ ��� ����������
                        }
                    }

                    visitedRows.Add(row); // �������� ������� ������ ��� ������������

                    // ���� ���� ����������, ��������� ���� � ���������� ��������
                    shippingDates = shippingDates.Where(d => !string.IsNullOrWhiteSpace(d)).ToList();
                    string dateRange = shippingDates.Count > 1
                        ? $"{shippingDates.Min()} - {shippingDates.Max()}"
                        : shippingDates.FirstOrDefault() ?? "";

                    // ��������� ������ �� ��������� �������
                    tempData.Add(new
                    {
                        Number = tempData.Count + 1, // ��������� � 1
                        ShippingDate = dateRange,
                        UnloadPoint = unloadPoint,
                        Quantity = matchingRowCount
                    });
                }
            }

            // ������������ ����� ���������� ����� � ����
            orderCount = endRow - startRow + 1;

            // ������� ���������
            MessageBox.Show($"������ \"{currentProductName}\". ���������� �������: {orderCount}",
                "����������", MessageBoxButtons.OK, MessageBoxIcon.Information);

            // ��������� label8
            label8.Text = $"���������� ��� ����� ��: {currentProductName}";

            // ��������� dataGridView4 ��������� ��������
            var tempTable = new DataTable();
            tempTable.Columns.Add("�����");
            tempTable.Columns.Add("���� ��������");
            tempTable.Columns.Add("����� ���������");
            tempTable.Columns.Add("����������");

            foreach (var entry in tempData)
            {
                tempTable.Rows.Add(entry.Number, entry.ShippingDate, entry.UnloadPoint, entry.Quantity);
            }

            dataGridView4.DataSource = tempTable;

            // ����������� ������ ������� � dataGridView4
            dataGridView4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dataGridView4.Columns[0].Width = 50; // �����
            dataGridView4.Columns[1].Width = 150; // ���� ��������
            dataGridView4.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; // ����� ���������
            dataGridView4.Columns[3].Width = 80; // ����������
        }







        private void button7_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                Title = "�������� ���� Excel"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;

                // ��������� ���� Excel � DataGridView1
                LoadExcelToDataGridView(filePath, dataGridView1);

                // �������� ������ 7 ����� �������� �������
                button7.Visible = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // ������� ������� DataGridView1
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();

            // ���������� ������ 7 �����
            button7.Visible = true;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                Title = "�������� ���� � ������ �������"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                headerFilePath = openFileDialog.FileName; // ��������� ���� � ���������� �����

                using (var package = new ExcelPackage(new FileInfo(headerFilePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    string cellValue = worksheet.Cells["A1"].Text; // ��������� �������� ������ A1

                    // ������� ����� "�������������� ���������������", ���� ��� ����
                    string cleanedValue = cellValue.Replace("�������������� ���������������", "").Trim();

                    // ������� �������� � label7
                    label7.Text = $"��������� �����: {cleanedValue}";
                }
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {
            headerFilePath = null; // ������� ���������� � ����� � �����
            label7.Text = "��������� �����: "; // ��������� ����� �����
        }

        private void button9_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                Title = "�������� ���� ��������"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                priceListFilePath = openFileDialog.FileName;
                label4.Text = "������: ����������";
            }
        }


        private void button3_Click(object sender, EventArgs e)
        {
            priceListFilePath = null; // ������� ���������� � ����� � �����
            label4.Text = "������: "; // ��������� ����� �����
        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Filter = "Excel Files|*.xlsx",
                    Title = "�������� ���� � �������� ���������"
                };

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    contractNumbersFilePath = openFileDialog.FileName;
                    label9.Text = "������: ����������";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"������ ��� ����������� �����: {ex.Message}", "������", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                contractNumbersFilePath = string.Empty;
                label9.Text = "������: ";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"������ ��� ������� ����: {ex.Message}", "������", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private double FindPriceInPriceList(string address)
        {
            try
            {
                if (string.IsNullOrEmpty(priceListFilePath))
                {
                    throw new Exception("���� �������� �� ������. ������� ��� � ������� ������ ������ ��������.");
                }

                using (ExcelPackage package = new ExcelPackage(new FileInfo(priceListFilePath)))
                {
                    var priceSheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (priceSheet == null || priceSheet.Dimension == null)
                    {
                        throw new Exception("���� �������� ���� ��� �� �������� ������.");
                    }

                    int startRow = priceSheet.Dimension.Start.Row;
                    int endRow = priceSheet.Dimension.End.Row;

                    for (int row = startRow; row <= endRow; row++)
                    {
                        // ������ ����� �� ������� D
                        string cellAddress = priceSheet.Cells[row, 4]?.Text?.Trim(); // ������� D
                        if (string.Equals(cellAddress, address, StringComparison.OrdinalIgnoreCase))
                        {
                            // ������ ���� �� ������� E
                            string priceText = priceSheet.Cells[row, 5]?.Text?.Trim(); // ������� E
                            if (double.TryParse(priceText, out double price))
                            {
                                return price;
                            }
                            else
                            {
                                throw new Exception($"�������� ������ ���� � ������ {row} ��� ������ \"{address}\".");
                            }
                        }
                    }
                }

                throw new Exception($"����� \"{address}\" �� ������ � ����� ��������.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"������ � ������ ����: {ex.Message}", "������", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 0.0; // ���������� 0, ���� ��������� ������
            }
        }

        private string FindAbbreviationInPriceList(string address)
        {
            try
            {
                if (string.IsNullOrEmpty(priceListFilePath))
                {
                    throw new Exception("���� �������� �� ������. ������� ��� � ������� ������ ������ ��������.");
                }

                using (ExcelPackage package = new ExcelPackage(new FileInfo(priceListFilePath)))
                {
                    var priceSheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (priceSheet == null || priceSheet.Dimension == null)
                    {
                        throw new Exception("���� �������� ���� ��� �� �������� ������.");
                    }

                    int startRow = priceSheet.Dimension.Start.Row;
                    int endRow = priceSheet.Dimension.End.Row;

                    for (int row = startRow; row <= endRow; row++)
                    {
                        // ������ ����� �� ������� D
                        string cellAddress = priceSheet.Cells[row, 4]?.Text?.Trim(); // ������� D
                        if (string.Equals(cellAddress, address, StringComparison.OrdinalIgnoreCase))
                        {
                            // ������ ���������� �� ������� C
                            string abbreviation = priceSheet.Cells[row, 3]?.Text?.Trim(); // ������� C
                            if (!string.IsNullOrEmpty(abbreviation))
                            {
                                return abbreviation;
                            }
                            else
                            {
                                throw new Exception($"���������� �� ������� � ������ {row} ��� ������ \"{address}\".");
                            }
                        }
                    }
                }

                throw new Exception($"����� \"{address}\" �� ������ � ����� ��������.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"������ ��� ������ ����������: {ex.Message}", "������", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return string.Empty;
            }
        }

        private string FindIGKInPriceList(string address)
        {
            try
            {
                if (string.IsNullOrEmpty(priceListFilePath))
                {
                    throw new Exception("���� �������� �� ������. ������� ��� � ������� ������ ������ ��������.");
                }

                using (ExcelPackage package = new ExcelPackage(new FileInfo(priceListFilePath)))
                {
                    var priceSheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (priceSheet == null || priceSheet.Dimension == null)
                    {
                        throw new Exception("���� �������� ���� ��� �� �������� ������.");
                    }

                    int startRow = priceSheet.Dimension.Start.Row;
                    int endRow = priceSheet.Dimension.End.Row;

                    for (int row = startRow; row <= endRow; row++)
                    {
                        // ������ ����� �� ������� D
                        string cellAddress = priceSheet.Cells[row, 4]?.Text?.Trim(); // ������� D
                        if (string.Equals(cellAddress, address, StringComparison.OrdinalIgnoreCase))
                        {
                            // ������ �������� �� ������� G (���)
                            string igk = priceSheet.Cells[row, 7]?.Text?.Trim(); // ������� G
                            return igk; // ������� �������� (����� ���� ������ ��� ���)
                        }
                    }
                }

                throw new Exception($"����� \"{address}\" �� ������ � ����� ��������.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"������ ��� ������ ���: {ex.Message}", "������", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return string.Empty;
            }
        }

        private string GetFootingFromTable(string igk, string productName)
        {

            using (ExcelPackage package = new ExcelPackage(new FileInfo(contractNumbersFilePath)))
            {
                var sheet = package.Workbook.Worksheets.FirstOrDefault();
                if (sheet == null)
                {
                    throw new Exception("�� ������� ������� ��������� �������.");
                }

                // �������� �� ���� ������� �������
                int rowCount = sheet.Dimension.Rows;
                for (int row = 1; row <= rowCount; row++)
                {
                    // ��������� �������� �� ������� B, G � H
                    string columnBValue = sheet.Cells[row, 2]?.Value?.ToString();
                    string columnGValue = sheet.Cells[row, 7]?.Value?.ToString(); // ������� G (7-�)
                    string columnHValue = sheet.Cells[row, 8]?.Value?.ToString(); // ������� H (8-�)

                    // ��������� �������� � �������� B � H � ��������
                    if (columnHValue == igk && columnBValue == productName)
                    {
                        return columnGValue; // ���������� �������� �� ������� G
                    }
                }
            }

            throw new Exception($"�� ������� ����� ������� � IGK '{igk}' � ��������� �������� '{productName}'.");
        }


        public static string ConvertSumToWords(double sum)
        {
            // ��������� �� ����� ������
            int rubles = (int)Math.Floor(sum);

            // ���������� ����� Format �� NickBuhro.NumToWords
            string rublesText = RussianConverter.Format(rubles) + " ������";

            // ��������� ������������� �������� ������
            string kopecksText = "00 ������";

            // �������� ������
            string result = $"{rublesText} {kopecksText}";

            // ����������� ������ ����� � ������� �������
            return char.ToUpper(result[0]) + result.Substring(1);
        }




        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                string accnum = textBox1.Text;
                string accdata = textBox2.Text;

                if (dataGridView4.Rows.Count == 0)
                {
                    MessageBox.Show("��������� ������� �����. ��������� �� ������� ����� ��������� ������.",
                                    "������", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (string.IsNullOrEmpty(headerFilePath))
                {
                    MessageBox.Show("�� ������ ���� ��� �����. ������� ��� � ������� ��������������� ������.",
                                    "������", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string tempFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "temp.xlsx");

                // �������� ����������� �������
                List<DataGridViewRow> excludedOrders = new List<DataGridViewRow>();
                int currentOrderNumber = 1; // ��������� �������

                // ��� 1: �������� ���������� �����
                using (ExcelPackage package = new ExcelPackage())
                {
                    var sheet = package.Workbook.Worksheets.Add("����");

                    // ������������� ����� ����� Arial
                    sheet.Cells.Style.Font.Name = "Arial";

                    // ��������� �����
                    sheet.Cells["A8:I8"].Merge = true;
                    sheet.Cells["A8:I8"].Value = $"���� �{accnum} �� {accdata}�";
                    sheet.Cells["A8:I8"].Style.Font.Size = 14;
                    sheet.Cells["A8:I8"].Style.Font.Bold = true;
                    sheet.Cells["A8:I8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    sheet.Cells["A8:I8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    // ����������
                    sheet.Cells["A9:I9"].Merge = true;
                    sheet.Cells["A9:I9"].Value = "����������: ��� ������� ��� 6164134558";
                    sheet.Cells["A9:I9"].Style.Font.Size = 10;
                    sheet.Cells["A9:I9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                    // ��������� �������
                    string[] headers = { "�", "������������ ������", "", "��. ���������", "����������", "����", "", "�����", "" };
                    for (int i = 0; i < headers.Length; i++)
                    {
                        sheet.Cells[12, i + 1].Value = headers[i];
                        sheet.Cells[12, i + 1].Style.Font.Size = 10;
                        sheet.Cells[12, i + 1].Style.WrapText = true;
                        sheet.Cells[12, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        sheet.Cells[12, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        sheet.Cells[12, i + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    }

                    // ���������� ���������
                    sheet.Cells["B12:C12"].Merge = true;
                    sheet.Cells["F12:G12"].Merge = true;
                    sheet.Cells["H12:I12"].Merge = true;

                    int currentRow = 13;
                    double totalSum = 0;

                    foreach (DataGridViewRow gridRow in dataGridView4.Rows)
                    {
                        if (gridRow.Cells[0].Value == null || gridRow.Cells[2].Value == null || gridRow.Cells[3].Value == null)
                            continue; // ���������� ������ ������

                        try
                        {
                            string �������������� = gridRow.Cells[2].Value.ToString();
                            string ������������ = gridRow.Cells[1].Value.ToString();
                            int ���������� = Convert.ToInt32(gridRow.Cells[3].Value);
                            string ���������������� = FindAbbreviationInPriceList(��������������);
                            double ���� = FindPriceInPriceList(��������������);

                            // ��������� ������� ���
                            string ��� = FindIGKInPriceList(��������������);
                            if (!string.IsNullOrEmpty(���))
                            {
                                excludedOrders.Add(gridRow);
                                continue; // ���������� ������ � ���
                            }

                            double ����� = ���������� * ����;
                            totalSum += �����;

                            // ��������� ������
                            sheet.Cells[currentRow, 1].Value = currentOrderNumber++;
                            sheet.Cells[currentRow, 2].Value = $"������������ ������ ��. �����������-{����������������} {������������}";
                            sheet.Cells[currentRow, 2, currentRow, 3].Merge = true;
                            sheet.Cells[currentRow, 4].Value = "����";
                            sheet.Cells[currentRow, 5].Value = ����������;
                            sheet.Cells[currentRow, 6].Value = $"{����:0.00}";
                            sheet.Cells[currentRow, 6, currentRow, 7].Merge = true;
                            sheet.Cells[currentRow, 8].Value = $"{�����:0,0.00}";
                            sheet.Cells[currentRow, 8, currentRow, 9].Merge = true;

                            for (int col = 1; col <= 9; col++)
                            {
                                sheet.Cells[currentRow, col].Style.Font.Size = 10;
                                sheet.Cells[currentRow, col].Style.WrapText = true;
                                sheet.Cells[currentRow, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                sheet.Cells[currentRow, col].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                                if (col != 2)
                                    sheet.Cells[currentRow, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            }
                            sheet.Cells[currentRow, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                            currentRow++;
                        }
                        catch (Exception exRow)
                        {
                            MessageBox.Show($"������ ��� ��������� ������: {exRow.Message}", "������ ������",
                                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }

                    // �������� ������
                    sheet.Cells[currentRow, 6, currentRow, 7].Merge = true;
                    sheet.Cells[currentRow, 6, currentRow, 7].Value = "�����:";
                    sheet.Cells[currentRow, 6, currentRow, 7].Style.Font.Size = 10;
                    sheet.Cells[currentRow, 6, currentRow, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    sheet.Cells[currentRow, 8, currentRow, 9].Merge = true;
                    sheet.Cells[currentRow, 8, currentRow, 9].Value = $"{totalSum:0,0.00}";
                    sheet.Cells[currentRow, 8, currentRow, 9].Style.Font.Size = 10;
                    sheet.Cells[currentRow, 8, currentRow, 9].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    sheet.Cells[currentRow, 8, currentRow, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    sheet.Cells[currentRow, 8, currentRow, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    currentRow++;
                    sheet.Cells[currentRow, 6, currentRow, 7].Merge = true;
                    sheet.Cells[currentRow, 6, currentRow, 7].Value = "��� ������(���).";
                    sheet.Cells[currentRow, 6, currentRow, 7].Style.Font.Size = 10;
                    sheet.Cells[currentRow, 6, currentRow, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    sheet.Cells[currentRow, 8, currentRow, 9].Merge = true;
                    sheet.Cells[currentRow, 8, currentRow, 9].Style.Font.Size = 10;
                    sheet.Cells[currentRow, 8, currentRow, 9].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    currentRow++;
                    sheet.Cells[currentRow, 6, currentRow, 7].Merge = true;
                    sheet.Cells[currentRow, 6, currentRow, 7].Value = "����� � ������:";
                    sheet.Cells[currentRow, 6, currentRow, 7].Style.Font.Size = 10;
                    sheet.Cells[currentRow, 6, currentRow, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    sheet.Cells[currentRow, 8, currentRow, 9].Merge = true;
                    sheet.Cells[currentRow, 8, currentRow, 9].Value = $"{totalSum:0,0.00}";
                    sheet.Cells[currentRow, 8, currentRow, 9].Style.Font.Size = 10;
                    sheet.Cells[currentRow, 8, currentRow, 9].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    sheet.Cells[currentRow, 8, currentRow, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    sheet.Cells[currentRow, 8, currentRow, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    // ����� ���������� � ������������� � �����
                    currentRow++;
                    int totalItems = currentOrderNumber - 1; // ��������� ����� ������������ �� ������� A

                    sheet.Cells[currentRow, 1].Merge = false;
                    sheet.Cells[currentRow, 1].Value = $"����� ������������ {totalItems}, �� ����� {Math.Floor(totalSum):N0}-{(totalSum % 1 * 100):00}";
                    sheet.Cells[currentRow, 1].Style.Font.Size = 10;
                    sheet.Cells[currentRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                    // ������� ����� � �����
                    currentRow++;
                    sheet.Cells[currentRow, 1].Merge = false;
                    sheet.Cells[currentRow, 1].Value = ConvertSumToWords(totalSum);
                    sheet.Cells[currentRow, 1].Style.Font.Size = 10;
                    sheet.Cells[currentRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                    // ������ ���� �� 3 ������
                    currentRow += 4;

                    // ��������� ������ � ���������
                    sheet.Cells[currentRow, 1].Value = $"������������ �����������____________________({currentProductName})";
                    sheet.Cells[currentRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    currentRow += 2;

                    sheet.Cells[currentRow, 1].Value = $"������� ��������� __________________________ ({currentProductName})";
                    sheet.Cells[currentRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                    // ��������� ������ � ������� ������ ��� ��������
                    sheet.Cells[currentRow - 2, 1, currentRow, 1].Style.Font.Size = 10;

                    // ���������� ���������� �����
                    package.SaveAs(new FileInfo(tempFilePath));
                }

                // ��� 2: ����������� ���������� ����� � ������ �����
                using (ExcelPackage headerPackage = new ExcelPackage(new FileInfo(headerFilePath)))
                using (ExcelPackage tempPackage = new ExcelPackage(new FileInfo(tempFilePath)))
                {
                    var headerSheet = headerPackage.Workbook.Worksheets.FirstOrDefault();
                    var tempSheet = tempPackage.Workbook.Worksheets.FirstOrDefault();

                    if (headerSheet == null || tempSheet == null)
                    {
                        throw new Exception("������ ������ ������ �� ���������� ����� ��� ����� �����.");
                    }

                    using (ExcelPackage mergedPackage = new ExcelPackage())
                    {
                        var mergedSheet = mergedPackage.Workbook.Worksheets.Add("�������� �����");

                        // �������� �����
                        var headerRange = headerSheet.Dimension;
                        headerSheet.Cells[headerRange.Start.Row, headerRange.Start.Column, headerRange.End.Row, headerRange.End.Column]
                            .Copy(mergedSheet.Cells[headerRange.Start.Row, headerRange.Start.Column]);

                        // �������� ��������� ������� ��� ������
                        var tempRange = tempSheet.Dimension;
                        int offsetRow = headerRange?.End.Row + 1 ?? 1;
                        tempSheet.Cells[tempRange.Start.Row, tempRange.Start.Column, tempRange.End.Row, tempRange.End.Column]
                            .Copy(mergedSheet.Cells[offsetRow, tempRange.Start.Column]);

                        // ����� ���������� �� ����������� �������
                        if (excludedOrders.Count > 0)
                        {
                            string excludedInfo = "��������� ������ ���� ��������� �� �����:\n";
                            foreach (var row in excludedOrders)
                            {
                                string �������������� = row.Cells[2].Value.ToString();
                                excludedInfo += $"- {��������������}\n";
                            }
                            MessageBox.Show(excludedInfo, "����������� ������", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }

                        // ��������� �������� ����
                        SaveFileDialog saveFileDialog = new SaveFileDialog
                        {
                            Filter = "Excel Files|*.xlsx",
                            Title = "��������� �������� ����"
                        };

                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            mergedPackage.SaveAs(new FileInfo(saveFileDialog.FileName));
                            MessageBox.Show("���� ������� ������!", "�����", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }

                // ������� ��������� ����
                if (File.Exists(tempFilePath))
                {
                    File.Delete(tempFilePath);
                }

                // ����� ����: ��������� excludedOrders
                while (excludedOrders.Count > 0)
                {
                    var currentRow = excludedOrders[0];
                    string address = currentRow.Cells[2].Value?.ToString();
                    double quantity = currentRow.Cells[3].Value != null && double.TryParse(currentRow.Cells[3].Value.ToString(), out double parsedQuantity)
    ? parsedQuantity
    : 0;
                    string date = currentRow.Cells[1].Value?.ToString();
                    string currentIGK = FindIGKInPriceList(address);
                    string currentFooting = GetFootingFromTable(currentIGK, currentProductName);


                    // ����� ���� ��� ��������� ������ � ����
                    using (Form dialog = new Form())
                    {
                        dialog.Width = 900;
                        dialog.Height = 300;
                        dialog.Text = "��������� ������ � ���";

                        // ����� � ����������� �� �������������� ������
                        Label labelOrderInfo = new Label
                        {
                            Text = $"����������: {currentProductName}\n�����: {address}\n���� ������: {date}\n���������� ������: {quantity}",
                            Left = 10,
                            Top = 10,
                            Width = 850, // ������ ����
                            Height = 80, // ������ ��� 4 ������ ������
                            Font = new Font("Arial", 12, FontStyle.Bold)
                        };

                        // ����� � TextBox ��� ������ ������
                        Label labelNumber = new Label
                        {
                            Text = "���� �:",
                            Left = 10,
                            Top = 100,
                            Width = 150,
                            Font = new Font("Arial", 10)
                        };
                        TextBox textBoxNumber = new TextBox
                        {
                            Left = 180,
                            Top = 100,
                            Width = 225
                        };

                        // ����� � TextBox ��� ����� ����
                        Label labelDate = new Label
                        {
                            Text = "��:",
                            Left = 10,
                            Top = 160,
                            Width = 150,
                            Font = new Font("Arial", 10)
                        };
                        TextBox textBoxDate = new TextBox
                        {
                            Left = 180,
                            Top = 160,
                            Width = 225
                        };

                        // ������ "OK"
                        Button buttonOk = new Button
                        {
                            Text = "OK",
                            DialogResult = DialogResult.OK,
                            Left = 180,
                            Top = 220,
                            Width = 120,
                            Height = 40,
                            Font = new Font("Arial", 10, FontStyle.Bold)
                        };

                        // ���������� ��������� � ������
                        dialog.Controls.Add(labelOrderInfo);
                        dialog.Controls.Add(labelNumber);
                        dialog.Controls.Add(textBoxNumber);
                        dialog.Controls.Add(labelDate);
                        dialog.Controls.Add(textBoxDate);
                        dialog.Controls.Add(buttonOk);

                        // ��������� ������ "OK" ��� AcceptButton
                        dialog.AcceptButton = buttonOk;


                        if (dialog.ShowDialog() == DialogResult.OK)
                        {
                            // �������� ����� ������ �� TextBox
                            string newNumber = textBoxNumber.Text;
                            string newDate = textBoxDate.Text;

                            // ��� 1: �������� ���������� �����
                            using (ExcelPackage package = new ExcelPackage())
                            {
                                var sheet = package.Workbook.Worksheets.Add("����");

                                // ������������� ����� ����� Arial
                                sheet.Cells.Style.Font.Name = "Arial";

                                // ��������� �����
                                sheet.Cells["A8:I8"].Merge = true;
                                sheet.Cells["A8:I8"].Value = $"���� �{newNumber} �� {newDate}�";
                                sheet.Cells["A8:I8"].Style.Font.Size = 14;
                                sheet.Cells["A8:I8"].Style.Font.Bold = true;
                                sheet.Cells["A8:I8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                sheet.Cells["A8:I8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                // ����������
                                sheet.Cells["A9:I9"].Merge = true;
                                sheet.Cells["A9:I9"].Value = "����������: ��� ������� ��� 6164134558";
                                sheet.Cells["A9:I9"].Style.Font.Size = 10;
                                sheet.Cells["A9:I9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                                // ���������
                                sheet.Cells["A10:I10"].Merge = true;
                                sheet.Cells["A10:I10"].Value = currentFooting;
                                sheet.Cells["A10:I10"].Style.Font.Size = 10;
                                sheet.Cells["A10:I10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                                // ���
                                sheet.Cells["A11:I11"].Merge = true;
                                sheet.Cells["A11:I11"].Value = currentIGK;
                                sheet.Cells["A11:I11"].Style.Font.Size = 10;
                                sheet.Cells["A11:I11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                                // ��������� �������
                                string[] headers = { "�", "������������ ������", "", "��. ���������", "����������", "����", "", "�����", "" };
                                for (int i = 0; i < headers.Length; i++)
                                {
                                    sheet.Cells[12, i + 1].Value = headers[i];
                                    sheet.Cells[12, i + 1].Style.Font.Size = 10;
                                    sheet.Cells[12, i + 1].Style.WrapText = true;
                                    sheet.Cells[12, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    sheet.Cells[12, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    sheet.Cells[12, i + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                }

                                // ���������� ���������
                                sheet.Cells["B12:C12"].Merge = true;
                                sheet.Cells["F12:G12"].Merge = true;
                                sheet.Cells["H12:I12"].Merge = true;

                                int IGKcurrentRow = 13;
                                double totalSum = 0;

                                try
                                {
                                    string ���������������� = FindAbbreviationInPriceList(address);
                                    double ���� = FindPriceInPriceList(address);

                                    double ����� = quantity * ����;
                                    totalSum += �����;

                                    // ��������� ������
                                    sheet.Cells[IGKcurrentRow, 1].Value = 1;
                                    sheet.Cells[IGKcurrentRow, 2].Value = $"������������ ������ ��. �����������-{����������������} {date}";
                                    sheet.Cells[IGKcurrentRow, 2, IGKcurrentRow, 3].Merge = true;
                                    sheet.Cells[IGKcurrentRow, 4].Value = "����";
                                    sheet.Cells[IGKcurrentRow, 5].Value = quantity;
                                    sheet.Cells[IGKcurrentRow, 6].Value = $"{����:0.00}";
                                    sheet.Cells[IGKcurrentRow, 6, IGKcurrentRow, 7].Merge = true;
                                    sheet.Cells[IGKcurrentRow, 8].Value = $"{�����:0,0.00}";
                                    sheet.Cells[IGKcurrentRow, 8, IGKcurrentRow, 9].Merge = true;

                                    for (int col = 1; col <= 9; col++)
                                    {
                                        sheet.Cells[IGKcurrentRow, col].Style.Font.Size = 10;
                                        sheet.Cells[IGKcurrentRow, col].Style.WrapText = true;
                                        sheet.Cells[IGKcurrentRow, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                        sheet.Cells[IGKcurrentRow, col].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                                        if (col != 2)
                                            sheet.Cells[IGKcurrentRow, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    }
                                    sheet.Cells[IGKcurrentRow, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                                    IGKcurrentRow++;
                                }
                                catch (Exception exRow)
                                {
                                    MessageBox.Show($"������ ��� ��������� ������: {exRow.Message}", "������ ������",
                                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }


                                // �������� ������
                                sheet.Cells[IGKcurrentRow, 6, IGKcurrentRow, 7].Merge = true;
                                sheet.Cells[IGKcurrentRow, 6, IGKcurrentRow, 7].Value = "�����:";
                                sheet.Cells[IGKcurrentRow, 6, IGKcurrentRow, 7].Style.Font.Size = 10;
                                sheet.Cells[IGKcurrentRow, 6, IGKcurrentRow, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                                sheet.Cells[IGKcurrentRow, 8, IGKcurrentRow, 9].Merge = true;
                                sheet.Cells[IGKcurrentRow, 8, IGKcurrentRow, 9].Value = $"{totalSum:0,0.00}";
                                sheet.Cells[IGKcurrentRow, 8, IGKcurrentRow, 9].Style.Font.Size = 10;
                                sheet.Cells[IGKcurrentRow, 8, IGKcurrentRow, 9].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                sheet.Cells[IGKcurrentRow, 8, IGKcurrentRow, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                sheet.Cells[IGKcurrentRow, 8, IGKcurrentRow, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                IGKcurrentRow++;
                                sheet.Cells[IGKcurrentRow, 6, IGKcurrentRow, 7].Merge = true;
                                sheet.Cells[IGKcurrentRow, 6, IGKcurrentRow, 7].Value = "��� ������(���).";
                                sheet.Cells[IGKcurrentRow, 6, IGKcurrentRow, 7].Style.Font.Size = 10;
                                sheet.Cells[IGKcurrentRow, 6, IGKcurrentRow, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                                sheet.Cells[IGKcurrentRow, 8, IGKcurrentRow, 9].Merge = true;
                                sheet.Cells[IGKcurrentRow, 8, IGKcurrentRow, 9].Style.Font.Size = 10;
                                sheet.Cells[IGKcurrentRow, 8, IGKcurrentRow, 9].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                                IGKcurrentRow++;
                                sheet.Cells[IGKcurrentRow, 6, IGKcurrentRow, 7].Merge = true;
                                sheet.Cells[IGKcurrentRow, 6, IGKcurrentRow, 7].Value = "����� � ������:";
                                sheet.Cells[IGKcurrentRow, 6, IGKcurrentRow, 7].Style.Font.Size = 10;
                                sheet.Cells[IGKcurrentRow, 6, IGKcurrentRow, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                                sheet.Cells[IGKcurrentRow, 8, IGKcurrentRow, 9].Merge = true;
                                sheet.Cells[IGKcurrentRow, 8, IGKcurrentRow, 9].Value = $"{totalSum:0,0.00}";
                                sheet.Cells[IGKcurrentRow, 8, IGKcurrentRow, 9].Style.Font.Size = 10;
                                sheet.Cells[IGKcurrentRow, 8, IGKcurrentRow, 9].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                sheet.Cells[IGKcurrentRow, 8, IGKcurrentRow, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                sheet.Cells[IGKcurrentRow, 8, IGKcurrentRow, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                // ����� ���������� � ������������� � �����
                                IGKcurrentRow++;
                                int totalItems = currentOrderNumber - 1; // ��������� ����� ������������ �� ������� A

                                sheet.Cells[IGKcurrentRow, 1].Merge = false;
                                sheet.Cells[IGKcurrentRow, 1].Value = $"����� ������������ {totalItems}, �� ����� {Math.Floor(totalSum):N0}-{(totalSum % 1 * 100):00}";
                                sheet.Cells[IGKcurrentRow, 1].Style.Font.Size = 10;
                                sheet.Cells[IGKcurrentRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                                // ������� ����� � �����
                                IGKcurrentRow++;
                                sheet.Cells[IGKcurrentRow, 1].Merge = false;
                                sheet.Cells[IGKcurrentRow, 1].Value = ConvertSumToWords(totalSum);
                                sheet.Cells[IGKcurrentRow, 1].Style.Font.Size = 10;
                                sheet.Cells[IGKcurrentRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                                // ������ ���� �� 3 ������
                                IGKcurrentRow += 4;

                                // ��������� ������ � ���������
                                sheet.Cells[IGKcurrentRow, 1].Value = $"������������ �����������____________________({currentProductName})";
                                sheet.Cells[IGKcurrentRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                IGKcurrentRow += 2;

                                sheet.Cells[IGKcurrentRow, 1].Value = $"������� ��������� __________________________ ({currentProductName})";
                                sheet.Cells[IGKcurrentRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                                // ��������� ������ � ������� ������ ��� ��������
                                sheet.Cells[IGKcurrentRow - 2, 1, IGKcurrentRow, 1].Style.Font.Size = 10;

                                // ���������� ���������� �����
                                package.SaveAs(new FileInfo(tempFilePath));
                            }

                            // ��� 2: ����������� ���������� ����� � ������ �����
                            using (ExcelPackage headerPackage = new ExcelPackage(new FileInfo(headerFilePath)))
                            using (ExcelPackage tempPackage = new ExcelPackage(new FileInfo(tempFilePath)))
                            {
                                var headerSheet = headerPackage.Workbook.Worksheets.FirstOrDefault();
                                var tempSheet = tempPackage.Workbook.Worksheets.FirstOrDefault();

                                if (headerSheet == null || tempSheet == null)
                                {
                                    throw new Exception("������ ������ ������ �� ���������� ����� ��� ����� �����.");
                                }

                                using (ExcelPackage mergedPackage = new ExcelPackage())
                                {
                                    var mergedSheet = mergedPackage.Workbook.Worksheets.Add("�������� �����");

                                    // �������� �����
                                    var headerRange = headerSheet.Dimension;
                                    headerSheet.Cells[headerRange.Start.Row, headerRange.Start.Column, headerRange.End.Row, headerRange.End.Column]
                                        .Copy(mergedSheet.Cells[headerRange.Start.Row, headerRange.Start.Column]);

                                    // �������� ��������� ������� ��� ������
                                    var tempRange = tempSheet.Dimension;
                                    int offsetRow = headerRange?.End.Row + 1 ?? 1;
                                    tempSheet.Cells[tempRange.Start.Row, tempRange.Start.Column, tempRange.End.Row, tempRange.End.Column]
                                        .Copy(mergedSheet.Cells[offsetRow, tempRange.Start.Column]);

                                    // ��������� �������� ����
                                    SaveFileDialog saveFileDialog = new SaveFileDialog
                                    {
                                        Filter = "Excel Files|*.xlsx",
                                        Title = "��������� �������� ����"
                                    };

                                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                                    {
                                        mergedPackage.SaveAs(new FileInfo(saveFileDialog.FileName));
                                        MessageBox.Show("���� ������� ������!", "�����", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    }
                                }
                            }

                            // ������� ��������� ����
                            if (File.Exists(tempFilePath))
                            {
                                File.Delete(tempFilePath);
                            }
                        }
                    }

                    // ������� ������� ������� �� ������
                    excludedOrders.RemoveAt(0);
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show($"������: {ex.Message}", "������", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }



        }

        private void button4_Click(object sender, EventArgs e)
        {
            // �������� ���� � �����������
            Form aboutDialog = new Form
            {
                Width = 480,
                Height = 350,
                Text = "� ���������",
                FormBorderStyle = FormBorderStyle.FixedDialog, // ������������� ������ ����
                MaximizeBox = false, // ������� ������ ������������
                MinimizeBox = false, // ������� ������ ����������
                StartPosition = FormStartPosition.CenterScreen // ���������� ���� �� ������
            };

            // ����� � �����������
            Label aboutText = new Label
            {
                Left = 10,
                Top = 10,
                Width = 360,
                Height = 300,
                Text = "��������� ������\n" +
                        "������: 1.0\n" +
                        "�����������: ������� ������, tg: @ReCream, github: recreamu\n" +
                        "\n" +
                        "������ ���������� ��������� ������������ ��������� ������� ������\n" +
                        "� ������� ����� Excel �� ������ ������ �� ���������� ������ ���������\n" +
                        "������ ���� �� �������. ��������� ������ � �������� � ��������� ������\n" +
                        "������ �������, �� �������� ���������� � ����������� ������������ ���\n" +
                        "������������. ��� �������� ��������� �������������� �������� ������:\n" +
                        "EPPlus(7.5.2) � NickVuhro.NumToWords(1.1.3), � ��� �� ChatGPT.\n" +
                        "�������� ��� ���� �� ���� GitHub, �������� ��� ������� ����� �����.\n" +
                        "\n" +
                        "��������� ��������������� ������������� � ��������������� �����.\n" +
                        "������������ ������������� �� ������������� �������������.",
                AutoSize = true,
                TextAlign = ContentAlignment.TopLeft
            };

            // ������ ��������
            Button closeButton = new Button
            {
                Text = "�������",
                Width = 80,
                Height = 30,
                Top = 260, // ������ ������ �� �������� ���� ����
                Left = (aboutDialog.ClientSize.Width - 80) / 2, // ���������� ������ ����
                DialogResult = DialogResult.OK
            };

            // ���������� ��������� � �����
            aboutDialog.Controls.Add(aboutText);
            aboutDialog.Controls.Add(closeButton);

            // ��������� ������ �������� �� ���������
            aboutDialog.AcceptButton = closeButton;

            // ����� ����
            aboutDialog.ShowDialog();
        }
    }
}

