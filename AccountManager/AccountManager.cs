using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using NickBuhro.NumToWords;
using NickBuhro.NumToWords.Russian;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;


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
        private string registryFilePath;

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
            button11.Click += button11_Click;
            button12.Click += button12_Click;
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

                // Добавляем заголовки как обычные строки
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    dataTable.Columns.Add($"Колонка {col}"); // Временно добавляем имена
                }

                // Записываем данные, включая заголовки
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

                // Отключаем автоматическое использование первой строки как заголовка
                targetGrid.ColumnHeadersVisible = false;
                targetGrid.AllowUserToAddRows = false; // Убираем пустую строку внизу
                targetGrid.RowHeadersVisible = false; // Убираем столбец с заголовками строк
            }
        }

        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex < 3 || e.RowIndex == dataGridView1.Rows.Count - 1) // Игнорируем строки 1-3 и последнюю строку
                return;

            int clickedRow = e.RowIndex;

            int orderCount = 0;

            // Для временной таблицы
            var tempData = new List<dynamic>();
            var visitedRows = new HashSet<int>();

            // Найти начало текущей зоны
            int startRow = -1;
            for (int row = clickedRow; row >= 3; row--)
            {
                string numberCell = dataGridView1.Rows[row].Cells[0].Value?.ToString();
                string nameCell = dataGridView1.Rows[row].Cells[1].Value?.ToString();

                if (!string.IsNullOrWhiteSpace(numberCell) && !string.IsNullOrWhiteSpace(nameCell))
                {
                    string trimmedNameCell = nameCell.Trim();

                    // Проверка на "ИП Фамилия Имя Отчество"
                    if (trimmedNameCell.StartsWith("ИП ") && trimmedNameCell.Split(' ').Length == 4)
                    {
                        string[] nameParts = trimmedNameCell.Split(' ');
                        // Формируем новое значение: "ИП Фамилия И.О."
                        currentProductName = $"{nameParts[0]} {nameParts[1]} {nameParts[2][0]}.{nameParts[3][0]}.";
                    }
                    else
                    {
                        // Если уже "ИП Фамилия И.О.", записываем как есть
                        currentProductName = trimmedNameCell;
                    }
                    startRow = row;
                    break;
                }
            }

            if (startRow == -1)
            {
                MessageBox.Show("Не удалось определить зону.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Найти конец текущей зоны
            int endRow = dataGridView1.Rows.Count - 2; // Учитываем, что последняя строка не обрабатывается
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

            // Обрабатываем строки в пределах зоны
            for (int row = startRow; row <= endRow; row++)
            {
                if (!visitedRows.Contains(row))
                {
                    string shippingDate = dataGridView1.Rows[row].Cells[4].Value?.ToString();  // Дата отгрузки
                    string unloadPoint = dataGridView1.Rows[row].Cells[13].Value?.ToString(); // Пункт разгрузки

                    // Подсчёт совпадений unloadPoint внутри зоны
                    int matchingRowCount = 1; // Текущая строка уже считается
                    List<string> shippingDates = new List<string> { shippingDate }; // Для сбора всех дат

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
                            visitedRows.Add(checkRow); // Отмечаем, что строка уже обработана
                        }
                    }

                    visitedRows.Add(row); // Отмечаем текущую строку как обработанную

                    // Если есть совпадения, сортируем даты и записываем диапазон
                    shippingDates = shippingDates.Where(d => !string.IsNullOrWhiteSpace(d)).ToList();
                    string dateRange = shippingDates.Count > 1
                        ? $"{shippingDates.Min()} - {shippingDates.Max()}"
                        : shippingDates.FirstOrDefault() ?? "";

                    // Добавляем строку во временную таблицу
                    tempData.Add(new
                    {
                        Number = tempData.Count + 1, // Нумерация с 1
                        ShippingDate = dateRange,
                        UnloadPoint = unloadPoint,
                        Quantity = matchingRowCount
                    });
                }
            }

            // Подсчитываем общее количество строк в зоне
            orderCount = endRow - startRow + 1;

            // Выводим сообщение
            MessageBox.Show($"Выбран \"{currentProductName}\". Количество заказов: {orderCount}",
                "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);

            // Обновляем label8
            label8.Text = $"Информация для счета по: {currentProductName}";

            // Заполняем dataGridView4 временной таблицей
            var tempTable = new DataTable();
            tempTable.Columns.Add("Номер");
            tempTable.Columns.Add("Дата отгрузки");
            tempTable.Columns.Add("Пункт разгрузки");
            tempTable.Columns.Add("Количество");

            foreach (var entry in tempData)
            {
                tempTable.Rows.Add(entry.Number, entry.ShippingDate, entry.UnloadPoint, entry.Quantity);
            }

            dataGridView4.DataSource = tempTable;

            // Настраиваем ширину колонок в dataGridView4
            dataGridView4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dataGridView4.Columns[0].Width = 50; // Номер
            dataGridView4.Columns[1].Width = 150; // Дата отгрузки
            dataGridView4.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; // Пункт разгрузки
            dataGridView4.Columns[3].Width = 80; // Количество
        }







        private void button7_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                Title = "Выберите файл Excel"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;

                // Загружаем файл Excel в DataGridView1
                LoadExcelToDataGridView(filePath, dataGridView1);

                // Скрываем кнопку 7 после загрузки таблицы
                button7.Visible = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Очистка таблицы DataGridView1
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();

            // Показываем кнопку 7 снова
            button7.Visible = true;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                Title = "Выберите файл с шапкой таблицы"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                headerFilePath = openFileDialog.FileName; // Сохраняем путь к выбранному файлу

                using (var package = new ExcelPackage(new FileInfo(headerFilePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    string cellValue = worksheet.Cells["A1"].Text; // Считываем значение ячейки A1

                    // Убираем слова "Индивидуальный предприниматель", если они есть
                    string cleanedValue = cellValue.Replace("Индивидуальный предприниматель", "").Trim();

                    // Выводим значение в label7
                    label7.Text = $"Выбранная шапка: {cleanedValue}";
                }
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {
            headerFilePath = null; // Очищаем переменную с путем к файлу
            label7.Text = "Выбранная шапка: "; // Обновляем текст метки
        }

        private void button9_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                Title = "Выберите файл ставок"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                priceListFilePath = openFileDialog.FileName;
                string fileName = Path.GetFileName(priceListFilePath);
                label4.Text = $"Статус: Подключено ({fileName})";
            }
        }


        private void button3_Click(object sender, EventArgs e)
        {
            priceListFilePath = null; // Очищаем переменную с путем к файлу
            label4.Text = "Статус: "; // Обновляем текст метки
        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Filter = "Excel Files|*.xlsx",
                    Title = "Выберите файл с номерами договоров"
                };

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    contractNumbersFilePath = openFileDialog.FileName;
                    string fileName = Path.GetFileName(contractNumbersFilePath);
                    label9.Text = $"Статус: Подключено ({fileName})";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при подключении файла: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                contractNumbersFilePath = string.Empty;
                label9.Text = "Статус: ";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при очистке пути: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private double FindPriceInPriceList(string address)
        {
            try
            {
                if (string.IsNullOrEmpty(priceListFilePath))
                {
                    throw new Exception("Файл расценок не выбран. Укажите его с помощью кнопки выбора расценок.");
                }

                using (ExcelPackage package = new ExcelPackage(new FileInfo(priceListFilePath)))
                {
                    var priceSheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (priceSheet == null || priceSheet.Dimension == null)
                    {
                        throw new Exception("Файл расценок пуст или не содержит данных.");
                    }

                    int startRow = priceSheet.Dimension.Start.Row;
                    int endRow = priceSheet.Dimension.End.Row;

                    for (int row = startRow; row <= endRow; row++)
                    {
                        string cellAddress = priceSheet.Cells[row, 4]?.Text?.Trim();
                        if (string.Equals(cellAddress, address, StringComparison.OrdinalIgnoreCase))
                        {
                            string priceWithoutGlueText = priceSheet.Cells[row, 5]?.Text?.Trim();
                            string priceWithGlueText = priceSheet.Cells[row, 6]?.Text?.Trim();

                            double priceWithoutGlue = double.TryParse(priceWithoutGlueText, out var tempWithoutGlue) ? tempWithoutGlue : 0;
                            double priceWithGlue = double.TryParse(priceWithGlueText, out var tempWithGlue) ? tempWithGlue : 0;

                            if (priceWithoutGlue > 0 && priceWithGlue > 0)
                            {
                                using (Form dialog = new Form())
                                {
                                    dialog.Text = "Выбор ставки";
                                    dialog.Padding = new Padding(10);
                                    dialog.StartPosition = FormStartPosition.CenterParent;
                                    dialog.AutoSizeMode = AutoSizeMode.GrowAndShrink;
                                    dialog.AutoSize = true;

                                    // Панель основного содержимого
                                    TableLayoutPanel contentPanel = new TableLayoutPanel
                                    {
                                        AutoSize = true,
                                        ColumnCount = 1,
                                        Dock = DockStyle.Fill,
                                        Padding = new Padding(10)
                                    };

                                    // Текстовая метка
                                    Label label = new Label
                                    {
                                        Text = $"Адрес \"{address}\" имеет две ставки:\nБез клея: {priceWithoutGlue}\nС клеем: {priceWithGlue}\nВыберите нужную.",
                                        AutoSize = true,
                                        MaximumSize = new Size(600, 0),
                                        Font = new Font("Arial", 10),
                                        Dock = DockStyle.Top
                                    };
                                    contentPanel.Controls.Add(label);

                                    // Панель для кнопок
                                    FlowLayoutPanel buttonPanel = new FlowLayoutPanel
                                    {
                                        FlowDirection = FlowDirection.LeftToRight,
                                        AutoSize = true,
                                        Dock = DockStyle.Top,
                                        Padding = new Padding(0, 10, 0, 0)
                                    };

                                    // Кнопка "Без клея"
                                    Button buttonWithoutGlue = new Button
                                    {
                                        Text = $"Без клея ({priceWithoutGlue})",
                                        AutoSize = true,
                                        Padding = new Padding(10)
                                    };
                                    buttonWithoutGlue.Click += (s, e) =>
                                    {
                                        dialog.DialogResult = DialogResult.Yes;
                                        dialog.Close();
                                    };

                                    // Кнопка "С клеем"
                                    Button buttonWithGlue = new Button
                                    {
                                        Text = $"С клеем ({priceWithGlue})",
                                        AutoSize = true,
                                        Padding = new Padding(10)
                                    };
                                    buttonWithGlue.Click += (s, e) =>
                                    {
                                        dialog.DialogResult = DialogResult.No;
                                        dialog.Close();
                                    };

                                    buttonPanel.Controls.Add(buttonWithoutGlue);
                                    buttonPanel.Controls.Add(buttonWithGlue);

                                    contentPanel.Controls.Add(buttonPanel);
                                    dialog.Controls.Add(contentPanel);

                                    DialogResult result = dialog.ShowDialog();

                                    if (result == DialogResult.Yes)
                                    {
                                        return priceWithoutGlue;
                                    }
                                    else if (result == DialogResult.No)
                                    {
                                        return priceWithGlue;
                                    }
                                    else
                                    {
                                        throw new Exception("Пользователь не выбрал ставку.");
                                    }
                                }
                            }
                            else if (priceWithoutGlue > 0)
                            {
                                return priceWithoutGlue;
                            }
                            else if (priceWithGlue > 0)
                            {
                                return priceWithGlue;
                            }
                            else
                            {
                                throw new Exception($"Неверный формат цен для адреса \"{address}\" в строке {row}.");
                            }
                        }
                    }
                }

                throw new Exception($"Адрес \"{address}\" не найден в файле расценок.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 0;
            }
        }



        private string FindAbbreviationInPriceList(string address)
        {
            try
            {
                if (string.IsNullOrEmpty(priceListFilePath))
                {
                    throw new Exception("Файл расценок не выбран. Укажите его с помощью кнопки выбора расценок.");
                }

                using (ExcelPackage package = new ExcelPackage(new FileInfo(priceListFilePath)))
                {
                    var priceSheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (priceSheet == null || priceSheet.Dimension == null)
                    {
                        throw new Exception("Файл расценок пуст или не содержит данных.");
                    }

                    int startRow = priceSheet.Dimension.Start.Row;
                    int endRow = priceSheet.Dimension.End.Row;

                    for (int row = startRow; row <= endRow; row++)
                    {
                        // Читаем адрес из колонки D
                        string cellAddress = priceSheet.Cells[row, 4]?.Text?.Trim(); // Колонка D
                        if (string.Equals(cellAddress, address, StringComparison.OrdinalIgnoreCase))
                        {
                            // Читаем сокращение из колонки C
                            string abbreviation = priceSheet.Cells[row, 3]?.Text?.Trim(); // Колонка C
                            if (!string.IsNullOrEmpty(abbreviation))
                            {
                                return abbreviation;
                            }
                            else
                            {
                                throw new Exception($"Сокращение не найдено в строке {row} для адреса \"{address}\".");
                            }
                        }
                    }
                }

                throw new Exception($"Адрес \"{address}\" не найден в файле расценок.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при поиске сокращения: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return string.Empty;
            }
        }

        private string FindIGKInPriceList(string address)
        {
            try
            {
                if (string.IsNullOrEmpty(priceListFilePath))
                {
                    throw new Exception("Файл расценок не выбран. Укажите его с помощью кнопки выбора расценок.");
                }

                using (ExcelPackage package = new ExcelPackage(new FileInfo(priceListFilePath)))
                {
                    var priceSheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (priceSheet == null || priceSheet.Dimension == null)
                    {
                        throw new Exception("Файл расценок пуст или не содержит данных.");
                    }

                    int startRow = priceSheet.Dimension.Start.Row;
                    int endRow = priceSheet.Dimension.End.Row;

                    for (int row = startRow; row <= endRow; row++)
                    {
                        // Читаем адрес из колонки D
                        string cellAddress = priceSheet.Cells[row, 4]?.Text?.Trim(); // Колонка D
                        if (string.Equals(cellAddress, address, StringComparison.OrdinalIgnoreCase))
                        {
                            // Читаем значение из колонки G (ИГК)
                            string igk = priceSheet.Cells[row, 7]?.Text?.Trim(); // Колонка G
                            return igk; // Вернуть значение (может быть пустым или ИГК)
                        }
                    }
                }

                throw new Exception($"Адрес \"{address}\" не найден в файле расценок.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при поиске ИГК: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    throw new Exception("Не удалось открыть указанную таблицу.");
                }

                // Проходим по всем строкам таблицы
                int rowCount = sheet.Dimension.Rows;
                for (int row = 1; row <= rowCount; row++)
                {
                    // Получение значений из колонок B, G и H
                    string columnBValue = sheet.Cells[row, 2]?.Value?.ToString();
                    string columnGValue = sheet.Cells[row, 7]?.Value?.ToString(); // Колонка G (7-я)
                    string columnHValue = sheet.Cells[row, 8]?.Value?.ToString(); // Колонка H (8-я)

                    // Сравнение значений в колонках B и H с искомыми
                    if (columnHValue == igk && columnBValue == productName)
                    {
                        return columnGValue; // Возвращаем значение из колонки G
                    }
                }
            }

            throw new Exception($"Не удалось найти строчку с ИГК '{igk}' и названием продукта '{productName}'.");
        }


        public static string ConvertSumToWords(double sum)
        {
            // Округляем до целых рублей
            int rubles = (int)Math.Floor(sum);

            // Используем метод Format из NickBuhro.NumToWords
            string rublesText = RussianConverter.Format(rubles) + " рублей";

            // Добавляем фиксированное значение копеек
            string kopecksText = "00 копеек";

            // Итоговая строка
            string result = $"{rublesText} {kopecksText}";

            // Преобразуем первую букву в верхний регистр
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
                    MessageBox.Show("Временная таблица пуста. Заполните ее данными перед созданием отчета.",
                                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (string.IsNullOrEmpty(headerFilePath))
                {
                    MessageBox.Show("Не выбран файл для шапки. Укажите его с помощью соответствующей кнопки.",
                                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string tempFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "temp.xlsx");

                // Хранение исключенных заказов
                List<DataGridViewRow> excludedOrders = new List<DataGridViewRow>();
                int currentOrderNumber = 1; // Нумерация заказов

                // Шаг 1: Создание временного файла
                using (ExcelPackage package = new ExcelPackage())
                {
                    var sheet = package.Workbook.Worksheets.Add("Счет");

                    // Устанавливаем общий шрифт Arial
                    sheet.Cells.Style.Font.Name = "Arial";

                    // Заголовок счета
                    sheet.Cells["A8:I8"].Merge = true;
                    sheet.Cells["A8:I8"].Value = $"СЧЕТ №{accnum} от {accdata}г";
                    sheet.Cells["A8:I8"].Style.Font.Size = 14;
                    sheet.Cells["A8:I8"].Style.Font.Bold = true;
                    sheet.Cells["A8:I8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    sheet.Cells["A8:I8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    // Плательщик
                    sheet.Cells["A9:I9"].Merge = true;
                    sheet.Cells["A9:I9"].Value = "Плательщик: ООО «Масикс» ИНН 6164134558";
                    sheet.Cells["A9:I9"].Style.Font.Size = 10;
                    sheet.Cells["A9:I9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                    // Заголовки таблицы
                    string[] headers = { "№", "Наименование товара", "", "Ед. Измерения", "Количество", "Цена", "", "Сумма", "" };
                    for (int i = 0; i < headers.Length; i++)
                    {
                        sheet.Cells[12, i + 1].Value = headers[i];
                        sheet.Cells[12, i + 1].Style.Font.Size = 10;
                        sheet.Cells[12, i + 1].Style.WrapText = true;
                        sheet.Cells[12, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        sheet.Cells[12, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        sheet.Cells[12, i + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    }

                    // Объединяем заголовки
                    sheet.Cells["B12:C12"].Merge = true;
                    sheet.Cells["F12:G12"].Merge = true;
                    sheet.Cells["H12:I12"].Merge = true;

                    int currentRow = 13;
                    double totalSum = 0;
                    string currentFooting = GetFootingFromTable(null, currentProductName);

                    foreach (DataGridViewRow gridRow in dataGridView4.Rows)
                    {
                        if (gridRow.Cells[0].Value == null || gridRow.Cells[2].Value == null || gridRow.Cells[3].Value == null)
                            continue; // Пропускаем пустые строки

                        try
                        {
                            string пунктРазгрузки = gridRow.Cells[2].Value.ToString();
                            string датаОтгрузки = gridRow.Cells[1].Value.ToString();
                            int количество = Convert.ToInt32(gridRow.Cells[3].Value);
                            string сокращениеПункта = FindAbbreviationInPriceList(пунктРазгрузки);

                            // Проверяем наличие ИГК
                            string игк = FindIGKInPriceList(пунктРазгрузки);
                            if (!string.IsNullOrEmpty(игк))
                            {
                                excludedOrders.Add(gridRow);
                                continue; // Пропускаем строки с ИГК
                            }

                            double цена = FindPriceInPriceList(пунктРазгрузки);
                            double сумма = количество * цена;
                            totalSum += сумма;

                            // Заполняем строку
                            sheet.Cells[currentRow, 1].Value = currentOrderNumber++;
                            sheet.Cells[currentRow, 2].Value = $"Транспортные услуги ст. Саратовская-{сокращениеПункта} {датаОтгрузки}";
                            sheet.Cells[currentRow, 2, currentRow, 3].Merge = true;
                            sheet.Cells[currentRow, 4].Value = "рейс";
                            sheet.Cells[currentRow, 5].Value = количество;
                            sheet.Cells[currentRow, 6].Value = $"{цена:0.00}";
                            sheet.Cells[currentRow, 6, currentRow, 7].Merge = true;
                            sheet.Cells[currentRow, 8].Value = $"{сумма:0,0.00}";
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
                            MessageBox.Show($"Ошибка при обработке строки: {exRow.Message}", "Ошибка строки",
                                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                    }

                    // Добавление записи в реестр (если включено)
                    if (checkBox1.Checked) // Проверяем, включен ли CheckBox1
                    {
                        FillRegistry(currentProductName, accnum, accdata, totalSum, currentFooting);
                    }
                    // Итоговые строки
                    sheet.Cells[currentRow, 6, currentRow, 7].Merge = true;
                    sheet.Cells[currentRow, 6, currentRow, 7].Value = "Итого:";
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
                    sheet.Cells[currentRow, 6, currentRow, 7].Value = "Без налога(НДС).";
                    sheet.Cells[currentRow, 6, currentRow, 7].Style.Font.Size = 10;
                    sheet.Cells[currentRow, 6, currentRow, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    sheet.Cells[currentRow, 8, currentRow, 9].Merge = true;
                    sheet.Cells[currentRow, 8, currentRow, 9].Style.Font.Size = 10;
                    sheet.Cells[currentRow, 8, currentRow, 9].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    currentRow++;
                    sheet.Cells[currentRow, 6, currentRow, 7].Merge = true;
                    sheet.Cells[currentRow, 6, currentRow, 7].Value = "Всего к оплате:";
                    sheet.Cells[currentRow, 6, currentRow, 7].Style.Font.Size = 10;
                    sheet.Cells[currentRow, 6, currentRow, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    sheet.Cells[currentRow, 8, currentRow, 9].Merge = true;
                    sheet.Cells[currentRow, 8, currentRow, 9].Value = $"{totalSum:0,0.00}";
                    sheet.Cells[currentRow, 8, currentRow, 9].Style.Font.Size = 10;
                    sheet.Cells[currentRow, 8, currentRow, 9].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    sheet.Cells[currentRow, 8, currentRow, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    sheet.Cells[currentRow, 8, currentRow, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    // Общая информация о наименованиях и сумме
                    currentRow++;
                    int totalItems = currentOrderNumber - 1; // Последний номер наименования из колонки A

                    sheet.Cells[currentRow, 1].Merge = false;
                    sheet.Cells[currentRow, 1].Value = $"Всего наименований {totalItems}, на сумму {Math.Floor(totalSum):N0}-{(totalSum % 1 * 100):00}";
                    sheet.Cells[currentRow, 1].Style.Font.Size = 10;
                    sheet.Cells[currentRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                    // Перевод суммы в текст
                    currentRow++;
                    sheet.Cells[currentRow, 1].Merge = false;
                    sheet.Cells[currentRow, 1].Value = ConvertSumToWords(totalSum);
                    sheet.Cells[currentRow, 1].Style.Font.Size = 10;
                    sheet.Cells[currentRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                    // Отступ вниз на 3 строки
                    currentRow += 4;

                    // Добавляем строки с подписями
                    sheet.Cells[currentRow, 1].Value = $"Руководитель предприятия____________________({currentProductName})";
                    sheet.Cells[currentRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    currentRow += 2;

                    sheet.Cells[currentRow, 1].Value = $"Главный бухгалтер __________________________ ({currentProductName})";
                    sheet.Cells[currentRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                    // Настройка шрифта и размера текста для подписей
                    sheet.Cells[currentRow - 2, 1, currentRow, 1].Style.Font.Size = 10;

                    // Сохранение временного файла
                    package.SaveAs(new FileInfo(tempFilePath));
                }

                // Шаг 2: Объединение временного файла с файлом шапки
                using (ExcelPackage headerPackage = new ExcelPackage(new FileInfo(headerFilePath)))
                using (ExcelPackage tempPackage = new ExcelPackage(new FileInfo(tempFilePath)))
                {
                    var headerSheet = headerPackage.Workbook.Worksheets.FirstOrDefault();
                    var tempSheet = tempPackage.Workbook.Worksheets.FirstOrDefault();

                    if (headerSheet == null || tempSheet == null)
                    {
                        throw new Exception("Ошибка чтения данных из временного файла или файла шапки.");
                    }

                    using (ExcelPackage mergedPackage = new ExcelPackage())
                    {
                        var mergedSheet = mergedPackage.Workbook.Worksheets.Add("Итоговый отчет");

                        // Копируем шапку
                        var headerRange = headerSheet.Dimension;
                        headerSheet.Cells[headerRange.Start.Row, headerRange.Start.Column, headerRange.End.Row, headerRange.End.Column]
                            .Copy(mergedSheet.Cells[headerRange.Start.Row, headerRange.Start.Column]);

                        // Копируем временную таблицу под шапкой
                        var tempRange = tempSheet.Dimension;
                        int offsetRow = headerRange?.End.Row + 1 ?? 1;
                        tempSheet.Cells[tempRange.Start.Row, tempRange.Start.Column, tempRange.End.Row, tempRange.End.Column]
                            .Copy(mergedSheet.Cells[offsetRow, tempRange.Start.Column]);

                        // Вывод информации об исключенных адресах
                        if (excludedOrders.Count > 0)
                        {
                            string excludedInfo = "Следующие адреса были исключены из основного счета из-за наличия ИГК, для них будет создан отдельный счет:\n";
                            foreach (var row in excludedOrders)
                            {
                                string пунктРазгрузки = row.Cells[2].Value.ToString();
                                excludedInfo += $"- {пунктРазгрузки}\n";
                            }
                            MessageBox.Show(excludedInfo, "Исключенные адреса", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }

                        // Сохраняем итоговый файл
                        SaveFileDialog saveFileDialog = new SaveFileDialog
                        {
                            Filter = "Excel Files|*.xlsx",
                            Title = "Сохранить итоговый файл"
                        };

                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            mergedPackage.SaveAs(new FileInfo(saveFileDialog.FileName));
                            MessageBox.Show("Файл успешно создан!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }

                // Удаляем временный файл
                if (File.Exists(tempFilePath))
                {
                    File.Delete(tempFilePath);
                }

                // Новый этап: обработка excludedOrders
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



                    // Показ окна для изменения номера и даты
                    using (Form dialog = new Form())
                    {
                        dialog.AutoSize = true;
                        dialog.AutoSizeMode = AutoSizeMode.GrowAndShrink;
                        dialog.Text = "Обработка заказа с ИГК";
                        dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                        dialog.StartPosition = FormStartPosition.CenterScreen;

                        // Основной контейнер
                        TableLayoutPanel layoutPanel = new TableLayoutPanel
                        {
                            Dock = DockStyle.Fill,
                            AutoSize = true,
                            AutoSizeMode = AutoSizeMode.GrowAndShrink,
                            ColumnCount = 2,
                            Padding = new Padding(10)
                        };
                        layoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
                        layoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

                        // Метка с информацией об обрабатываемом заказе
                        Label labelOrderInfo = new Label
                        {
                            Text = $"Перевозчик: {currentProductName}\nАдрес: {address}\nДата заказа: {date}\nКоличество рейсов: {quantity}",
                            Font = new Font("Arial", 12, FontStyle.Bold),
                            Dock = DockStyle.Fill,
                            AutoSize = true,
                            TextAlign = ContentAlignment.MiddleLeft
                        };

                        // Метка и TextBox для нового номера
                        Label labelNumber = new Label
                        {
                            Text = "Счет №:",
                            Font = new Font("Arial", 10),
                            TextAlign = ContentAlignment.MiddleLeft,
                            Anchor = AnchorStyles.Left
                        };
                        TextBox textBoxNumber = new TextBox
                        {
                            Width = 225,
                            Anchor = AnchorStyles.Left
                        };

                        // Метка и TextBox для новой даты
                        Label labelDate = new Label
                        {
                            Text = "От:",
                            Font = new Font("Arial", 10),
                            TextAlign = ContentAlignment.MiddleLeft,
                            Anchor = AnchorStyles.Left
                        };
                        TextBox textBoxDate = new TextBox
                        {
                            Width = 225,
                            Anchor = AnchorStyles.Left
                        };

                        // Кнопка "OK"
                        Button buttonOk = new Button
                        {
                            Text = "OK",
                            DialogResult = DialogResult.OK,
                            Font = new Font("Arial", 10, FontStyle.Bold),
                            AutoSize = true,
                            Anchor = AnchorStyles.None
                        };

                        // Добавление элементов в таблицу
                        layoutPanel.Controls.Add(labelOrderInfo, 0, 0);
                        layoutPanel.SetColumnSpan(labelOrderInfo, 2); // Распространение на 2 столбца

                        layoutPanel.Controls.Add(labelNumber, 0, 1);
                        layoutPanel.Controls.Add(textBoxNumber, 1, 1);

                        layoutPanel.Controls.Add(labelDate, 0, 2);
                        layoutPanel.Controls.Add(textBoxDate, 1, 2);

                        layoutPanel.Controls.Add(buttonOk, 0, 3);
                        layoutPanel.SetColumnSpan(buttonOk, 2); // Распространение кнопки на 2 столбца
                        layoutPanel.SetCellPosition(buttonOk, new TableLayoutPanelCellPosition(0, 3)); // Центрирование кнопки

                        dialog.Controls.Add(layoutPanel);

                        dialog.AcceptButton = buttonOk;


                        if (dialog.ShowDialog() == DialogResult.OK)
                        {
                            // Получаем новые данные из TextBox
                            string newNumber = textBoxNumber.Text;
                            string newDate = textBoxDate.Text;

                            // Шаг 1: Создание временного файла
                            using (ExcelPackage package = new ExcelPackage())
                            {
                                var sheet = package.Workbook.Worksheets.Add("Счет");

                                // Устанавливаем общий шрифт Arial
                                sheet.Cells.Style.Font.Name = "Arial";

                                // Заголовок счета
                                sheet.Cells["A8:I8"].Merge = true;
                                sheet.Cells["A8:I8"].Value = $"СЧЕТ №{newNumber} от {newDate}г";
                                sheet.Cells["A8:I8"].Style.Font.Size = 14;
                                sheet.Cells["A8:I8"].Style.Font.Bold = true;
                                sheet.Cells["A8:I8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                sheet.Cells["A8:I8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                // Плательщик
                                sheet.Cells["A9:I9"].Merge = true;
                                sheet.Cells["A9:I9"].Value = "Плательщик: ООО «Масикс» ИНН 6164134558";
                                sheet.Cells["A9:I9"].Style.Font.Size = 10;
                                sheet.Cells["A9:I9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                                // Основание
                                sheet.Cells["A10:I10"].Merge = true;
                                sheet.Cells["A10:I10"].Value = currentFooting;
                                sheet.Cells["A10:I10"].Style.Font.Size = 10;
                                sheet.Cells["A10:I10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                                // ИГК
                                sheet.Cells["A11:I11"].Merge = true;
                                sheet.Cells["A11:I11"].Value = currentIGK;
                                sheet.Cells["A11:I11"].Style.Font.Size = 10;
                                sheet.Cells["A11:I11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                                // Заголовки таблицы
                                string[] headers = { "№", "Наименование товара", "", "Ед. Измерения", "Количество", "Цена", "", "Сумма", "" };
                                for (int i = 0; i < headers.Length; i++)
                                {
                                    sheet.Cells[12, i + 1].Value = headers[i];
                                    sheet.Cells[12, i + 1].Style.Font.Size = 10;
                                    sheet.Cells[12, i + 1].Style.WrapText = true;
                                    sheet.Cells[12, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    sheet.Cells[12, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    sheet.Cells[12, i + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                }

                                // Объединяем заголовки
                                sheet.Cells["B12:C12"].Merge = true;
                                sheet.Cells["F12:G12"].Merge = true;
                                sheet.Cells["H12:I12"].Merge = true;

                                int IGKcurrentRow = 13;
                                double totalSum = 0;

                                try
                                {
                                    string сокращениеПункта = FindAbbreviationInPriceList(address);
                                    double цена = FindPriceInPriceList(address);

                                    double сумма = quantity * цена;
                                    totalSum += сумма;

                                    // Добавление записи в реестр (если включено)
                                    if (checkBox1.Checked) // Проверяем, включен ли CheckBox1
                                    {
                                        FillRegistry(currentProductName, newNumber, newDate, totalSum, currentFooting);
                                    }

                                    // Заполняем строку
                                    sheet.Cells[IGKcurrentRow, 1].Value = 1;
                                    sheet.Cells[IGKcurrentRow, 2].Value = $"Транспортные услуги ст. Саратовская-{сокращениеПункта} {date}";
                                    sheet.Cells[IGKcurrentRow, 2, IGKcurrentRow, 3].Merge = true;
                                    sheet.Cells[IGKcurrentRow, 4].Value = "рейс";
                                    sheet.Cells[IGKcurrentRow, 5].Value = quantity;
                                    sheet.Cells[IGKcurrentRow, 6].Value = $"{цена:0.00}";
                                    sheet.Cells[IGKcurrentRow, 6, IGKcurrentRow, 7].Merge = true;
                                    sheet.Cells[IGKcurrentRow, 8].Value = $"{сумма:0,0.00}";
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
                                    MessageBox.Show($"Ошибка при обработке строки: {exRow.Message}", "Ошибка строки",
                                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }


                                // Итоговые строки
                                sheet.Cells[IGKcurrentRow, 6, IGKcurrentRow, 7].Merge = true;
                                sheet.Cells[IGKcurrentRow, 6, IGKcurrentRow, 7].Value = "Итого:";
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
                                sheet.Cells[IGKcurrentRow, 6, IGKcurrentRow, 7].Value = "Без налога(НДС).";
                                sheet.Cells[IGKcurrentRow, 6, IGKcurrentRow, 7].Style.Font.Size = 10;
                                sheet.Cells[IGKcurrentRow, 6, IGKcurrentRow, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                                sheet.Cells[IGKcurrentRow, 8, IGKcurrentRow, 9].Merge = true;
                                sheet.Cells[IGKcurrentRow, 8, IGKcurrentRow, 9].Style.Font.Size = 10;
                                sheet.Cells[IGKcurrentRow, 8, IGKcurrentRow, 9].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                                IGKcurrentRow++;
                                sheet.Cells[IGKcurrentRow, 6, IGKcurrentRow, 7].Merge = true;
                                sheet.Cells[IGKcurrentRow, 6, IGKcurrentRow, 7].Value = "Всего к оплате:";
                                sheet.Cells[IGKcurrentRow, 6, IGKcurrentRow, 7].Style.Font.Size = 10;
                                sheet.Cells[IGKcurrentRow, 6, IGKcurrentRow, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                                sheet.Cells[IGKcurrentRow, 8, IGKcurrentRow, 9].Merge = true;
                                sheet.Cells[IGKcurrentRow, 8, IGKcurrentRow, 9].Value = $"{totalSum:0,0.00}";
                                sheet.Cells[IGKcurrentRow, 8, IGKcurrentRow, 9].Style.Font.Size = 10;
                                sheet.Cells[IGKcurrentRow, 8, IGKcurrentRow, 9].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                sheet.Cells[IGKcurrentRow, 8, IGKcurrentRow, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                sheet.Cells[IGKcurrentRow, 8, IGKcurrentRow, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                // Общая информация о наименованиях и сумме
                                IGKcurrentRow++;
                                int totalItems = currentOrderNumber - 1; // Последний номер наименования из колонки A

                                sheet.Cells[IGKcurrentRow, 1].Merge = false;
                                sheet.Cells[IGKcurrentRow, 1].Value = $"Всего наименований {totalItems}, на сумму {Math.Floor(totalSum):N0}-{(totalSum % 1 * 100):00}";
                                sheet.Cells[IGKcurrentRow, 1].Style.Font.Size = 10;
                                sheet.Cells[IGKcurrentRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                                // Перевод суммы в текст
                                IGKcurrentRow++;
                                sheet.Cells[IGKcurrentRow, 1].Merge = false;
                                sheet.Cells[IGKcurrentRow, 1].Value = ConvertSumToWords(totalSum);
                                sheet.Cells[IGKcurrentRow, 1].Style.Font.Size = 10;
                                sheet.Cells[IGKcurrentRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                                // Отступ вниз на 3 строки
                                IGKcurrentRow += 4;

                                // Добавляем строки с подписями
                                sheet.Cells[IGKcurrentRow, 1].Value = $"Руководитель предприятия____________________({currentProductName})";
                                sheet.Cells[IGKcurrentRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                IGKcurrentRow += 2;

                                sheet.Cells[IGKcurrentRow, 1].Value = $"Главный бухгалтер __________________________ ({currentProductName})";
                                sheet.Cells[IGKcurrentRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                                // Настройка шрифта и размера текста для подписей
                                sheet.Cells[IGKcurrentRow - 2, 1, IGKcurrentRow, 1].Style.Font.Size = 10;

                                // Сохранение временного файла
                                package.SaveAs(new FileInfo(tempFilePath));
                            }

                            // Шаг 2: Объединение временного файла с файлом шапки
                            using (ExcelPackage headerPackage = new ExcelPackage(new FileInfo(headerFilePath)))
                            using (ExcelPackage tempPackage = new ExcelPackage(new FileInfo(tempFilePath)))
                            {
                                var headerSheet = headerPackage.Workbook.Worksheets.FirstOrDefault();
                                var tempSheet = tempPackage.Workbook.Worksheets.FirstOrDefault();

                                if (headerSheet == null || tempSheet == null)
                                {
                                    throw new Exception("Ошибка чтения данных из временного файла или файла шапки.");
                                }

                                using (ExcelPackage mergedPackage = new ExcelPackage())
                                {
                                    var mergedSheet = mergedPackage.Workbook.Worksheets.Add("Итоговый отчет");

                                    // Копируем шапку
                                    var headerRange = headerSheet.Dimension;
                                    headerSheet.Cells[headerRange.Start.Row, headerRange.Start.Column, headerRange.End.Row, headerRange.End.Column]
                                        .Copy(mergedSheet.Cells[headerRange.Start.Row, headerRange.Start.Column]);

                                    // Копируем временную таблицу под шапкой
                                    var tempRange = tempSheet.Dimension;
                                    int offsetRow = headerRange?.End.Row + 1 ?? 1;
                                    tempSheet.Cells[tempRange.Start.Row, tempRange.Start.Column, tempRange.End.Row, tempRange.End.Column]
                                        .Copy(mergedSheet.Cells[offsetRow, tempRange.Start.Column]);

                                    // Сохраняем итоговый файл
                                    SaveFileDialog saveFileDialog = new SaveFileDialog
                                    {
                                        Filter = "Excel Files|*.xlsx",
                                        Title = "Сохранить итоговый файл"
                                    };

                                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                                    {
                                        mergedPackage.SaveAs(new FileInfo(saveFileDialog.FileName));
                                        MessageBox.Show("Файл успешно создан!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    }
                                }
                            }

                            if (File.Exists(tempFilePath))
                            {
                                File.Delete(tempFilePath);
                            }
                        }
                    }

                    excludedOrders.RemoveAt(0);
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }



        }

        private void button11_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                Title = "Выберите файл таблицы реестра"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                registryFilePath = openFileDialog.FileName;
                label11.Text = $"Статус: Подключено ({Path.GetFileName(registryFilePath)})";
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            registryFilePath = null;
            label11.Text = "Статус: ";
        }

        private void FillRegistry(string currentProductName, string accnum, string accdata, double totalsum, string currentFooting)
        {
            try
            {
                if (string.IsNullOrEmpty(registryFilePath))
                {
                    throw new Exception("Файл реестра не выбран. Укажите его с помощью соответствующей кнопки.");
                }

                using (ExcelPackage package = new ExcelPackage(new FileInfo(registryFilePath)))
                {
                    var sheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (sheet == null)
                    {
                        sheet = package.Workbook.Worksheets.Add("Реестр");
                    }

                    // Проверка и добавление заголовков
                    if (string.IsNullOrWhiteSpace(sheet.Cells[3, 1]?.Text)) // Проверяем ячейку А3
                    {
                        sheet.Cells[3, 1].Value = "№";
                        sheet.Cells[3, 2].Value = "Наименование контрагента";
                        sheet.Cells[3, 3].Value = "Номер документа";
                        sheet.Cells[3, 4].Value = "Сумма";
                        sheet.Cells[3, 5].Value = "Общая сумма по получателю";
                        sheet.Cells[3, 6].Value = "Наименование";
                        sheet.Cells[3, 7].Value = "№ договора";

                        // Применяем стиль для заголовков
                        using (var range = sheet.Cells[3, 1, 3, 7])
                        {
                            range.Style.Font.Bold = true;
                            range.Style.Font.Name = "Times New Roman";
                            range.Style.Font.Size = 11;
                            range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            range.Style.WrapText = true;
                            range.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin); // Границы вокруг ячейки
                        }
                    }

                    // Поиск последней заполненной строки
                    int startRow = 4; // Данные начинаются с 4-й строки
                    int currentRow = startRow;
                    while (!string.IsNullOrWhiteSpace(sheet.Cells[currentRow, 1]?.Text) ||
                           !string.IsNullOrWhiteSpace(sheet.Cells[currentRow, 2]?.Text))
                    {
                        currentRow++;
                    }

                    // Заполнение новой строки
                    sheet.Cells[currentRow, 1].Value = currentRow - startRow + 1; // Номер строки
                    sheet.Cells[currentRow, 2].Value = currentProductName;
                    sheet.Cells[currentRow, 3].Value = $"№ {accnum} от {accdata}";
                    sheet.Cells[currentRow, 4].Value = totalsum;
                    sheet.Cells[currentRow, 5].Value = totalsum;
                    sheet.Cells[currentRow, 6].Value = "перевозка гб 4.3";
                    sheet.Cells[currentRow, 7].Value = currentFooting;

                    // Применяем стиль для строки данных
                    using (var range = sheet.Cells[currentRow, 1, currentRow, 7])
                    {
                        range.Style.Font.Name = "Times New Roman";
                        range.Style.Font.Size = 11;
                        range.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        range.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        range.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin); // Границы вокруг ячейки
                    }

                    // Сохраняем изменения в файл
                    package.Save();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при заполнении реестра: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }



        private void button4_Click(object sender, EventArgs e)
        {
            Form aboutDialog = new Form
            {
                AutoSize = true, // Автоматическое изменение размера окна
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Text = "О программе",
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                StartPosition = FormStartPosition.CenterScreen
            };

            // Текстовое описание программы
            Label descriptionLabel = new Label
            {
                AutoSize = true,
                Text = "Rcmu Account Manager\n" +
                        "Версия: 1.1 (unlimited)\n" +
                        "Разработчик: Дмитрий Кремов, tg: @ReCream, github: recreamu\n" +
                        "\n" +
                        "Данное приложение позволяет пользователю создавать таблицу счетов\n" +
                        "в формате книги Excel на основе данных из нескольких других выбранных\n" +
                        "таблиц того же формата. Структура данных у входящих и выходящих таблиц\n" +
                        "всегда строгая, за примером обратитесь к предыдущему пользователю или\n" +
                        "разработчику. При создании программы использовались открытые пакеты:\n" +
                        "EPPlus(7.5.2) и NickVuhro.NumToWords(1.1.3), а так же ChatGPT.\n" +
                        "Исходный код есть на моем GitHub, желающим его изучить желаю удачи.\n" +
                        "\n" +
                        "Авторские права защищены GNU GENERAL PUBLIC LICENSE v3 © 2025",
                Location = new Point(20, 20)
            };

            // Кнопка закрытия
            Button closeButton = new Button
            {
                Text = "Закрыть",
                AutoSize = true,
                Anchor = AnchorStyles.None
            };
            closeButton.Click += (s, args) => aboutDialog.Close();

            // Панель размещения с вертикальным выравниванием
            FlowLayoutPanel layoutPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Padding = new Padding(20, 20, 20, 20), // Отступы вокруг содержимого
                BorderStyle = BorderStyle.None
            };

            // Выравнивание по центру
            layoutPanel.Controls.Add(descriptionLabel);
            layoutPanel.SetFlowBreak(descriptionLabel, true); // Прерываем поток после текста

            layoutPanel.Controls.Add(closeButton);
            closeButton.Margin = new Padding(0, 10, 0, 0); // Небольшой отступ сверху

            aboutDialog.Controls.Add(layoutPanel);

            aboutDialog.ShowDialog();
        }


    }
}

