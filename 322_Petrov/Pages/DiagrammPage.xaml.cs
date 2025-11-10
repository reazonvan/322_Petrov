using System;
using System.Windows.Forms.DataVisualization.Charting;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace _322_Petrov.Pages
{
    /// <summary>
    /// Логика взаимодействия для DiagrammPage.xaml
    /// </summary>
    public partial class DiagrammPage : Page
    {
        private Entities _context = new Entities();

        public DiagrammPage()
        {
            InitializeComponent();
            // Инициализация диаграммы
            InitializeChart();

            // Загрузка данных в ComboBox
            LoadComboBoxData();
        }

        private void InitializeChart()
        {
            // Создание области построения диаграммы
            ChartPayments.ChartAreas.Add(new System.Windows.Forms.DataVisualization.Charting.ChartArea("Main"));

            // Добавление набора данных
            var currentSeries = new System.Windows.Forms.DataVisualization.Charting.Series("Платежи")
            {
                IsValueShownAsLabel = true
            };
            ChartPayments.Series.Add(currentSeries);
        }

        private void LoadComboBoxData()
        {
            // Загрузка пользователей
            CmbUser.ItemsSource = _context.Users.ToList();

            // Загрузка типов диаграмм
            CmbDiagram.ItemsSource = Enum.GetValues(typeof(SeriesChartType));

            // Установка значений по умолчанию
            if (CmbUser.Items.Count > 0)
                CmbUser.SelectedIndex = 0;

            if (CmbDiagram.Items.Count > 0)
                CmbDiagram.SelectedIndex = 0;
        }

        // Обработчик обновления диаграммы
        private void UpdateChart(object sender, SelectionChangedEventArgs e)
        {
            if (CmbUser.SelectedItem is User currentUser &&
                CmbDiagram.SelectedItem is SeriesChartType currentType)
            {
                System.Windows.Forms.DataVisualization.Charting.Series currentSeries = ChartPayments.Series.FirstOrDefault();

                if (currentSeries != null)
                {
                    currentSeries.ChartType = currentType;
                    currentSeries.Points.Clear();

                    var categoriesList = _context.Categories.ToList();
                    foreach (var category in categoriesList)
                    {
                        double sum = (double)_context.Payments.ToList()
                            .Where(u => u.User == currentUser && u.Category == category)
                            .Sum(u => u.Price * u.Num);

                        currentSeries.Points.AddXY(category.Name, sum);
                    }
                }
            }
        }

        private void BtnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Получаем список пользователей с сортировкой по ФИО
                var allUsers = _context.Users.ToList().OrderBy(u => u.FIO).ToList();

                // Создаем новую книгу Excel
                var application = new Excel.Application();
                application.SheetsInNewWorkbook = allUsers.Count();
                Excel.Workbook workbook = application.Workbooks.Add(Type.Missing);

                // Переменная для общего итога
                decimal grandTotal = 0;

                // Запускаем цикл по пользователям
                for (int i = 0; i < allUsers.Count(); i++)
                {
                    int startRowIndex = 1;
                    Excel.Worksheet worksheet = application.Worksheets.Item[i + 1];
                    worksheet.Name = allUsers[i].FIO;

                    // Добавляем названия колонок
                    worksheet.Cells[1, 1] = "Дата платежа";
                    worksheet.Cells[1, 2] = "Название";
                    worksheet.Cells[1, 3] = "Стоимость";
                    worksheet.Cells[1, 4] = "Количество";
                    worksheet.Cells[1, 5] = "Сумма";

                    // Форматируем заголовки колонок
                    Excel.Range columnHeaderRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 5]];
                    columnHeaderRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    columnHeaderRange.Font.Bold = true;
                    startRowIndex++;

                    // Группируем платежи по категориям
                    var userCategories = allUsers[i].Payments
                        .OrderBy(u => u.Date)
                        .GroupBy(u => u.Category)
                        .OrderBy(u => u.Key.Name);

                    // Вложенный цикл по категориям платежей
                    foreach (var groupCategory in userCategories)
                    {
                        // Добавляем заголовок категории
                        Excel.Range headerRange = worksheet.Range[worksheet.Cells[1, startRowIndex], worksheet.Cells[5, startRowIndex]];
                        headerRange.Merge();
                        headerRange.Value = groupCategory.Key.Name;
                        headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        headerRange.Font.Italic = true;
                        startRowIndex++;

                        // Вложенный цикл по платежам
                        foreach (var payment in groupCategory)
                        {
                            worksheet.Cells[startRowIndex, 1] = payment.Date.ToString("dd.MM.yyyy");
                            worksheet.Cells[startRowIndex, 2] = payment.Name;
                            worksheet.Cells[startRowIndex, 3] = payment.Price;
                            (worksheet.Cells[startRowIndex, 3] as Excel.Range).NumberFormat = "0.00";
                            worksheet.Cells[startRowIndex, 4] = payment.Num;
                            worksheet.Cells[startRowIndex, 5].Formula = $"=C{startRowIndex}*D{startRowIndex}";
                            (worksheet.Cells[startRowIndex, 5] as Excel.Range).NumberFormat = "0.00";
                            startRowIndex++;
                        }

                        // Добавляем строку "ИТОГО" для категории
                        Excel.Range sumRange = worksheet.Range[worksheet.Cells[1, startRowIndex], worksheet.Cells[4, startRowIndex]];
                        sumRange.Merge();
                        sumRange.Value = "ИТОГО:";
                        sumRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

                        // Рассчитываем сумму для категории
                        int startSumRow = startRowIndex - groupCategory.Count();
                        int endSumRow = startRowIndex - 1;
                        worksheet.Cells[startRowIndex, 5].Formula = $"=SUM(E{startSumRow}:E{endSumRow})";

                        // Добавляем сумму категории к общему итогу
                        decimal categoryTotal = groupCategory.Sum(p => p.Price * p.Num);
                        grandTotal += categoryTotal;

                        sumRange.Font.Bold = true;
                        (worksheet.Cells[startRowIndex, 5] as Excel.Range).Font.Bold = true;
                        startRowIndex++;
                    }

                    // Добавляем границы таблицы
                    if (startRowIndex > 1)
                    {
                        Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[5, startRowIndex - 1]];
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle =
                        Excel.XlLineStyle.xlContinuous;
                    }

                    // Устанавливаем автоширину столбцов
                    worksheet.Columns.AutoFit();
                }

                // Добавляем лист "Общий итог"
                Excel.Worksheet summarySheet = workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
                summarySheet.Name = "Общий итог";

                // Запись заголовка и значения
                summarySheet.Cells[1, 1] = "Общий итог:";
                summarySheet.Cells[1, 2] = grandTotal;
                (summarySheet.Cells[1, 2] as Excel.Range).NumberFormat = "0.00";

                // Форматирование: красный цвет и жирный шрифт
                Excel.Range summaryRange = summarySheet.Range[summarySheet.Cells[1, 1], summarySheet.Cells[1, 2]];
                summaryRange.Font.Color = Excel.XlRgbColor.rgbRed;
                summaryRange.Font.Bold = true;

                // Автоподбор ширины столбцов
                summarySheet.Columns.AutoFit();

                // Разрешаем отображение Excel
                application.Visible = true;

                // Сохраняем файл
                string basePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string excelPath = System.IO.Path.Combine(basePath, "Payments_Report.xlsx");
                workbook.SaveAs(excelPath);

                MessageBox.Show($"Excel файл успешно сохранен:\n{excelPath}", "Экспорт завершен", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Excel: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnExportWord_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Получаем список пользователей и категорий из базы данных
                var allUsers = _context.Users.ToList();
                var allCategories = _context.Categories.ToList();

                // Создаем новый документ Word
                var application = new Word.Application();
                Word.Document document = application.Documents.Add();

                // Запускаем цикл по пользователям
                foreach (var user in allUsers)
                {
                    // Создаем абзац для ФИО пользователя
                    Word.Paragraph userParagraph = document.Paragraphs.Add();
                    Word.Range userRange = userParagraph.Range;
                    userRange.Text = user.FIO;

                    // Пробуем разные варианты названий стилей
                    try
                    {
                        userRange.set_Style("Заголовок");
                    }
                    catch
                    {
                        try
                        {
                            userRange.set_Style("Заголовок 1");
                        }
                        catch
                        {
                            userRange.set_Style("Heading 1");
                        }
                    }

                    userRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    userRange.InsertParagraphAfter();
                    document.Paragraphs.Add(); // Пустая строка

                    // Добавляем таблицу с платежами
                    Word.Paragraph tableParagraph = document.Paragraphs.Add();
                    Word.Range tableRange = tableParagraph.Range;
                    Word.Table paymentsTable = document.Tables.Add(tableRange, allCategories.Count() + 1, 2);

                    // Форматируем таблицу
                    paymentsTable.Borders.InsideLineStyle = paymentsTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    paymentsTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                    // Добавляем названия колонок
                    Word.Range cellRange;

                    cellRange = paymentsTable.Cell(1, 1).Range;
                    cellRange.Text = "Категория";
                    cellRange = paymentsTable.Cell(1, 2).Range;
                    cellRange.Text = "Сумма расходов";

                    // Форматируем заголовки таблицы
                    paymentsTable.Rows[1].Range.Font.Name = "Times New Roman";
                    paymentsTable.Rows[1].Range.Font.Size = 14;
                    paymentsTable.Rows[1].Range.Bold = 1;
                    paymentsTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    // Заполняем таблицу данными
                    for (int i = 0; i < allCategories.Count(); i++)
                    {
                        var currentCategory = allCategories[i];

                        // Название категории
                        cellRange = paymentsTable.Cell(i + 2, 1).Range;
                        cellRange.Text = currentCategory.Name;
                        cellRange.Font.Name = "Times New Roman";
                        cellRange.Font.Size = 12;

                        // Сумма расходов
                        cellRange = paymentsTable.Cell(i + 2, 2).Range;
                        decimal sum = user.Payments.ToList()
                            .Where(u => u.Category == currentCategory)
                            .Sum(u => u.Num * u.Price);
                        cellRange.Text = sum.ToString("N2") + " руб.";
                        cellRange.Font.Name = "Times New Roman";
                        cellRange.Font.Size = 12;
                    }

                    document.Paragraphs.Add(); // Пустая строка

                    // Добавляем информацию о максимальном платеже
                    Payment maxPayment = user.Payments.OrderByDescending(u => u.Price * u.Num).FirstOrDefault();
                    if (maxPayment != null)
                    {
                        Word.Paragraph maxPaymentParagraph = document.Paragraphs.Add();
                        Word.Range maxPaymentRange = maxPaymentParagraph.Range;
                        maxPaymentRange.Text = $"Самый дорогостоящий платеж - {maxPayment.Name} за {(maxPayment.Price * maxPayment.Num).ToString("N2")} руб. от {maxPayment.Date.ToString("dd.MM.yyyy")}";

                        // Пробуем разные варианты названий стилей
                        try
                        {
                            maxPaymentRange.set_Style("Подзаголовок");
                        }
                        catch
                        {
                            try
                            {
                                maxPaymentRange.set_Style("Заголовок 2");
                            }
                            catch
                            {
                                maxPaymentRange.set_Style("Heading 2");
                            }
                        }

                        maxPaymentRange.Font.Color = Word.WdColor.wdColorDarkRed;
                        maxPaymentRange.InsertParagraphAfter();
                    }

                    // Добавляем информацию о минимальном платеже
                    Payment minPayment = user.Payments.OrderBy(u => u.Price * u.Num).FirstOrDefault();
                    if (minPayment != null)
                    {
                        Word.Paragraph minPaymentParagraph = document.Paragraphs.Add();
                        Word.Range minPaymentRange = minPaymentParagraph.Range;
                        minPaymentRange.Text = $"Самый дешевый платеж - {minPayment.Name} за {(minPayment.Price * minPayment.Num).ToString("N2")} руб. от {minPayment.Date.ToString("dd.MM.yyyy")}";

                        // Пробуем разные варианты названий стилей
                        try
                        {
                            minPaymentRange.set_Style("Подзаголовок");
                        }
                        catch
                        {
                            try
                            {
                                minPaymentRange.set_Style("Заголовок 2");
                            }
                            catch
                            {
                                minPaymentRange.set_Style("Heading 2");
                            }
                        }

                        minPaymentRange.Font.Color = Word.WdColor.wdColorDarkGreen;
                        minPaymentRange.InsertParagraphAfter();
                    }

                    document.Paragraphs.Add(); // Пустая строка

                    // Добавляем разрыв страницы (кроме последнего пользователя)
                    if (user != allUsers.LastOrDefault())
                        document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
                }

                // Разрешаем отображение документа
                application.Visible = true;

                // Сохраняем документ
                string basePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string docxPath = System.IO.Path.Combine(basePath, "Payments.docx");
                string pdfPath = System.IO.Path.Combine(basePath, "Payments.pdf");

                document.SaveAs2(docxPath);
                document.SaveAs2(pdfPath, Word.WdExportFormat.wdExportFormatPDF);

                MessageBox.Show($"Документы успешно сохранены:\n{docxPath}\n{pdfPath}", "Экспорт завершен", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Word: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            try
            {
                // Получаем список пользователей и категорий из базы данных
                var allUsers = _context.Users.ToList();
                var allCategories = _context.Categories.ToList();

                // Создаем новый документ Word
                var application = new Word.Application();
                Word.Document document = application.Documents.Add();

                // ДОБАВЛЯЕМ КОЛОНТИТУЛЫ ПЕРЕД ОСНОВНЫМ СОДЕРЖАНИЕМ

                // Добавляем верхний колонтитул с текущей датой
                foreach (Word.Section section in document.Sections)
                {
                    // Получаем диапазон верхнего колонтитула
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

                    // Устанавливаем форматирование верхнего колонтитула
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    headerRange.Font.ColorIndex = Word.WdColorIndex.wdBlack;
                    headerRange.Font.Size = 10;
                    headerRange.Text = $"Отчет создан: {DateTime.Now.ToString("dd/MM/yyyy")}";
                }

                // Добавляем нижний колонтитул с номерами страниц
                foreach (Word.Section section in document.Sections)
                {
                    // Получаем нижний колонтитул
                    Word.HeaderFooter footer = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];

                    // Добавляем номер страницы в центр нижнего колонтитула
                    footer.PageNumbers.Add(Word.WdPageNumberAlignment.wdAlignPageNumberCenter);
                }

                // Запускаем цикл по пользователям
                foreach (var user in allUsers)
                {
                    // Создаем абзац для ФИО пользователя
                    Word.Paragraph userParagraph = document.Paragraphs.Add();
                    Word.Range userRange = userParagraph.Range;
                    userRange.Text = user.FIO;

                    // Пробуем разные варианты названий стилей
                    try
                    {
                        userRange.set_Style("Заголовок");
                    }
                    catch
                    {
                        try
                        {
                            userRange.set_Style("Заголовок 1");
                        }
                        catch
                        {
                            userRange.set_Style("Heading 1");
                        }
                    }

                    userRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    userRange.InsertParagraphAfter();
                    document.Paragraphs.Add(); // Пустая строка

                    // Добавляем таблицу с платежами
                    Word.Paragraph tableParagraph = document.Paragraphs.Add();
                    Word.Range tableRange = tableParagraph.Range;
                    Word.Table paymentsTable = document.Tables.Add(tableRange, allCategories.Count() + 1, 2);

                    // Форматируем таблицу
                    paymentsTable.Borders.InsideLineStyle = paymentsTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    paymentsTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                    // Добавляем названия колонок
                    Word.Range cellRange;

                    cellRange = paymentsTable.Cell(1, 1).Range;
                    cellRange.Text = "Категория";
                    cellRange = paymentsTable.Cell(1, 2).Range;
                    cellRange.Text = "Сумма расходов";

                    // Форматируем заголовки таблицы
                    paymentsTable.Rows[1].Range.Font.Name = "Times New Roman";
                    paymentsTable.Rows[1].Range.Font.Size = 14;
                    paymentsTable.Rows[1].Range.Bold = 1;
                    paymentsTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    // Заполняем таблицу данными
                    for (int i = 0; i < allCategories.Count(); i++)
                    {
                        var currentCategory = allCategories[i];

                        // Название категории
                        cellRange = paymentsTable.Cell(i + 2, 1).Range;
                        cellRange.Text = currentCategory.Name;
                        cellRange.Font.Name = "Times New Roman";
                        cellRange.Font.Size = 12;

                        // Сумма расходов
                        cellRange = paymentsTable.Cell(i + 2, 2).Range;
                        decimal sum = user.Payments.ToList()
                            .Where(u => u.Category == currentCategory)
                            .Sum(u => u.Num * u.Price);
                        cellRange.Text = sum.ToString("N2") + " руб.";
                        cellRange.Font.Name = "Times New Roman";
                        cellRange.Font.Size = 12;
                    }

                    document.Paragraphs.Add(); // Пустая строка

                    // Добавляем информацию о максимальном платеже
                    Payment maxPayment = user.Payments.OrderByDescending(u => u.Price * u.Num).FirstOrDefault();
                    if (maxPayment != null)
                    {
                        Word.Paragraph maxPaymentParagraph = document.Paragraphs.Add();
                        Word.Range maxPaymentRange = maxPaymentParagraph.Range;
                        maxPaymentRange.Text = $"Самый дорогостоящий платеж - {maxPayment.Name} за {(maxPayment.Price * maxPayment.Num).ToString("N2")} руб. от {maxPayment.Date.ToString("dd.MM.yyyy")}";

                        // Пробуем разные варианты названий стилей
                        try
                        {
                            maxPaymentRange.set_Style("Подзаголовок");
                        }
                        catch
                        {
                            try
                            {
                                maxPaymentRange.set_Style("Заголовок 2");
                            }
                            catch
                            {
                                maxPaymentRange.set_Style("Heading 2");
                            }
                        }

                        maxPaymentRange.Font.Color = Word.WdColor.wdColorDarkRed;
                        maxPaymentRange.InsertParagraphAfter();
                    }

                    // Добавляем информацию о минимальном платеже
                    Payment minPayment = user.Payments.OrderBy(u => u.Price * u.Num).FirstOrDefault();
                    if (minPayment != null)
                    {
                        Word.Paragraph minPaymentParagraph = document.Paragraphs.Add();
                        Word.Range minPaymentRange = minPaymentParagraph.Range;
                        minPaymentRange.Text = $"Самый дешевый платеж - {minPayment.Name} за {(minPayment.Price * minPayment.Num).ToString("N2")} руб. от {minPayment.Date.ToString("dd.MM.yyyy")}";

                        // Пробуем разные варианты названий стилей
                        try
                        {
                            minPaymentRange.set_Style("Подзаголовок");
                        }
                        catch
                        {
                            try
                            {
                                minPaymentRange.set_Style("Заголовок 2");
                            }
                            catch
                            {
                                minPaymentRange.set_Style("Heading 2");
                            }
                        }

                        minPaymentRange.Font.Color = Word.WdColor.wdColorDarkGreen;
                        minPaymentRange.InsertParagraphAfter();
                    }

                    document.Paragraphs.Add(); // Пустая строка

                    // Добавляем разрыв страницы (кроме последнего пользователя)
                    if (user != allUsers.LastOrDefault())
                        document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
                }

                // Разрешаем отображение документа
                application.Visible = true;

                // Сохраняем документ
                string basePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string docxPath = System.IO.Path.Combine(basePath, "Payments_Report.docx");
                string pdfPath = System.IO.Path.Combine(basePath, "Payments_Report.pdf");

                document.SaveAs2(docxPath);
                document.SaveAs2(pdfPath, Word.WdExportFormat.wdExportFormatPDF);

                MessageBox.Show($"Отчет успешно сохранен:\n{docxPath}\n{pdfPath}", "Экспорт завершен", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Word: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

    }
}
