using System;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;



namespace ExcelToWord
{
    public partial class iron66 : Form
    {
        public string SelectedFilePath;
        public string City;
        public string DeliveryLetter;
        public string FileName;
        private List<string> SelectedFilePaths = new List<string>();

        public iron66()
        {
            InitializeComponent();
            this.processBtn.Click += new System.EventHandler(async (sender, e) => await processBtn_ClickAsync(sender, e));
        }

        // Извлечение номера заказа
        private string ExtractOrderId(string text)
        {
            Match match = Regex.Match(text, @"№ (\d+)");
            return match.Success ? match.Groups[1].Value : null;
        }

        // Извлечение буквы, соответствующей типу доставки
        private string ExtractDeliveryLetter(Excel.Worksheet sheet)
        {
            int startRow = 11;
            int currentRow = startRow;
            dynamic productTextCell = sheet.Cells[currentRow, 7].Value;
            while (productTextCell != null)
            {
                currentRow++;
                productTextCell = sheet.Cells[currentRow, 7].Value;
            }

            currentRow++;
            productTextCell = sheet.Cells[currentRow, 2].Value;
            string letterText = (productTextCell != null) ? productTextCell.ToString() : string.Empty;

            for (int i = 0; i < 11; i++)
            {
                if (letterText != "")
                {
                    if (letterText.Contains("ТК") || letterText.Contains("тк") || letterText.Contains("Тк") ||
                        letterText.Contains("Самовывоз") || letterText.Contains("самовывоз") || letterText.Contains("САМОВЫВОЗ"))
                    {
                        DeliveryLetter = "Т";
                        return DeliveryLetter;
                    }
                }

                currentRow++;
                productTextCell = sheet.Cells[currentRow, 2].Value;
                letterText = (productTextCell != null) ? productTextCell.ToString() : string.Empty;
            }

            DeliveryLetter = "Н/Д";
            return DeliveryLetter;
        }


        // Извлечение информации о доставке
        private string ExtractDeliveryInfo(string text, dynamic deliveryInfo, Excel.Worksheet sheet)
        {
            string cityText = deliveryInfo;
            if (DeliveryLetter != null)
            {
                return DeliveryLetter;
            }

            if (cityText.Contains("Бравиум") || cityText.Contains("БРАВИУМ") ||
            cityText.Contains("Меделия") || cityText.Contains("МЕДЕЛИЯ") ||
            cityText.Contains("ДСД Проект") || cityText.Contains("Зерц") ||
            cityText.Contains("Бурдонов"))
            {
                City = "Москва";
                DeliveryLetter = "M";
                return DeliveryLetter;
            }

            else return ExtractDeliveryLetter(sheet);
        }

        // Извлечение информации о заказчике
        private string ExtractCustomer(string text)
        {
            return text.Split(',')[0].Trim();
        }

        // Чтение списка городов
        private List<string> LoadCitiesFromFile(string filePath)
        {
            List<string> cities = new List<string>();

            try
            {
                // Чтение всех строк из файла и добавление их в список городов
                cities = File.ReadAllLines(filePath, Encoding.UTF8).ToList();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при чтении файла городов: {ex.Message}");
            }

            return cities;
        }

        // Извлечение информации о городе
        private string ExtractCityInfo(string text)
        {
            if (City == "Москва")
            {
                return "Москва";
            }

            string citiesFilePath = "City.txt"; // Путь к файлу с городами
            List<string> knownCities = LoadCitiesFromFile(citiesFilePath);

            string[] words = text.Split(new char[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries);

            // Отсекаем первые 6 и последние 3 слова
            words = words.Skip(6).Take(words.Length - 9).ToArray();

            string shortenedText = string.Join(" ", words);

            foreach (string knownCity in knownCities)
            {
                if (shortenedText.Contains(knownCity))
                {
                    return knownCity;
                }
            }

            return null;
        }

        // Извлечение информации о товарах
        private Dictionary<string, string> ExtractProductInfo(string productText)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            Match match = Regex.Match(productText, @"\((.*?)\)");

            if (match.Success)
            {
                string dimensions = match.Groups[1].Value.Trim();

                if (dimensions.Contains("*"))
                {
                    string info = Regex.Replace(productText, $@"\({Regex.Escape(dimensions)}\)", "").Trim();

                    Match colorMatch = Regex.Match(info, @"\((Корпус.*?)\)");
                    string colorInfo = colorMatch.Success ? colorMatch.Groups[1].Value.Trim() : "";

                    info = Regex.Replace(info, $@"\({Regex.Escape(colorInfo)}\)", "").Trim();
                    info = Regex.Replace(info, @"\s+", " ").Trim();

                    result["info"] = info;
                    result["dimensions"] = dimensions;
                    result["color"] = colorInfo;
                }
                else
                {
                    result["info"] = Regex.Replace(productText, @"\s+", " ").Trim();
                    result["dimensions"] = "";
                    result["color"] = "";
                }
            }
            else
            {
                result["info"] = Regex.Replace(productText, @"\s+", " ").Trim();
                result["dimensions"] = "";
                result["color"] = "";
            }

            return result;
        }

        // Извлечение информации о количестве продукта
        private int FindQuantityColumn(Excel.Worksheet sheet)
        {
            foreach (Excel.Range cell in sheet.Rows[9].Cells)
            {
                if (cell.Value != null && cell.Value.ToString().Contains("Кол-во"))
                {
                    return cell.Column;
                }
            }

            return -1;
        }

        private void ReadExcelAndCreateWordDocument(string excelFilePath)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
            Excel.Worksheet sheet = workbook.ActiveSheet;

            // Создание документа Word
            Word.Application wordApp = new Word.Application();
            Word.Document document = wordApp.Documents.Add();
            Word.Style style = document.Styles.Add("MyStyle", Word.WdStyleType.wdStyleTypeParagraph);
            
            style.Font.Name = "Times New Roman";
            style.Font.Size = 24;

            foreach (Word.Paragraph paragraph in document.Paragraphs)
            {
                paragraph.set_Style(style);
                //paragraph.LineSpacing = 1;
            }

            // Настройка отступов
            Word.PageSetup pageSetup = document.PageSetup;
            pageSetup.LeftMargin = 20;
            pageSetup.RightMargin = 20;
            pageSetup.TopMargin = 20;
            pageSetup.BottomMargin = 20;

            int startRow = 11;
            int currentRow = startRow;
            string order_id;

            // Поиск ячейки с дополнительной информацией
            dynamic otherColumn = sheet.Cells[9, 16].Value;
            Excel.Range otherCell = sheet.Cells[9, 7];

            // Переход вправо, учитывая объединение ячеек
            otherCell = otherCell.MergeArea[otherCell.MergeArea.Rows.Count, otherCell.MergeArea.Columns.Count];
            otherCell = sheet.Cells[otherCell.Row, otherCell.Column + 1];

            if (otherCell.Value == null)
            {
                otherColumn = otherCell.Value;
            }

            // Чтение данных из Excel и создание документа Word
            while (true)
            {
                dynamic orderCell = sheet.Cells[3, 2].Value;
                dynamic orderStrCell = sheet.Cells[7, 6].Value;

                string order_id_str = (orderCell != null) ? orderCell.ToString() : string.Empty;
                string order_str = (orderStrCell != null) ? orderStrCell.ToString() : string.Empty;

                order_id = ExtractOrderId(order_id_str);
                string delivery_info = ExtractDeliveryInfo(order_id_str, sheet.Cells[7, 6].Value, sheet);
                string customer = ExtractCustomer(order_str);
                string city_info = ExtractCityInfo(order_str);
                int quantityColumn = FindQuantityColumn(sheet);
                
                dynamic productTextCell = sheet.Cells[currentRow, 7].Value;
                string productText = (productTextCell != null) ? productTextCell.ToString() : string.Empty;

                if (productTextCell != null)
                {
                    Dictionary<string, string> productInfo = ExtractProductInfo(productText);

                    if (productInfo.Count > 0 && sheet.Cells[currentRow, quantityColumn].Value != null)
                    {
                        int quantity = (int)sheet.Cells[currentRow, quantityColumn].Value;
                        dynamic other = null;

                        if (otherColumn == null)
                        {
                            other = sheet.Cells[currentRow, otherCell.Column].Value;
                            if (other != null)
                            {
                                other = Regex.Replace(other.ToString(), @"\s+", " ").Trim();
                            }
                        }

                        for (int i = 0; i < quantity; i++)
                        {
                            // Счёт
                            Word.Paragraph mainParagraph = document.Content.Paragraphs.Add();
                            Word.Range orderRun = mainParagraph.Range;
                            orderRun.Text = "Счёт №: ";
                            orderRun.Font.Size = 24;
                            orderRun.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                            orderRun.InsertAfter($"{order_id} {delivery_info}");
                            orderRun.Bold = 1;
                            orderRun.Font.Size = 72;
                            mainParagraph.Range.InsertParagraphAfter();

                            // Заказчик
                            Word.Paragraph customerParagraph = document.Content.Paragraphs.Add();
                            Word.Range customerRun = customerParagraph.Range;
                            customerRun.Text = "Заказчик: ";
                            customerRun.Font.Size = 24;
                            customerRun.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                            customerRun.InsertAfter($"{customer}");
                            customerRun.Bold = 1;
                            customerRun.Font.Size = 28;
                            customerParagraph.Range.InsertParagraphAfter();

                            // Город
                            Word.Paragraph cityParagraph = document.Content.Paragraphs.Add();
                            Word.Range cityRun = cityParagraph.Range;
                            cityRun.Text = $"Город: {city_info}";
                            cityRun.Font.Size = 24;
                            cityParagraph.Range.InsertParagraphAfter();

                            // Наименование
                            Word.Paragraph productNameParagraph = document.Content.Paragraphs.Add();
                            Word.Range productNameRun = productNameParagraph.Range;
                            productNameRun.Text = "Наименование: ";
                            cityRun.Font.Size = 24;
                            productNameRun.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                            productNameRun.InsertAfter($"{productInfo["info"]} {(other != null ? other : string.Empty)}");
                            productNameRun.Bold = 1;
                            productNameRun.Font.Size = 28;
                            productNameParagraph.Range.InsertParagraphAfter();

                            // Габариты
                            if (!string.IsNullOrEmpty(productInfo["dimensions"]))
                            {
                                Word.Paragraph dimensionsParagraph = document.Content.Paragraphs.Add();
                                Word.Range dimensionsRun = dimensionsParagraph.Range;
                                dimensionsRun.Text = $"Габариты изделия: {productInfo["dimensions"]}";
                                dimensionsRun.Font.Size = 16;
                                dimensionsParagraph.Range.InsertParagraphAfter();
                            }

                            // Цвет
                            if (!string.IsNullOrEmpty(productInfo["color"]))
                            {
                                Word.Paragraph colorParagraph = document.Content.Paragraphs.Add();
                                Word.Range colorRun = colorParagraph.Range;
                                colorRun.Text = $"Цвет: ";
                                colorRun.Font.Size = 24;
                                colorRun.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                                colorRun.InsertAfter($"{productInfo["color"]}");
                                colorRun.Bold = 1;
                                colorRun.Font.Size = 28;
                            }

                            //Картинка @"D:\Programs\VisualStudio\Programs\ExcelToWord\Fragile.png"
                            string imagePath = Path.Combine(Application.StartupPath, "Fragile.png");
                            Word.Shape imgShape = document.Shapes.AddPicture(imagePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                            // Устанавливаем позицию изображения
                            imgShape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                            imgShape.Width = document.Application.CentimetersToPoints(5);
                            imgShape.Height = document.Application.CentimetersToPoints(5);
                            imgShape.WrapFormat.Type = Word.WdWrapType.wdWrapBehind;
                            imgShape.Left = document.PageSetup.PageWidth - document.PageSetup.RightMargin - imgShape.Width - 3;
                            imgShape.Top = document.PageSetup.TopMargin - 30;

                            document.Bookmarks[@"\EndOfDoc"].Range.Select();
                            object breakTypePage = Word.WdBreakType.wdPageBreak;
                            document.Application.Selection.InsertBreak(ref breakTypePage);
                        }
                    }

                    currentRow++;
                }
                else
                {
                    //Удаление лишней страницы в конце документа и выход из цикла
                    Word.Selection selection = wordApp.Selection;
                    selection.EndKey(Word.WdUnits.wdStory);
                    selection.MoveLeft(Word.WdUnits.wdCharacter, 2);
                    selection.Delete();

                    break;
                }
            }

            // Получение ячейки, которая находится на 2 пункта ниже таблицы
            Excel.Range belowCell = sheet.Cells[currentRow + 3, 2];

            string excelFileName = Path.GetFileNameWithoutExtension(excelFilePath);
            FileName = $"{excelFileName}.docx";
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string fileDir = Path.Combine(desktopPath, FileName);

            document.SaveAs(fileDir);
            document.Close();
            wordApp.Quit();
            excelApp.Quit();
            City = null;
            DeliveryLetter = null;
            // Process.Start(filePath); Открытие документа по завершении обработки
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            
        }

        private async Task processBtn_ClickAsync(object sender, EventArgs e)
        {
            progressBar1.Minimum = 0;
            progressBar1.Maximum = SelectedFilePaths.Count + 1;
            progressBar1.Value = 0;
            progressBar1.Value++;

            foreach (string filePath in SelectedFilePaths)
            {
                await Task.Run(() =>
                {
                    ReadExcelAndCreateWordDocument(filePath);
                });

                // Обновляем состояние полосы загрузки
                progressBar1.Value++;

                // Проигрываем звук оповещения
                System.Media.SystemSounds.Asterisk.Play();

                // Создаем экземпляр формы CustomMessageBox и передаем сообщение
                CustomMessageBox customMessageBox = new CustomMessageBox(FileName);

                // Отображаем окно модально (асинхронно)
                customMessageBox.ShowDialog();

                FileName = null;
            }
        }

        private void uploadBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Выберите файлы";
            openFileDialog.Filter = "Excel таблица (*.xls;*.xlsx)|*.xls;*.xlsx|Все файлы (*.*)|*.*";
            openFileDialog.Multiselect = true; // Разрешить выбирать несколько файлов
            DialogResult result = openFileDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                SelectedFilePaths = openFileDialog.FileNames.ToList();
                fileDirLabel.Text = string.Join(Environment.NewLine, SelectedFilePaths);
                processBtn.Enabled = true;
            }
        }

        private void iron66_Load(object sender, EventArgs e)
        {

        }

        private void fileDirLabel_Click(object sender, EventArgs e)
        {

        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }
    }
}

