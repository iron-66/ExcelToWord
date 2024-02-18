using System;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;
using System.ComponentModel;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;



namespace ExcelToWord
{
    public partial class iron66 : Form
    {
        public string SelectedFilePath;

        public iron66()
        {
            InitializeComponent();
        }

        // Извлечение номера заказа
        private string ExtractOrderId(string text)
        {
            Match match = Regex.Match(text, @"№ (\d+)");
            return match.Success ? match.Groups[1].Value : null;
        }

        // Извлечение информации о доставке
        private string ExtractDeliveryInfo(string text)
        {
            Match match = Regex.Match(text, @"\.(.*?)$");
            return match.Success ? match.Groups[1].Value.Trim() : null;
        }

        // Извлечение информации о заказчике
        private string ExtractCustomer(string text)
        {
            return text.Split(',')[0].Trim();
        }

        // Извлечение информации о городе доставки
        private string ExtractCityInfo(string text)
        {
            Match match = Regex.Match(text, @", г[ .]([А-Яа-я\s\-]+),");
            return match.Success ? match.Groups[1].Value.Trim() : null;
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
                string delivery_info = ExtractDeliveryInfo(order_id_str);
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
                            productNameRun.Text = "Наименование:";
                            cityRun.Font.Size = 24;
                            productNameParagraph.Range.InsertParagraphAfter();

                            Word.Paragraph productParagraph = document.Content.Paragraphs.Add();
                            Word.Range productRun = productParagraph.Range;
                            productRun.Text = $"{productInfo["info"]} {(other != null ? other : string.Empty)}";
                            productRun.Bold = 1;
                            productRun.Font.Size = 28;
                            productParagraph.Range.InsertParagraphAfter();

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

            string outputFileName = $"Заказ {order_id}.docx";
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string filePath = Path.Combine(desktopPath, outputFileName);
            document.SaveAs(filePath);
            document.Close();
            wordApp.Quit();
            excelApp.Quit();
            Process.Start(filePath);
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            
        }

        private void processBtn_Click(object sender, EventArgs e)
        {
            ReadExcelAndCreateWordDocument(SelectedFilePath);

            // Проиграть звук оповещения
            System.Media.SystemSounds.Asterisk.Play();

            // Создаем экземпляр формы CustomMessageBox и передаем сообщение
            CustomMessageBox customMessageBox = new CustomMessageBox();

            // Отображаем окно
            customMessageBox.Show();
        }

        private void uploadBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Выберите файл";
            openFileDialog.Filter = "Excel таблица (*.xls;*.xlsx)|*.xls;*.xlsx|Все файлы (*.*)|*.*";
            DialogResult result = openFileDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                SelectedFilePath = openFileDialog.FileName;
                fileDirLabel.Text = SelectedFilePath;
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

