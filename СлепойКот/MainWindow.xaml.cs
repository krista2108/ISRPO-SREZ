using Word = Microsoft.Office.Interop.Word;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Diagnostics;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Windows;
using СлепойКот.Models;
using System.IO;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace СлепойКот
{
    
    public partial class MainWindow : Window
    {
        string TemplateFileName = Directory.GetCurrentDirectory() + @"\Шаблон товарный чек.doc";

        static HttpClient client = new HttpClient();
        BindingList<Sale> sales = new BindingList<Sale>();
        private MyDataClass MyData;

        static Sale saleObj;

        Dictionary<string, double> data = new Dictionary<string, double>();

        public MainWindow()
        {
            InitializeComponent();
            MyData = new MyDataClass();
            DataContext = MyData;

            dpDateStar.DisplayDateEnd = DateTime.Now;
            dpDateStar.DisplayDateStart = DateTime.Parse("01.01.2000");
            dpDateEnd.DisplayDateStart = DateTime.Parse("01.01.2000");
            dpDateEnd.DisplayDateEnd = DateTime.Parse("01.01.2030");

            

        }
        private async void btnEnter_Click(object sender, RoutedEventArgs e)
        {
            if (dpDateStar.SelectedDate == null & dpDateEnd.SelectedDate == null)
                MessageBox.Show("Выберите даты!", "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
            else if (dpDateStar.SelectedDate == null)
                MessageBox.Show("Выберите дату начала!", "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
            else if (dpDateEnd.SelectedDate == null)
                MessageBox.Show("Выберите дату окончания!", "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
            else if (dpDateEnd.SelectedDate < dpDateStar.SelectedDate)
                MessageBox.Show("Дата окончания не может быть больше даты начала!", "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Warning);
            else
            {
                DateTime dateStart = (DateTime)dpDateStar.SelectedDate;
                DateTime dateEnd = (DateTime)dpDateEnd.SelectedDate;

                using (HttpClient request = new HttpClient())
                {
                    var context = new StringContent("", Encoding.UTF8, "applocation/json");
                    HttpResponseMessage httpResponseMessage = await client.PostAsync($"https://localhost:7100/api/Sale?dateStart={dateStart}&dateEnd={dateEnd}", context);
                    string json = httpResponseMessage.Content.ReadAsStringAsync().Result;
                    sales = JsonConvert.DeserializeObject<BindingList<Sale>>(json);

                    foreach (var items in sales)
                    {
                        MyData.MyCategory.Add(new Sale
                        {
                            Client = items.Client,
                            Telephones = items.Telephones,
                            DateSale = items.DateSale
                        });

                        for (int i = 0; i < items.Telephones.Count; i++)
                        {
                            if (!data.Keys.Contains(items.Telephones[i].Manufacturer))
                            {
                                data.Add(items.Telephones[i].Manufacturer, 0);
                            }
                            else continue;
                            data[items.Telephones[i].Manufacturer] += items.Telephones[i].Count;
                        }
                    }
                }
            }
        }

        private List<Telephone> GetTelephones(List<Sale> sales)
        {
            List<Telephone> telephones = new List<Telephone>();
            foreach (Sale s in sales)
            {
                telephones.AddRange(s.Telephones);
            }
            return telephones;
        }
       

        private void btnCheckExcel_Click(object sender, RoutedEventArgs e)
        {
            saleObj = lvSales.SelectedItem as Sale;
            if (saleObj == null)
                MessageBox.Show("Запись не выбрана!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            else
            {
                if (File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "\\Шаблон товарный чек.xls"))
                {
                    try
                    {
                        FileStream fs = File.Open(Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "\\Шаблон товарный чек.xls", FileMode.Open);
                        fs.Close();
                    }
                    catch
                    {
                        MessageBox.Show("Файл Шаблон товарный чек.xls запущен на компьютере. Пожалуйста выключите его",
                            "Файл недоступен",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error);

                        return;
                    }
                }
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                ExcelPackage package = new ExcelPackage();

                ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Чек");
                sheet.Columns[1].Width = 4.43;
                sheet.Columns[2].Width = 3.29;
                sheet.Columns[3].Width = 10.57;
                sheet.Columns[4].Width = 1.71;
                sheet.Columns[5].Width = 0.17;
                sheet.Columns[6].Width = 17.57;
                sheet.Columns[7].Width = 16.14;
                sheet.Columns[8].Width = 4.43;
                sheet.Columns[9].Width = 8.71;
                sheet.Columns[10].Width = 0.08;
                sheet.Columns[11].Width = 5.86;
                sheet.Columns[12].Width = 2.86;
                sheet.Columns[13].Width = 4;
                sheet.Columns[14].Width = 6;
                sheet.Columns[15].Width = 0.08;
                sheet.Rows[1].Height = 15.25;
                sheet.Rows[2].Height = 9.25;
                sheet.Rows[3].Height = 3.75;
                sheet.Rows[4].Height = 23;
                sheet.Cells.Style.Font.Size = 9;
                sheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;


                sheet.Cells[1, 1, 1, 15].Merge = true;
                sheet.Cells[1, 1, 1, 15].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                sheet.Cells[2, 1, 2, 15].Merge = true;
                sheet.Cells[2, 1, 2, 15].Value = "наименование организации, ИНН";
                sheet.Cells[2, 1, 2, 15].Style.Font.Size = 6;

                sheet.Cells[3, 1, 3, 15].Merge = true;
                sheet.Cells[4, 1, 4, 15].Merge = true;
                sheet.Cells[4, 1, 4, 15].Style.Font.Bold = true;
                sheet.Cells[4, 1, 4, 15].Style.Font.Size = 12;

                int Cheque = 1;
                if (File.Exists("int_i.txt"))
                    using (StreamReader reader = new StreamReader("int_i.txt"))
                    {
                        Cheque = int.Parse(reader.ReadToEnd());
                    }
                sheet.Cells[4, 1, 4, 15].Value = $"Товарный чек № {Cheque} от {saleObj.DateSale.ToShortDateString()} г.";
                using (StreamWriter writer = new StreamWriter("int_i.txt", false))
                {
                    writer.WriteLine(Cheque + 1);
                }
                sheet.Cells[5, 1].Value = "№ п/п";
                sheet.Cells[5, 1].Style.WrapText = true;

                sheet.Cells[5, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                sheet.Cells[5, 2, 5, 7].Merge = true;
                sheet.Cells[5, 2, 5, 7].Value = "Наименование, характеристика товара";
                sheet.Cells[5, 2, 5, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Top;

                sheet.Cells[5, 2, 5, 7].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                sheet.Cells[5, 8].Value = "Ед.";
                sheet.Cells[5, 8].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                sheet.Cells[5, 9].Value = "Кол-во";
                sheet.Cells[5, 9].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                sheet.Cells[5, 10, 5, 12].Merge = true;
                sheet.Cells[5, 13, 5, 14].Merge = true;
                sheet.Cells[5, 10, 5, 12].Value = "Цена";
                sheet.Cells[5, 10, 5, 12].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                sheet.Cells[5, 13, 5, 14].Value = "Сумма";
                sheet.Cells[5, 13, 5, 14].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                decimal sum = 0;
                int i = 1;
                foreach (Telephone telephone in saleObj.Telephones)
                {
                    sheet.Rows[i + 5].Height = 19;
                    sheet.Cells[i + 5, 1].Value = i;
                    sheet.Cells[i + 5, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    sheet.Cells[i + 5, 2, i + 5, 7].Merge = true;
                    sheet.Cells[i + 5, 2, i + 5, 7].Value = telephone.NameTelephone + ", " + telephone.Articul;
                    sheet.Cells[i + 5, 2, i + 5, 7].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    sheet.Cells[i + 5, 8].Value = "шт";
                    sheet.Cells[i + 5, 8].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    sheet.Cells[i + 5, 9].Value = telephone.Count;
                    sheet.Cells[i + 5, 9].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    sheet.Cells[i + 5, 10, i + 5, 12].Merge = true;
                    sheet.Cells[i + 5, 13, i + 5, 14].Merge = true;
                    sheet.Cells[i + 5, 10, i + 5, 12].Value = telephone.Cost;
                    sheet.Cells[i + 5, 10, i + 5, 12].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    sum += telephone.Cost * telephone.Count;
                    sheet.Cells[i + 5, 13, i + 5, 14].Value = telephone.Cost * telephone.Count;
                    sheet.Cells[i + 5, 13, i + 5, 14].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    i++;
                }
                i += 5;
                sheet.Rows[i].Height = 19;
                sheet.Cells[i, 1, i, 12].Merge = true;
                sheet.Cells[i, 1, i, 12].Value = "Всего";
                sheet.Cells[i, 1, i, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                sheet.Cells[i, 13, i, 14].Merge = true;

                sheet.Cells[i, 13, i, 14].Value = sum;
                sheet.Cells[i, 13, i, 14].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                i++;
                sheet.Rows[i].Height = 3.75;
                sheet.Cells[i, 1, i, 15].Merge = true;
                i++;
                sheet.Rows[i].Height = 12.25;
                sheet.Cells[i, 1, i, 4].Merge = true;
                sheet.Cells[i, 1, i, 4].Value = "Всего отпущено на сумму:";
                sheet.Cells[i, 5, i, 14].Merge = true;
                sheet.Cells[i, 5, i, 14].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                i++;
                sheet.Rows[i].Height = 7;
                sheet.Cells[i, 1, i, 15].Merge = true;
                i++;
                sheet.Rows[i].Height = 0.75;
                sheet.Cells[i, 1, i, 15].Merge = true;
                i++;
                sheet.Rows[i].Height = 11.5;
                sheet.Cells[i, 1, i, 10].Merge = true;
                sheet.Cells[i, 1, i, 10].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                sheet.Cells[i, 11].Value = "руб.";
                sheet.Cells[i, 12, i, 13].Merge = true;
                sheet.Cells[i, 12, i, 13].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                sheet.Cells[i, 14].Value = "коп.";
                i++;
                sheet.Rows[i].Height = 0.75;
                sheet.Cells[i, 1, i, 15].Merge = true;
                i++;
                sheet.Rows[i].Height = 13.75;
                sheet.Cells[i, 1, i, 15].Merge = true;
                i++;
                sheet.Rows[i].Height = 11.5;
                sheet.Cells[i, 1, i, 2].Merge = true;
                sheet.Cells[i, 1, i, 2].Value = "Продавец";
                sheet.Cells[i, 3].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                sheet.Cells[i, 4, i, 5].Merge = true;
                sheet.Cells[i, 6].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                sheet.Cells[i, 7, i, 15].Merge = true;
                i++;
                sheet.Rows[i].Height = 11.5;
                sheet.Cells[i, 1, i, 2].Merge = true;
                sheet.Cells[i, 3].Value = "подпись";
                sheet.Cells[i, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                sheet.Cells[i, 4, i, 5].Merge = true;
                sheet.Cells[i, 6].Value = "ф.и.о.";
                sheet.Cells[i, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                sheet.Cells[i, 7, i, 15].Merge = true;
                sheet.Cells[i, 6].Style.Font.Size = 6;
                sheet.Cells[i, 3].Style.Font.Size = 6;
                sheet.Cells[1, 1, i, 15].Style.Border.BorderAround(ExcelBorderStyle.Medium, System.Drawing.Color.Blue);
                sheet.PrinterSettings.FitToPage = true;
                File.WriteAllBytes(Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "\\Cheque.xlsx", package.GetAsByteArray());
                Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "\\Cheque.xlsx");
            }
        }

        private void ReplaceWordStud(string studToReplace, string text, Word.Document wordDocument)
        {
            var range = wordDocument.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: studToReplace, ReplaceWith: text);
        }
        private void btnCheckWord_Click(object sender, RoutedEventArgs e)
        {
            var wordApp = new Word.Application();
            wordApp.Visible = false;
            saleObj = lvSales.SelectedItem as Sale;
            if (saleObj == null)
                MessageBox.Show("Запись не выбрана!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            else
            {
                try
                {
                    var wordDocument = wordApp.Documents.Open(TemplateFileName);

                    int count = 0;
                    foreach (var sale in saleObj.Telephones)
                    {
                        var sum = sale.Cost * sale.Count;
                        ReplaceWordStud("{Client}", saleObj.Client.FullName, wordDocument);
                        //ReplaceWordStud("{Date_of_Birth}", dateOfBirthChild.ToString("D"), wordDocument);
                        ReplaceWordStud("NameTelef" + count, sale.NameTelephone, wordDocument);
                        ReplaceWordStud("Art" + count, sale.Articul.ToString(), wordDocument);
                        ReplaceWordStud("{Edizm}", "Шт", wordDocument);
                        ReplaceWordStud("Kol" + count, sale.Count.ToString(), wordDocument);
                        ReplaceWordStud("Price" + count, sale.Cost.ToString(), wordDocument);
                        ReplaceWordStud("Sum" + count, sum.ToString(), wordDocument);
                        count++;
                    }
                    if (count < 10)
                    {
                        for (int i = count; i < 10; i++)
                        {
                            //ReplaceWordStud("{Date_of_Birth}", dateOfBirthChild.ToString("D"), wordDocument);
                            ReplaceWordStud("NameTelef" + count, "", wordDocument);
                            ReplaceWordStud("Art" + count, "", wordDocument);
                            ReplaceWordStud("{Edizm}", "", wordDocument);
                            ReplaceWordStud("Kol" + count, "", wordDocument);
                            ReplaceWordStud("Price" + count, "", wordDocument);
                            ReplaceWordStud("Sum" + count, "", wordDocument);
                            count++;
                        }
                    }
                    wordDocument.SaveAs2(Directory.GetCurrentDirectory() + @"\товарный чек.doc");
                    wordApp.Visible = true;
                }
                catch
                {
                    MessageBox.Show("Произошла ошибка при добавлении!");
                }
            }
        }

        

        
    }
}
