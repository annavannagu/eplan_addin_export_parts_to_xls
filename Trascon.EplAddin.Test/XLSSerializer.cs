using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Trascon.EplAddin.ExportToXLS
{
    class XLSSerializer
    {
        public XLSSerializer() { }

        private string[] articles;
        private string[] places;
        private string[] header;
        private string[] headerStart;
        private string[] headerEnd;
        private string[] manufacturers;

        public void Serialize(string filepath, List<Part> parts)
        {
            articles = (from p in parts orderby p.PartNo select p.PartNo).Distinct().ToArray();
            places = (from p in parts orderby p.Place select p.Place).Distinct().ToArray();
            manufacturers = (from p in parts orderby p.Manufacturer select p.Manufacturer).Distinct().ToArray();

            using (ExcelPackage excel = new ExcelPackage())
            {
                // добавление листа в книгу

                excel.Workbook.Worksheets.Add("Спецификация");
                excel.Workbook.Worksheets.Add("Ссылки");
                //заголовки
                string[] vs = CreateHeader();


                //ряд заголовков и колонка изделий
                var headerRow = new List<string[]>() { vs };
                var articlesColl = new List<string[]>();
                foreach (string art in articles)
                {
                    articlesColl.Add(new string[] { art.TrimEnd(';') });
                }


                //граничные ячейки
                int lastColumn = headerRow[0].Length;
                int lastRow = articles.Count() + 11;

                //выбор активного листа вкниге Excel
                var worksheet = excel.Workbook.Worksheets["Спецификация"];

                #region HeaderStyle
                worksheet.Cells[11, 1, 11, lastColumn].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[11, 1, 11, lastColumn].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[11, 1, 11, lastColumn].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[11, 1, 11, lastColumn].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[11, 1, 11, lastColumn].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Cells[11, 1, 11, lastColumn].Style.WrapText = true;
                worksheet.Cells[11, 1, 11, lastColumn].Style.Font.Bold = true;
                worksheet.Cells[11, 1, 11, lastColumn].Style.Font.Size = 12;
                worksheet.Cells[11, 1, 11, lastColumn].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[11, 1, 11, lastColumn].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                #endregion

                #region DataStyle
                worksheet.Cells[12, 4, lastRow, lastColumn].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[12, 4, lastRow, lastColumn].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[12, 4, lastRow, lastColumn].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[12, 4, lastRow, lastColumn].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[12, 4, lastRow, lastColumn].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Cells[12, 4, lastRow, lastColumn].Style.Font.Bold = false;
                worksheet.Cells[12, 4, lastRow, lastColumn].Style.Font.Size = 10;
                #endregion

                #region ArticlesStyle
                worksheet.Cells[12, 2, lastRow, 2].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[12, 2, lastRow, 2].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[12, 2, lastRow, 2].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[12, 2, lastRow, 2].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[12, 2, lastRow, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                worksheet.Cells[12, 2, lastRow, 2].Style.Font.Bold = false;
                worksheet.Cells[12, 2, lastRow, 2].Style.Font.Size = 10;
                worksheet.Cells[12, 2, lastRow, 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[12, 2, lastRow, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                #endregion

                #region DescriptionStyle
                worksheet.Cells[12, 3, lastRow, 3].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[12, 3, lastRow, 3].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[12, 3, lastRow, 3].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[12, 3, lastRow, 3].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[12, 3, lastRow, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                worksheet.Cells[12, 3, lastRow, 3].Style.WrapText = true;
                worksheet.Cells[12, 3, lastRow, 3].Style.Font.Bold = false;
                worksheet.Cells[12, 3, lastRow, 3].Style.Font.Size = 10;
                worksheet.Cells[12, 3, lastRow, 3].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[12, 3, lastRow, 3].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                #endregion

                #region PositionStyle
                worksheet.Cells[12, 1, lastRow, 1].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[12, 1, lastRow, 1].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[12, 1, lastRow, 1].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[12, 1, lastRow, 1].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[12, 1, lastRow, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Cells[12, 1, lastRow, 1].Style.Font.Bold = false;
                worksheet.Cells[12, 1, lastRow, 1].Style.Font.Size = 10;
                #endregion


                //заполнение ячеек из массивов заголовков и изделий проекта
                worksheet.Cells[11, 1, 11, lastColumn].LoadFromArrays(headerRow);
                worksheet.Cells[12, 2, lastRow, 2].LoadFromArrays(articlesColl);

                //порядковые номера
                int number = 1;
                foreach (var art in articles)
                {
                    worksheet.Cells[number + 11, 1].Value = number;
                    number++;
                };

                //заполнение количества изделий
                foreach (Part part in parts)
                {
                    int i = Array.IndexOf(headerRow[0], part.Place) + 1;
                    int j = Array.IndexOf(articles, part.PartNo) + 12;
                    int pr = vs.Length;
                    double temp = Convert.ToDouble(worksheet.Cells[j, i].Value);
                    worksheet.Cells[j, i].Value = temp + part.Quantity;
                    worksheet.Cells[j, 3].Value = part.Description.TrimEnd(';');
                    worksheet.Cells[j, pr].Value = part.Manufacturer;
                };


                //подсчет суммы по изделиям(вставка формулы)
                for (int k = 0; k < articles.Length; k++)
                {
                    //сумма по ячейкам в линию
                    var adr = new ExcelAddress((k + 12), 4, (k + 12), (vs.Length - 5));
                    worksheet.Cells[k + 12, vs.Length - 4].Formula = "=SUM(" + adr.Address + ")";
                    worksheet.Cells[k + 12, vs.Length - 3].Value = 1;
                    worksheet.Cells[k + 12, vs.Length - 3].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                    //делить на упаковки
                    adr = new ExcelAddress((k + 12), (vs.Length - 4), (k + 12), (vs.Length - 4));
                    string start = adr.Address;
                    adr = new ExcelAddress((k + 12), (vs.Length - 3), (k + 12), (vs.Length - 3));
                    string end = adr.Address;
                    worksheet.Cells[k + 12, vs.Length - 2].Formula = "=" + start + "/" + end;
                    worksheet.Cells[k + 12, vs.Length - 2].Style.Font.Bold = true;
                    worksheet.Cells[k + 12, vs.Length - 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[k + 12, vs.Length - 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
                    worksheet.Cells[k + 12, vs.Length - 2].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                    worksheet.Cells[k + 12, vs.Length - 2].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                    worksheet.Cells[k + 12, vs.Length - 1].Value = 0;
                }

                //вставка расходников и работ
                #region Insert strings
                worksheet.Cells[lastRow + 1, 2].Value = "Расходники";
                worksheet.Cells[lastRow + 1, 3].Value = "Провод, наконечники, маркеры, гермовводы";
                worksheet.Cells[lastRow + 1, vs.Length - 1].Value = 3000;

                worksheet.Cells[lastRow + 2, 3].Value = "Работы по сборке, человеко-дней";
                worksheet.Cells[lastRow + 2, vs.Length - 1].Value = 12000;

                worksheet.Cells[lastRow + 4, 3].Value = "ИТОГО ОБОРУДОВАНИЕ, РУБ с НДС";
                worksheet.Cells[lastRow + 4, 3].Style.Font.Bold = true;

                worksheet.Cells[lastRow + 5, 3].Value = "Работы, РУБ с НДС";
                worksheet.Cells[lastRow + 5, 3].Style.Font.Bold = true;

                worksheet.Cells[lastRow + 6, 3].Value = "ИТОГО, РУБ с НДС";
                worksheet.Cells[lastRow + 6, 3].Style.Font.Bold = true;
                #endregion

                //посчет суммы-произведения внизу под шкафом =СУММПРОИЗВ(D12:D110;AC12:AC110)
                for (int k = 4; k < (places.Length + 4); k++)
                {
                    //итого оборудование
                    var adr = new ExcelAddress(12, k, (lastRow + 1), k);
                    string temp = "R12C" + (vs.Length - 1);
                    string priceTop = ExcelCellBase.TranslateFromR1C1(temp, 0, 0);
                    temp = "R" + (lastRow + 1) + "C" + (vs.Length - 1);
                    string priceBottom = ExcelCellBase.TranslateFromR1C1(temp, 0, 0);
                    temp = "=SUMPRODUCT(" + adr.Address + "," + priceTop + ":" + priceBottom + ")";
                    worksheet.Cells[lastRow + 4, k].Formula = temp;

                    //итого работы
                    adr = new ExcelAddress((lastRow + 2), k, (lastRow + 2), k);
                    string work = adr.Address;
                    temp = "R" + (lastRow + 2) + "C" + (vs.Length - 1);
                    string price = ExcelCellBase.TranslateFromR1C1(temp, 0, 0);
                    temp = "=" + work + "*" + price;
                    worksheet.Cells[lastRow + 5, k].Formula = temp;

                    //итого сумма оборудование + работа
                    adr = new ExcelAddress((lastRow + 5), k, (lastRow + 5), k);
                    work = adr.Address;
                    adr = new ExcelAddress((lastRow + 4), k, (lastRow + 4), k);
                    string equip = adr.Address;
                    temp = "=" + work + "+" + equip;
                    worksheet.Cells[lastRow + 6, k].Formula = temp;
                    worksheet.Cells[lastRow + 6, k].Style.Font.Bold = true;
                    worksheet.Cells[lastRow + 6, k].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                    worksheet.Cells[lastRow + 6, k].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                }


                //выравнивание ширины колонок
                worksheet.Column(3).Width = 50;
                worksheet.Column(2).AutoFit();
                worksheet.Column(lastColumn).AutoFit();
                worksheet.View.FreezePanes(12, 4);

                //выбор активного листа вкниге Excel и заполнение прайсов
                worksheet = excel.Workbook.Worksheets["Ссылки"];                
                var manufacturersColl = manufacturers.ToList();
                worksheet.Cells[1, 1, (manufacturers.Length), 1].LoadFromCollection(manufacturersColl);
                for (int i = 1; i < manufacturersColl.Count(); i++)
                {
                    worksheet.Cells[i, 2].Value = "";
                }                

                //сохранение файла
                FileInfo excelFile = new FileInfo(filepath);
                excel.SaveAs(excelFile);
            }            
        }

        private string[] CreateHeader()
        {
            headerStart = new string[] { "#", "Оборудование", "Описание" };
            headerEnd = new string[] { "Всего, шт", "Штук в упаковке", "Всего упаковок", "Цена", "Производитель" };
            header = new string[(headerStart.Length + places.Length + headerEnd.Length)];
            Array.Copy(headerStart, header, headerStart.Length);
            Array.Copy(places, 0, header, headerStart.Length, places.Length);
            Array.Copy(headerEnd, 0, header, headerStart.Length + places.Length, headerEnd.Length);
            return header;
        }
    }
}
