using Azure;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Buffers;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _30Form
{
    public class ExcelGenerator
    {
        public string worksheetName { get; set; }
        public FileInfo newFile { get; set; }

        public ExcelGenerator(FileInfo newFile, string worksheetName)
        {
            this.newFile = newFile;
            this.worksheetName = worksheetName;
        }

        public byte[] Generate_(string filePath, string worksheetName, List <ReportRow> data)
        {
            FileInfo newFile = new FileInfo(filePath);

            if (newFile.Exists)
            {
                newFile.Delete();  // ensures we create a new workbook
                newFile = new FileInfo(filePath);
            }

            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                // Добавляем новый лист в пустую книгу
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(worksheetName);
                // Если лист новый - добавляем заголовки
                if (worksheet.Dimension == null)
                {
                    AddHeaders(worksheet);
                }

                // Находим первую пустую строку
                int startRow = worksheet.Dimension.End.Row + 1;

                // Вставляем данные
                for (int i = 0; i < data.Count; i++)
                {
                    int currentRow = startRow + i;

                    worksheet.Cells[currentRow, 1].Value = data[i].rowName;
                    worksheet.Cells[currentRow, 2].Value = data[i].hspColumn;
                    worksheet.Cells[currentRow, 3].Value = data[i].expressColumn;
                    worksheet.Cells[currentRow, 4].Value = data[i].consultColumn;
                    worksheet.Cells[currentRow, 5].Value = data[i].ruspoleColumn;
                }

                FormatWorksheet(worksheet); 

                return package.GetAsByteArray();
            }
        }


        // создаем excel файл
        //public void GenerateExcel(FileInfo newFile, string worksheetName)
        public void GenerateExcel()
        {
            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                // Добавляем новый лист в пустую книгу
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(worksheetName);
                // Если лист новый - добавляем заголовки
                if (worksheet.Dimension == null)
                {
                    AddHeaders(worksheet);
                }

                AddTableHeader(worksheet);

                package.Save();
            }
        }


        // вставка данных в excel
        //public byte[] InsertData(FileInfo newFile, string worksheetName, List<ReportRow> data)
        public byte[] InsertData(List<ReportRow> data) 
        {
            string[] totalnums = { "1.", "2.", "3.", "4.", "5.", "6." };
            string[] subtotalnums = { "1.1", "1.2", "1.3", "1.4", "1.5", "1.6", "1.7", "1.8", "1.9", "4.1", "4.2", "4.3", "4.4", "4u", 
                                      "6.1", "6.2", "6.3", "6.4", "6.5", "6.5.1" };

            using(ExcelPackage package = new ExcelPackage(newFile))
            {
                var worksheet = package.Workbook.Worksheets[worksheetName];

                if (worksheet != null)
                {
                    // Находим первую пустую строку
                    int startRow = worksheet.Dimension.End.Row + 1;

                    // Вставляем данные
                    for (int i = 0; i < data.Count; i++)
                    {
                        int currentRow = startRow + i;

                        worksheet.Cells[currentRow, 1].Value = data[i].rowNumber;
                        worksheet.Cells[currentRow, 2].Value = data[i].rowName;
                        worksheet.Cells[currentRow, 3].Value = data[i].hspColumn;
                        worksheet.Cells[currentRow, 4].Value = data[i].expressColumn;
                        worksheet.Cells[currentRow, 5].Value = data[i].consultColumn;
                        worksheet.Cells[currentRow, 6].Value = data[i].ruspoleColumn;

                        // сли строка - подитог, то выделяем серым
                        if (subtotalnums.Contains(data[i].rowNumber))
                        {
                            FormatRow(worksheet, currentRow);
                        }
                        else if (totalnums.Contains(data[i].rowNumber))
                        {
                            FormatTotals(worksheet, currentRow);
                        }

                    }

                    FormatWorksheet(worksheet);

                    //FormatTotals(worksheet);

                    package.Save();

                }

                return package.GetAsByteArray();
            }
        }

        // добавление и форматирование шапки таблицы
        static void AddTableHeader(ExcelWorksheet worksheet)
        {
            //Add a jpg image and apply some effects (EPPlus 6+ interface).
            //var pic = worksheet.Drawings.AddPicture("fnkc", new FileInfo(AppDomain.CurrentDomain.BaseDirectory + "\\" + "fnkc.png"));
            //pic.SetPosition(1, 1, 1, worksheet.Dimension.End.Column);

            // вставка и форматирование шапки таблицы
            var rangeheader = worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column];
            rangeheader.Merge = true;
            rangeheader.Value = "Отчет КДЛ по форме №30";
            rangeheader.Style.Font.Name = "Tahoma";
            rangeheader.Style.Font.Bold = true;
            rangeheader.Style.Font.Size = 14;
            rangeheader.Style.Font.Color.SetColor(Color.White);
            rangeheader.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rangeheader.Style.Fill.BackgroundColor.SetColor(Color.SteelBlue);
            rangeheader.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            rangeheader.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }
        // добавление заголовков таблицы
        static void AddHeaders(ExcelWorksheet worksheet)
        {
            string[] headers = { "№",
                                 "Наименование раздела с указанием видов проведенных  исследований",
                                 "Стационар (тестов по пробам, без контроля качества и калибровок)",
                                 "в том числе ЭКСПРЕСС (тестов по пробам, без контроля качества и калибровок)",
                                 "Консультативный отдел (тестов по пробам, без контроля качества и калибровок)",
                                 "Русское поле (забор бм в Русском поле, исследование-на Самора Машела)",
                                 "Калибровки",
                                 "Контроль качества"};

            for (int i = 0; i < headers.Length; i++)
            {
                worksheet.Cells[2, i + 1].Value = headers[i];
            }
        }

        // форматирование итогов
        static void FormatTotals(ExcelWorksheet worksheet, int row)
        {
            //worksheet.Cells[3, 1, 3, worksheet.Dimension.End.Column].Style.Font.Bold = true;
            //worksheet.Cells[3, 1, 3, worksheet.Dimension.End.Column].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //worksheet.Cells[3, 1, 3, worksheet.Dimension.End.Column].Style.Fill.BackgroundColor.SetColor(Color.Orange);
            //worksheet.Cells[3, 1, 3, worksheet.Dimension.End.Column].Style.WrapText = true;

            worksheet.Cells[row, 1, row, worksheet.Dimension.End.Column].Style.Font.Bold = true;
            worksheet.Cells[row, 1, row, worksheet.Dimension.End.Column].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, 1, row, worksheet.Dimension.End.Column].Style.Fill.BackgroundColor.SetColor(Color.Orange);
            worksheet.Cells[row, 1, row, worksheet.Dimension.End.Column].Style.WrapText = true;

        }
        // форматирование итогов
        static void FormatTotals_(ExcelWorksheet worksheet)
        {
            // Поиск первой ячейки, содержащей "Искомое значение"
            var foundCell = worksheet.Cells["A:A"];

            foreach (var cell in foundCell)
            {
                if (cell.Value.ToString() == "1")
                {
                    Console.WriteLine(cell.Value);
                }
                //Console.WriteLine(cell.Value);
            }
            
        }

        // заливка подитогов
        static void FormatRow(ExcelWorksheet worksheet, int row )
        {
            worksheet.Cells[row, 1, row, worksheet.Dimension.End.Column].Style.Font.Bold = true;
            worksheet.Cells[row, 1, row, worksheet.Dimension.End.Column].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, 1, row, worksheet.Dimension.End.Column].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            worksheet.Cells[row, 1, row, worksheet.Dimension.End.Column].Style.WrapText = true;
        }

        // форматирование 
        static void FormatWorksheet(ExcelWorksheet worksheet)
        {
            //Add a jpg image and apply some effects (EPPlus 6+ interface).
            //var pic = worksheet.Drawings.AddPicture("fnkc", new FileInfo(AppDomain.CurrentDomain.BaseDirectory + "\\" + "fnkc.png"));
            //pic.SetPosition(1, 1, 1, worksheet.Dimension.End.Column);

            /*
            // вставка и форматирование шапки таблицы
            var rangeheader = worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column];
            rangeheader.Merge = true;
            rangeheader.Value = "Отчет КДЛ по форме №30";
            rangeheader.Style.Font.Name = "Arial";
            rangeheader.Style.Font.Bold = true;
            rangeheader.Style.Font.Size = 14;
            rangeheader.Style.Font.Color.SetColor(Color.White);
            rangeheader.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rangeheader.Style.Fill.BackgroundColor.SetColor(Color.SteelBlue);
            rangeheader.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            rangeheader.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            */

            // Форматирование заголовков
            using (var range = worksheet.Cells[2, 1, 2, worksheet.Dimension.End.Column])
            {
                range.AutoFilter = true;

                range.Style.Font.Name = "Tahoma";
                range.Style.Font.Bold = true;
                range.Style.Font.Color.SetColor(Color.White);
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.SteelBlue);
                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                // 
                range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Left.Style = ExcelBorderStyle.Thin;

                range.Style.WrapText = true;

             
            }

            // ширина столбцов
            worksheet.Column(1).Width = 10;
            worksheet.Column(2).Width = 50;
            for (int i = 3; i <= worksheet.Dimension.End.Column; i++) 
            {
                worksheet.Column(i).Width = 30;
            }

            // Границы для всех данных
            using (var range = worksheet.Cells[1, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column])
            {
                range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

        }


    }
}
