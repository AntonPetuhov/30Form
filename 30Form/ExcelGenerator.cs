using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _30Form
{
    public class ExcelGenerator
    {

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
        public void GenerateExcel(ExcelPackage package, string worksheetName)
        {
            using (package)
            {
                var ws = package.Workbook.Worksheets[worksheetName];

                // Если лист новый - добавляем заголовки
                if (ws.Dimension == null)
                {
                    AddHeaders(ws);
                }

                package.Save();
            }
        }
        

        // вставка данных в excel
        public byte[] InsertData(ExcelPackage package, string worksheetName, List<ReportRow> data) 
        {
            using (package)
            {
                var ws = package.Workbook.Worksheets[worksheetName];
                // Находим первую пустую строку
                int startRow = ws.Dimension.End.Row + 1;

                // Вставляем данные
                for (int i = 0; i < data.Count; i++)
                {
                    int currentRow = startRow + i;

                    ws.Cells[currentRow, 1].Value = data[i].rowName;
                    ws.Cells[currentRow, 2].Value = data[i].hspColumn;
                    ws.Cells[currentRow, 3].Value = data[i].expressColumn;
                    ws.Cells[currentRow, 4].Value = data[i].consultColumn;
                    ws.Cells[currentRow, 5].Value = data[i].ruspoleColumn;
                }

                FormatWorksheet(ws);

                return package.GetAsByteArray();
            } 
        }
        public string Create(string filePath, string worksheetName, List<ReportRow> data)
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

                FormatWorksheet(worksheet);

                // save our new workbook in the output directory and we are done!
                package.SaveAs(newFile); // нужно чтобы в названии файлы было расширение .xlsx
                return newFile.FullName;
            }
        }

        // добавление заголовков таблицы
        static void AddHeaders(ExcelWorksheet worksheet)
        {
            string[] headers = { "Наименование раздела с указанием видов проведенных  исследований",
                                 "Стационар (тестов по пробам, без контроля качества и калибровок)",
                                 "в том числе ЭКСПРЕСС (тестов по пробам, без контроля качества и калибровок)",
                                 "Консультативный отдел (тестов по пробам, без контроля качества и калибровок)",
                                 "Русское поле (забор бм в Русском поле, исследование-на Самора Машела)" };

            for (int i = 0; i < headers.Length; i++)
            {
                worksheet.Cells[2, i + 1].Value = headers[i];
            }
        }

        // форматирование 
        static void FormatWorksheet(ExcelWorksheet worksheet)
        {
            //Add a jpg image and apply some effects (EPPlus 6+ interface).
            //var pic = worksheet.Drawings.AddPicture("fnkc", new FileInfo(AppDomain.CurrentDomain.BaseDirectory + "\\" + "fnkc.png"));
            //pic.SetPosition(1, 1, 1, worksheet.Dimension.End.Column);

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

            // Форматирование заголовков
            using (var range = worksheet.Cells[2, 1, 2, worksheet.Dimension.End.Column])
            {
                range.AutoFilter = true;

                range.Style.Font.Name = "Arial";
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
            worksheet.Column(1).Width = 50;
            for (int i = 2; i <= worksheet.Dimension.End.Column; i++) 
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
