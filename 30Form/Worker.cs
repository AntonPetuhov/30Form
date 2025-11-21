using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace _30Form
{
    public class Worker : BackgroundService
    {

        public Worker()
        {
            
        }


        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            ExcelPackage.License.SetNonCommercialPersonal("FNKC");

            /*
            string filePath = AppDomain.CurrentDomain.BaseDirectory + "\\" + $"Отчет 30 форма КДЛ " + DateTime.Now.ToShortDateString() + ".xlsx";
            string worksheetName = "Форма № 30";


            // создаем файл
            FileInfo newFile = new FileInfo(filePath);

            if (newFile.Exists)
            {
                newFile.Delete();  // ensures we create a new workbook
                newFile = new FileInfo(filePath);
            }

            ExcelGenerator reportExcel = new ExcelGenerator(); // объект для работы с excel
            reportExcel.GenerateExcel(newFile, worksheetName); // создаем файл с заголовками
            */

            Reporter reporter = new Reporter();
            reporter.StartReporter();




            while (!stoppingToken.IsCancellationRequested)
            {
                await Task.Delay(1000, stoppingToken);
            }
        }
    }

    
}
