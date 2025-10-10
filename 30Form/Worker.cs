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

            //string filePath = AppDomain.CurrentDomain.BaseDirectory + "\\" + $"Отчет 30 форма КДЛ (за {DateTime.Now.AddMonths(-1).ToString()}) {DateTime.Now.ToString()}";
           // string filePath = AppDomain.CurrentDomain.BaseDirectory + "\\" + $"Отчет 30 форма КДЛ.xlsx";
            string filePath = AppDomain.CurrentDomain.BaseDirectory + "\\" + $"Отчет 30 форма КДЛ " + DateTime.Now.ToShortDateString() + ".xlsx";
            string worksheetName = "Форма № 30";


            // создаем файл
            FileInfo newFile = new FileInfo(filePath);

            if (newFile.Exists)
            {
                newFile.Delete();  // ensures we create a new workbook
                newFile = new FileInfo(filePath);
            }


            ExcelPackage package = new ExcelPackage(newFile);
            using (package) 
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(worksheetName);
            }

            
            ExcelGenerator reportExcel = new ExcelGenerator(); // объект для работы с excel

            reportExcel.GenerateExcel(package, worksheetName);


            Reporter reporter = new Reporter();
            var reportData = reporter.GetReport();

            reportExcel.InsertData(package, worksheetName, reportData);

            /*
            ExcelGenerator reportExcel = new ExcelGenerator();
            //reportExcel.Generate(filePath, worksheetName, reportData);

            byte[] report = reportExcel.Generate(filePath, worksheetName, reportData);
            File.WriteAllBytes(filePath, report);

            //reportExcel.Create(filePath, worksheetName, reportData);
            */
 

            while (!stoppingToken.IsCancellationRequested)
            {
                await Task.Delay(1000, stoppingToken);
            }
        }
    }

    
}
