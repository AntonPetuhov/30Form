using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data.SqlClient;
using static System.Runtime.InteropServices.JavaScript.JSType;
using System.Xml.Linq;
using System.IO;
using static Azure.Core.HttpHeader;
using System.Reflection.Metadata;
using System.Data;
using Microsoft.Extensions.Primitives;
//using Microsoft.Data.SqlClient;

namespace _30Form
{
    public class ReportRow
    {
        public string rowNumber { get; set; }
        public string rowName { get; set; }
        public int? hspColumn { get; set; }
        public int? expressColumn { get; set; }
        public int? consultColumn { get; set; }
        public int? ruspoleColumn { get; set; }
    }

    public class Reporter
    {
        string filePath = AppDomain.CurrentDomain.BaseDirectory + "\\" + $"Отчет 30 форма КДЛ " + DateTime.Now.ToShortDateString() + ".xlsx";
        string worksheetName = "Форма № 30";

        string user = "mielogrammauser";
        string password = "Qw123456";

        //string validateDateFrom = "01.09.2025 00:00:00.000";
        //string validateDateTo = "01.09.2025 23:59:59.000";
        DateTime validateDateFrom = new DateTime(2025, 10, 01, 00, 00, 0);
        DateTime validateDateTo = new DateTime(2025, 10, 31, 23, 59, 0);


        public void StartReporter()
        {
            // создаем файл с отчетом
            FileInfo newFile = new FileInfo(filePath);

            if (newFile.Exists)
            {
                newFile.Delete();  // ensures we create a new workbook
                newFile = new FileInfo(filePath);
            }

            //ExcelGenerator reportExcel = new ExcelGenerator(); // объект для работы с excel
            ExcelGenerator reportExcel = new ExcelGenerator(newFile, worksheetName); // объект для работы с excel
            reportExcel.GenerateExcel(); // создаем файл с заголовками

            #region Собираем отчет

            Console.WriteLine(validateDateFrom.ToString());
            Console.WriteLine(validateDateTo.ToString());

            var reportItem = new List<ReportRow>();

            bool debugger = false;

            #region 1. ОКИ

            #region Тесты, по которым будем формировать отчет
            // моча
            string[] urineTests = new string[] { "КИ0125", "КИ0220", "БМ0001", "БМ0020", "КИ0250", "МЗ0080" };
            // Кал
            string[] coproTests = new string[] { "КИ0001" }; // копрограмма
            string[] helminthTests = new string[] { "КИ0009", "КИ0110" }; // яйца гельминтов
            string[] simpleTests = new string[] { "КИ0118" }; // на простейшие
            string[] enterobiosTests = new string[] { "КИ0112" }; // соскоб на энтеробиоз
            string[] hidebloodTests = new string[] { "КИ0021" }; // скрытая кровь
            string[] elastTests = new string[] { "КИ0300", "КИ0290" }; // эластаза, углеводы
            string[] fecTests = coproTests.Union(helminthTests).Union(simpleTests).Union(enterobiosTests).Union(hidebloodTests).Union(elastTests).ToArray();
            // Мокрота
            string[] sputumTests = new string[] { "КИ0415" }; // мокрота
            // Спинномозговая жидкость
            string[] liquorTests = new string[] { "КИ0350", "КИ0360", "КИ0380" };
            // Жидкости
            string[] fluidsTests = new string[] { "КИ0397", "КИ0750", "КИ1401", "КИ0930", "КИ0464", "КИ0550", "КИ0610" };
            // Подитог
            string[] OKI = urineTests.Union(fecTests).Union(sputumTests).Union(liquorTests).Union(fluidsTests).ToArray();
            #endregion

            // добавляем итоги по разделу
            //reportItem.Add(GetTotal("1."));
            reportItem.Add(GetTotalsCountFromDB("1.", OKI, validateDateFrom, validateDateTo));

            #region 1.1 Исследования мочи
            reportItem.Add(GetTotalsCountFromDB("1.1", urineTests, validateDateFrom, validateDateTo));

            int i = 0; //счетчик тестов для нумерации
            foreach (string test in urineTests)
            {
                i++;
                ReportRow testRow = GetDataFromDB(i, test, validateDateFrom, validateDateTo);
                reportItem.Add(testRow);
            }

            #endregion

            #region 1.2 Исследование кала
            // добавляем подитоги
            //reportItem.Add(GetTotal("1.2"));
            reportItem.Add(GetTotalsCountFromDB("1.2", fecTests, validateDateFrom, validateDateTo));

            i = 0; //счетчик тестов для нумерации

            // Копрограмма
            foreach (string test in coproTests)
            {
                i++;
                ReportRow testRow = GetDataFromDB(i, test, validateDateFrom, validateDateTo);
                reportItem.Add(testRow);
            }
            // Обнаружение  яиц гельминтов
            foreach (string test in helminthTests)
            {
                i++; 
                ReportRow testRow = GetDataForArrayFromDB(i, helminthTests, validateDateFrom, validateDateTo);
                reportItem.Add(testRow);
                break;
            }
            // Простейшие
            foreach (string test in simpleTests)
            {
                i++;
                ReportRow testRow = GetDataFromDB(i, test, validateDateFrom, validateDateTo);
                reportItem.Add(testRow);
            }
            // Соскоб на энтеробиоз 
            foreach (string test in enterobiosTests)
            {
                i++;
                ReportRow testRow = GetDataFromDB(i, test, validateDateFrom, validateDateTo);
                reportItem.Add(testRow);
            }
            // Скрытая кровь в кале
            foreach (string test in hidebloodTests)
            {
                i++;
                ReportRow testRow = GetDataFromDB(i, test, validateDateFrom, validateDateTo);
                reportItem.Add(testRow);
            }
            // Эластаза 1, углеводы
            foreach (string test in elastTests)
            {
                i++;
                ReportRow testRow = GetDataFromDB(i, test, validateDateFrom, validateDateTo);
                reportItem.Add(testRow);
            }
            #endregion

            #region 1.3 Исследование мокроты
            //reportItem.Add(GetTotal("1.3"));
            reportItem.Add(GetTotalsCountFromDB("1.3", sputumTests, validateDateFrom, validateDateTo));
            i = 0; //счетчик тестов для нумерации

            foreach (string test in sputumTests)
            {
                i++;
                ReportRow testRow = GetDataFromDB(i, test, validateDateFrom, validateDateTo);
                reportItem.Add(testRow);
            }

            #endregion

            #region  1.4 Исследование спинномозговой жидкости
            //reportItem.Add(GetTotal("1.4"));
            reportItem.Add(GetTotalsCountFromDB("1.4", liquorTests, validateDateFrom, validateDateTo));
            i = 0; //счетчик тестов для нумерации
            
            foreach (string test in liquorTests)
            {
                i++;
                ReportRow testRow = GetDataFromDB(i, test, validateDateFrom, validateDateTo);
                reportItem.Add(testRow);
            }

            #endregion

            #region 1.5 Исследование выпотных жидкостей 
            reportItem.Add(GetTotalsCountFromDB("1.5", fluidsTests, validateDateFrom, validateDateTo));
            i = 0; //счетчик тестов для нумерации

            foreach (string test in fluidsTests)
            {
                i++;
                ReportRow testRow = GetDataFromDB(i, test, validateDateFrom, validateDateTo);
                reportItem.Add(testRow);
            }

            #endregion

            #endregion

            #region 2. Гематологические исследования
            string[] gematologyTests = new string[] { "Г0001", "Г0490", "Г0235", "Г0501", "Г0256", "Г0011", "Г0012", 
                                                      "Г0010", "Г0008", "Г0009", "Г0560", "Г0540", "Г0545" };

            // добавляем итоги
            //reportItem.Add(GetTotal("2."));
            reportItem.Add(GetTotalsCountFromDB("2.", gematologyTests, validateDateFrom, validateDateTo));

            i = 0; //счетчик тестов для нумерации
            foreach (string test in gematologyTests)
            {
                i++;
                ReportRow testRow = GetDataFromDB(i, test, validateDateFrom, validateDateTo);
                reportItem.Add(testRow);
            }

            #endregion

            #region 3. ЦИТОЛОГИЧЕСКИЕ ИССЛЕДОВАНИЯ
            string[] cytologyTests = new string[] { "Ц0020", "Ц0001", "Ц0015", "Ц0010", "Ц0025", "КИ1325", "КИ1115" };

            // добавляем итоги
            reportItem.Add(GetTotalsCountFromDB("3.", cytologyTests, validateDateFrom, validateDateTo));

            i = 0; //счетчик тестов для нумерации
            foreach (string test in cytologyTests)
            {
                i++;
                ReportRow testRow = GetDataFromDB(i, test, validateDateFrom, validateDateTo);
                reportItem.Add(testRow);
            }

            // Пунктат щитовидной железы
            string[] thyroidTests = new string[] {"КИ1325"};

            #endregion

            #region 4. Биохимические исследования

            #region списки тестов
            // Метаболиты, ферменты, электролиты, витамины
            string[] biochemicalTests_4_1 = new string[] { "Б0001", "Б0180", "Б0005", "БМ0011", "Б0155", "Б0150", "Б0010", "Б0015", "Б0020", "Б0025", "Б0030",
                                                           "Б0045", "Б0035", "Б0040", "Б0050", "Б0235", "Б0240", "Б0060", "Б0070", "Б0055", "Б0140", "Б0045",
                                                           "Б0297", "Б0080", "Б0085", "Б0099", "Б0090", "Б0305", "Б0105", "Б0110", "Б0285", "Б0125", "Б0130", 
                                                           "Б0131", "Б0120", "Б0135", "Б0145", "ИМ0231", "ИМ0075", "ИМ0080", "Б0065", "ИМ0112", "БМ0095", 
                                                           "Б0260", "ИМ0232", "КФ0001", "КФ0005", "Э0050", "Б0245", "ОС0010", "Б0255", "Б0300"};
            // КЩС
            string[][] abb_tests =
            {
                new string[] { "КЩС0040", "КЩС0030", "КЩС0035" },
                new string[] { "КЩС0025", "КЩС0015", "КЩС0020" },
                new string[] { "КЩС0010", "КЩС0001", "КЩС0005" },
                new string[] { "КЩС0045", "КЩС0050", "КЩС0055" },
                new string[] { "КЩС0060", "КЩС0065", "КЩС0070" },
                new string[] { "КЩС0075", "КЩС0080", "КЩС0085" },
                new string[] { "КЩС0090", "КЩС0095", "КЩС0100" },
                new string[] { "КЩС0105", "КЩС0110", "КЩС0115" },
                new string[] { "КЩС0120", "КЩС0125", "КЩС0130" },
                new string[] { "КЩС0135", "КЩС0140", "КЩС0145" },
                new string[] { "КЩС0150", "КЩС0155", "КЩС0160" },
                new string[] { "КЩС0165", "КЩС0170", "КЩС0175" },
                new string[] { "КЩС0180", "КЩС0185", "КЩС0190" },
                new string[] { "КЩС0195", "КЩС0200", "КЩС0205" },
                new string[] { "КЩС0210", "КЩС0215", "КЩС0220" },
                new string[] { "КЩС0240", "КЩС0245", "КЩС0250" },
                new string[] { "КЩС0255", "КЩС0260", "КЩС0265" },
                new string[] { "КЩС0270", "КЩС0275", "КЩС0280" },
                new string[] { "КЩС0285", "КЩС0315", "КЩС0295" },
                new string[] { "КЩС0310", "КЩС0290", "КЩС0320" },
                new string[] { "КЩС0325", "КЩС0330", "КЩС0335" },
                new string[] { "КЩС0340", "КЩС0345", "КЩС0350" },
                new string[] { "КЩС0355", "КЩС0360", "КЩС0365" }
            };
            /*
            string[] ab_po2 = new string[] { "КЩС0040", "КЩС0030", "КЩС0035" };
            string[] ab_pCo2 = new string[] { "КЩС0025", "КЩС0015", "КЩС0020" };
            string[] ab_ph = new string[] { "КЩС0010", "КЩС0001", "КЩС0005" };
            string[] ab_so2 = new string[] { "КЩС0045", "КЩС0050", "КЩС0055" };
            string[] ab_thb = new string[] { "КЩС0060", "КЩС0065", "КЩС0070" };
            string[] ab_O2hb = new string[] { "КЩС0075", "КЩС0080", "КЩС0085" };
            string[] ab_CoHb = new string[] { "КЩС0090", "КЩС0095", "КЩС0100" };
            string[] ab_Hhb = new string[] { "КЩС0105", "КЩС0110", "КЩС0115" };
            string[] ab_MetHb = new string[] { "КЩС0120", "КЩС0125", "КЩС0130" };
            string[] ab_k = new string[] { "КЩС0135", "КЩС0140", "КЩС0145" };
            string[] ab_Na = new string[] { "КЩС0150", "КЩС0155", "КЩС0160" };
            string[] ab_Ca = new string[] { "КЩС0165", "КЩС0170", "КЩС0175" };
            string[] ab_Cl = new string[] { "КЩС0180", "КЩС0185", "КЩС0190" };
            string[] ab_glu = new string[] { "КЩС0195", "КЩС0200", "КЩС0205" };
            string[] ab_lac = new string[] { "КЩС0210", "КЩС0215", "КЩС0220" };
            string[] ab_phT = new string[] { "КЩС0240", "КЩС0245", "КЩС0250" };
            string[] ab_Co2T = new string[] { "КЩС0255", "КЩС0260", "КЩС0265" };
            string[] ab_po2T = new string[] { "КЩС0270", "КЩС0275", "КЩС0280" };
            string[] ab_ctO2T = new string[] { "КЩС0285", "КЩС0315", "КЩС0295" };
            string[] ab_p50 = new string[] { "КЩС0310", "КЩС0290", "КЩС0320" };
            string[] ab_cBase = new string[] { "КЩС0325", "КЩС0330", "КЩС0335" };
            string[] ab_cHCO3 = new string[] { "КЩС0340", "КЩС0345", "КЩС0350" };
            string[] ab_cBaseB = new string[] { "КЩС0355", "КЩС0360", "КЩС0365" };
            string[] biochemicalTests_4_2 = ab_po2.Union(ab_pCo2).Union(ab_ph).Union(ab_so2).Union(ab_thb).Union(ab_O2hb).Union(ab_CoHb).Union(ab_Hhb).
                                            Union(ab_MetHb).Union(ab_k).Union(ab_Na).Union(ab_Ca).Union(ab_Cl).Union(ab_glu).Union(ab_lac).Union(ab_phT).
                                            Union(ab_Co2T).Union(ab_po2T).Union(ab_ctO2T).Union(ab_p50).Union(ab_cBase).Union(ab_cHCO3).Union(ab_cBaseB).ToArray();
            */
            string[] biochemicalTests_4_2 = abb_tests[0];
            foreach (string[] abb_arr in abb_tests)
            {
                biochemicalTests_4_2 = biochemicalTests_4_2.Union(abb_arr).ToArray();
            }
            // 4.3 Гормоны и биологически активные соединения
            string[] biochemicalTests_4_3 = new string[] { "ИМ0005", "ИМ0036", "ИМ0037", "ИМ0275", "ИМ0040", "ИМ0015", "ИМ0010", "ИМ0020", "ИМ0001", "ИМ0229",
                                                           "ИМ0271", "ИМ0228", "ИМ0280", "ИМ0060", "ИМ0035", "КИ0385 ", "ИМ0045", "ИМ0270", "ИМ0065", "ИМ0102",
                                                           "ОМ0010", "ИМ0104", "ИМ0101", "ИМ0285", "ОМ383", "ИМ0290", "ИМ0111", "ОМ391", "ИМ0073", "ИМ0322",
                                                           "ИМ0234"};
            // Лекарственный мониторинг
            string[] biochemicalTests_4_4 = new string[] { "АП0645", "ИМ0272", "Б0251", "ИМ1015", "ИМ0625", "ИМ0400", "ИМ1005", "ИМ1010", "Б0298"};
            // БХ мочи
            string[] biochemicalTests_4u = new string[] { "БМ0070", "БМ0135", "БМ0061", "БМ0016", "БМ0031", "БМ0056", "БМ0036", "ОС0005", "ОС0001",
                                                          "БМ0041", "БМ0045", "БМ0080", "БМ0090", "БМ0100"};
            // Итог
            string[] biochemicalAllTests = biochemicalTests_4_1.Union(biochemicalTests_4_2).Union(biochemicalTests_4_3).
                                           Union(biochemicalTests_4_4).Union(biochemicalTests_4u).ToArray();
            #endregion 

            // добавляем итоги
            //reportItem.Add(GetTotal("4."));
            reportItem.Add(GetTotalsCountFromDB("4.", biochemicalAllTests, validateDateFrom, validateDateTo));

            #region 4.1 Метаболиты, ферменты, электролиты, витамины
            reportItem.Add(GetTotalsCountFromDB("4.1", biochemicalTests_4_1, validateDateFrom, validateDateTo));

            i = 0; //счетчик тестов для нумерации
            foreach (string test in biochemicalTests_4_1)
            {
                i++;
                ReportRow testRow = GetDataFromDB(i, test, validateDateFrom, validateDateTo);
                reportItem.Add(testRow);
            }
            #endregion

            #region 4.2 Газообмен крови и выдыхаемого воздуха, соединения гемоглобина
            // итоги
            reportItem.Add(GetTotalsCountFromDB("4.2", biochemicalTests_4_2, validateDateFrom, validateDateTo));

            i = 0; //счетчик тестов для нумерации
            /*
            i++; 
            ReportRow po2Row = GetDataForArrayFromDB(i, ab_po2, validateDateFrom, validateDateTo);
            i++;
            ReportRow pCo2Row = GetDataForArrayFromDB(i, ab_pCo2, validateDateFrom, validateDateTo);
            i++;
            ReportRow phRow = GetDataForArrayFromDB(i, ab_ph, validateDateFrom, validateDateTo);
            i++;
            ReportRow so2Row = GetDataForArrayFromDB(i, ab_so2, validateDateFrom, validateDateTo);
            i++;
            ReportRow thbRow = GetDataForArrayFromDB(i, ab_thb, validateDateFrom, validateDateTo);
            i++;
            ReportRow O2hbRow = GetDataForArrayFromDB(i, ab_O2hb, validateDateFrom, validateDateTo);
            i++;
            ReportRow CoHbRow = GetDataForArrayFromDB(i, ab_CoHb, validateDateFrom, validateDateTo);
            i++;
            ReportRow HhbRow = GetDataForArrayFromDB(i, ab_Hhb, validateDateFrom, validateDateTo);
            i++;
            ReportRow HhbRow = GetDataForArrayFromDB(i, ab_Hhb, validateDateFrom, validateDateTo);
            */

            foreach (string[] abb_array in abb_tests)
            {
                i++;
                ReportRow testRow = GetDataForArrayFromDB(i, abb_array, validateDateFrom, validateDateTo);
                reportItem.Add(testRow);
            }

            #endregion

            #region 4.3 Гормоны и биологически активные соединения.
            // итоги
            reportItem.Add(GetTotalsCountFromDB("4.3", biochemicalTests_4_3, validateDateFrom, validateDateTo));

            i = 0; //счетчик тестов для нумерации
            foreach (string test in biochemicalTests_4_3)
            {
                i++;
                ReportRow testRow = GetDataFromDB(i, test, validateDateFrom, validateDateTo);
                reportItem.Add(testRow);
            }
            #endregion

            #region 4.4 Лекарственный мониторинг (концентрация лекарственных препаратов)
            // итоги
            reportItem.Add(GetTotalsCountFromDB("4.4", biochemicalTests_4_4, validateDateFrom, validateDateTo));

            i = 0; //счетчик тестов для нумерации
            foreach (string test in biochemicalTests_4_4)
            {
                i++;
                ReportRow testRow = GetDataFromDB(i, test, validateDateFrom, validateDateTo);
                reportItem.Add(testRow);
            }
            #endregion

            #region 4u Биохимические исследования мочи
            // итоги
            reportItem.Add(GetTotalsCountFromDB("4u", biochemicalTests_4u, validateDateFrom, validateDateTo));

            i = 0; //счетчик тестов для нумерации
            foreach (string test in biochemicalTests_4u)
            {
                i++;
                ReportRow testRow = GetDataFromDB(i, test, validateDateFrom, validateDateTo);
                reportItem.Add(testRow);
            }
            #endregion

            #endregion

            #region 5. КОАГУЛОГИЧЕСКИЕ ИССЛЕДОВАНИЯ
            string[] coagulogramTests = new string[] { "ГЕ0050", "ГЕ0132", "ГЕ0055", "ГЕ0075", "ГЕ0105", "ГЕ0056", "ГЕ0046", "ГЕ0045" };

            // добавляем итоги
            reportItem.Add(GetTotalsCountFromDB("5.", coagulogramTests, validateDateFrom, validateDateTo));

            i = 0; //счетчик тестов для нумерации
            foreach (string test in coagulogramTests)
            {
                i++;
                ReportRow testRow = GetDataFromDB(i, test, validateDateFrom, validateDateTo);
                reportItem.Add(testRow);
            }
            #endregion

            #region 6. ИММУНОЛОГИЧЕСКИЕ ИССЛЕДОВАНИЯ

            #region Тесты, по которым будем формировать отчет
            // Иммуногематология
            string[] immunologyTests_6_1 = new string[] { "Ф0210", "Ф0215", "Ф0219", "Ф0240", "Ф0230", "Ф0260", "Ф0275", "Ф0290", 
                                                          "Ф0218", "Ф0270", "Ф0225", "Ф0220", "Ф0310" };
            // Иммунологические маркеры резистентности
            string[] immunologyTests_6_2 = new string[] { "Б0275", "Б0271", "Б0220", "Б0302", "Б0290", "ИМ0210", "ИМ0215", "ИМ0069",
                                                          "ИМ0130", "ИМ0140", "ИМ0135", "ИМ0192", "ИМ0405", "ИМ0220", "ИМ0225" };
            // Показатели клеточного иммунитета 
            string[] immunologyTests_6_3 = new string[] { "ИМ0072" };
            // Онкомаркеры
            string[] immunologyTests_6_4 = new string[] { "ИМ0265", "ИМ0266", "ИМ0100", "ИМ0095", "ИМ0085", "ИМ0090", "ИМ0071", "ОМ0001" };
            // Аутоантитела другие 
            string[] immunologyTests_6_5 = new string[] { "ИМ0030", "ИМ0025", "Б0310", "АИ0015", "АИ0020", "АИ0040", "АИ0005", "АИ0010",
                                                          "АИ0050", "АИ0055", "ОМ0391", "АИ0015", "АИ0020", "ИМ0355", "ИМ0360", "ИМ0365",
                                                          "ИМ0370", "ИМ0375", "ИМ0380", "ИМ0385", "ИМ0074", "Ф0316", "АИ0025", "АИ0030", "ИФА0100", "ИФА0359"};
            // Итоги
            string[] immunologyAllTests = immunologyTests_6_1.Union(immunologyTests_6_2).Union(immunologyTests_6_3).Union(immunologyTests_6_4).Union(immunologyTests_6_5).ToArray();
            #endregion

            // добавляем итоги по разделу
            //reportItem.Add(GetTotal("6."));
            reportItem.Add(GetTotalsCountFromDB("6.", immunologyAllTests, validateDateFrom, validateDateTo));

            #region 6.1 Иммуногематология
            reportItem.Add(GetTotalsCountFromDB("6.1", immunologyTests_6_1, validateDateFrom, validateDateTo));

            i = 0; //счетчик тестов для нумерации
            foreach (string test in immunologyTests_6_1)
            {
                i++;
                ReportRow testRow = GetDataFromDB(i, test, validateDateFrom, validateDateTo);
                reportItem.Add(testRow);
            }
            #endregion

            #region 6.2 Иммунологические маркеры резистентности
            reportItem.Add(GetTotalsCountFromDB("6.2", immunologyTests_6_2, validateDateFrom, validateDateTo));

            i = 0; //счетчик тестов для нумерации
            foreach (string test in immunologyTests_6_2)
            {
                i++;
                ReportRow testRow = GetDataFromDB(i, test, validateDateFrom, validateDateTo);
                reportItem.Add(testRow);
            }
            #endregion

            #region 6.3 Показатели клеточного иммунитета 
            reportItem.Add(GetTotalsCountFromDB("6.3", immunologyTests_6_3, validateDateFrom, validateDateTo));

            i = 0; //счетчик тестов для нумерации
            foreach (string test in immunologyTests_6_3)
            {
                i++;
                ReportRow testRow = GetDataFromDB(i, test, validateDateFrom, validateDateTo);
                reportItem.Add(testRow);
            }
            #endregion

            #region 6.4 Онкомаркеры 
            reportItem.Add(GetTotalsCountFromDB("6.4", immunologyTests_6_4, validateDateFrom, validateDateTo));

            i = 0; //счетчик тестов для нумерации
            foreach (string test in immunologyTests_6_4)
            {
                i++;
                ReportRow testRow = GetDataFromDB(i, test, validateDateFrom, validateDateTo);
                reportItem.Add(testRow);
            }
            #endregion

            #region 6.5 Аутоантитела другие 
            reportItem.Add(GetTotalsCountFromDB("6.5", immunologyTests_6_5, validateDateFrom, validateDateTo));

            i = 0; //счетчик тестов для нумерации
            foreach (string test in immunologyTests_6_5)
            {
                i++;
                ReportRow testRow = GetDataFromDB(i, test, validateDateFrom, validateDateTo);
                reportItem.Add(testRow);
            }
                #endregion


                #endregion

            #region Аллергопанель

            string[] allergenTests = GetTestsArray("Аллергопанель", validateDateFrom, validateDateTo);

            reportItem.Add(GetTotalsCountFromDB("6.5.1", allergenTests, validateDateFrom, validateDateTo));

            i = 0;
            foreach (string test in allergenTests) 
            {
                i++;
                ReportRow testRow = GetDataFromDB(i, test, validateDateFrom, validateDateTo);
                reportItem.Add(testRow);
            }

            #endregion


            // втсавляем данные в excel
            reportExcel.InsertData(reportItem);

            #endregion
        }

        // получаем итоги и подитоги
        public ReportRow GetTotal(string totalPar)
        {
            switch (totalPar)
            {
                case "1.": 
                    return new ReportRow { rowNumber = "1.", rowName = "Химико-микроскопическое исследование биологических жидкостей (ОБЩЕКЛИНИЧЕСКИЕ)", consultColumn = null, expressColumn = null, hspColumn = null, ruspoleColumn = null };
                case "1.1":
                    return new ReportRow { rowNumber = "1.1", rowName = "Исследование мочи", consultColumn = null, expressColumn = null, hspColumn = null, ruspoleColumn = null };
                case "1.2":
                    return new ReportRow { rowNumber = "1.2", rowName = "Исследование кала", consultColumn = null, expressColumn = null, hspColumn = null, ruspoleColumn = null };
                case "1.3":
                    return new ReportRow { rowNumber = "1.3", rowName = "Исследование мокроты", consultColumn = null, expressColumn = null, hspColumn = null, ruspoleColumn = null };
                case "1.4":
                    return new ReportRow { rowNumber = "1.4", rowName = "Исследование спинномозговой жидкости", consultColumn = null, expressColumn = null, hspColumn = null, ruspoleColumn = null };
                case "2.":
                    return new ReportRow { rowNumber = "2.", rowName = "ГЕМАТОЛОГИЧЕСКИЕ ИССЛЕДОВАНИЯ", consultColumn = null, expressColumn = null, hspColumn = null, ruspoleColumn = null };
                case "4.":
                    return new ReportRow { rowNumber = "4.", rowName = "БИОХИМИЧЕСКИЕ ИССЛЕДОВАНИЯ", consultColumn = null, expressColumn = null, hspColumn = null, ruspoleColumn=null };
                case "4.1":
                    return new ReportRow { rowNumber = "4.1", rowName = "Метаболиты, ферменты, электролиты, витамины", consultColumn = null, expressColumn = null, hspColumn = null, ruspoleColumn = null };
                case "4.2":
                    return new ReportRow { rowNumber = "4.2", rowName = "Газообмен крови и выдыхаемого воздуха, соединения гемоглобина", consultColumn = null, expressColumn = null, hspColumn = null, ruspoleColumn = null };
                case "4.3":
                    return new ReportRow { rowNumber = "4.3", rowName = " Гормоны и биологически активные соединения", consultColumn = null, expressColumn = null, hspColumn = null, ruspoleColumn = null };
                case "5.":
                    return new ReportRow { rowNumber = "5.", rowName = "КОАГУЛОГИЧЕСКИЕ ИССЛЕДОВАНИЯ", consultColumn = null, expressColumn = null, hspColumn = null, ruspoleColumn = null };
                case "6.":
                    return new ReportRow { rowNumber = "6.", rowName = "ИММУНОЛОГИЧЕСКИЕ ИССЛЕДОВАНИЯ", consultColumn = null, expressColumn = null, hspColumn = null, ruspoleColumn = null };
                case "6.1":
                    return new ReportRow { rowNumber = "6.1", rowName = "Иммуногематология", consultColumn = null, expressColumn = null, hspColumn = null, ruspoleColumn = null };
                case "6.2":
                    return new ReportRow { rowNumber = "6.2", rowName = "Иммунологические маркеры резистентности", consultColumn = null, expressColumn = null, hspColumn = null, ruspoleColumn = null };
                case "6.3":
                    return new ReportRow { rowNumber = "6.3", rowName = "Показатели клеточного иммунитета", consultColumn = null, expressColumn = null, hspColumn = null, ruspoleColumn = null };
                case "6.4":
                    return new ReportRow { rowNumber = "6.4", rowName = "Онкомаркеры", consultColumn = null, expressColumn = null, hspColumn = null, ruspoleColumn = null };
                default:
                    return null;
            }
        }

        // получаем наименование строк итогов
        public string GetTotalName(string totalPar)
        {
            switch (totalPar)
            {
                case "1.":
                    return"Химико-микроскопическое исследование биологических жидкостей (ОБЩЕКЛИНИЧЕСКИЕ)";
                case "1.1":
                    return "Исследование мочи";
                case "1.2":
                    return "Исследование кала";
                case "1.3":
                    return "Исследование мокроты";
                case "1.4":
                    return "Исследование спинномозговой жидкости";
                case "1.5":
                    return "Исследование выпотных жидкостей (экссудатов и транссудатов)";
                case "2.":
                    return "ГЕМАТОЛОГИЧЕСКИЕ ИССЛЕДОВАНИЯ";
                case "3.":
                    return "ЦИТОЛОГИЧЕСКИЕ ИССЛЕДОВАНИЯ";
                case "4.":
                    return "БИОХИМИЧЕСКИЕ ИССЛЕДОВАНИЯ";
                case "4.1":
                    return "Метаболиты, ферменты, электролиты, витамины";
                case "4.2":
                    return "Газообмен крови и выдыхаемого воздуха, соединения гемоглобина";
                case "4.3":
                    return "Гормоны и биологически активные соединения";
                case "4.4":
                    return "Лекарственный мониторинг (концентрация лекарственных препаратов)";
                case "4u":
                    return "Биохимические исследования мочи";
                case "5.":
                    return "КОАГУЛОГИЧЕСКИЕ ИССЛЕДОВАНИЯ";
                case "6.":
                    return "ИММУНОЛОГИЧЕСКИЕ ИССЛЕДОВАНИЯ";
                case "6.1":
                    return "Иммуногематология";
                case "6.2":
                    return "Иммунологические маркеры резистентности";
                case "6.3":
                    return "Показатели клеточного иммунитета";
                case "6.4":
                    return "Онкомаркеры";
                case "6.5":
                    return "Аутоантитела другие";
                case "6.5.1":
                    return "Аллергопанель";
                default:
                    return null;
            }
        }

        // если название строки в отчете не совпадает в именем теста
        public string GetRowName(string testCodePar)
        {
            switch (testCodePar) 
            {
                case "КИ0125":
                    return "Клинический (общий) анализ мочи (ОАМ)";
                case "КИ0220":
                    return "Микроскопия осадка мочи в нативном препарате или исследование осадка на анализаторе";
                case "КИ0250":
                    return "Анализ мочи по Нечипоренко";
                case "МЗ0080":
                    return "Анализ мочи по Зимницкому";
                case "КИ0001":
                    return "Копрограмма";
                case "КИ0009":
                    return "Обнаружение  яиц гельминтов";
                case "КИ0021":
                    return "Скрытая кровь в кале";
                case "КИ0400":
                    return "Исследование физических свойств мокроты";
                case "КИ0415":
                    return "Микроскопическое исследование мокроты";
                case "КИ0350":
                    return "Общий анализ спинно-мозговой жидкости";
                case "КИ0397":
                    return "Общий анализ плевральной жидкости";
                case "КИ1401":
                    return "Общий анализ перитонеальной (асцитической) жидкости";
                case "КИ0464":
                    return "Исследование физических свойств синовиальной жидкости";
                case "КИ0610":
                    return "Микроскопическое исследование лаважной жидкости ";
                case "КИ0326":
                    return "Морфологическое и цитохимическое исследование биологических жидкостей при остром лейкозе";
                case "Г0001":
                    return "Клинический (общий) анализ крови (ОАК)";
                case "Г0235":
                    return "Лейкоцитарная формула";
                case "КИ1325":
                    return "Цитологическое исследование микропрепарата тонкоигольной аспирационной биопсии щитовидной железы";
                case "КИ1115":
                    return "Микроскопическое исследование пунктатов органов кроветворения (костный мозг, селезенка, лимфатические узлы) на лейшмании (Leishmania spp.)";
                case "Б0180":
                    return "белковые фракции";
                case "КФ0005":
                    return "Дифференциальная диагностика метгемоглобинемий";
                case "Э0050":
                    return "Количественная оценка соотношения типов гемоглобина (фракции гемоглобина)";
                case "КЩС0040":
                    return "Парциальное давление кислорода";
                case "КЩС0025":
                    return "парциальное давление углекислого газа";
                case "КЩС0010":
                    return "рН крови";
                case "КЩС0045":
                    return "Сатурация кислородом (sO2)";
                case "КЩС0060":
                    return "Гемоглобин, массовая концентрация в крови";
                case "КЩС0075":
                    return "Оксигемоглобин";
                case "КЩС0090":
                    return "Карбоксигемоглобин (fCOHb)";
                case "КЩС0105":
                    return "Восстановленный гемоглобин (fHHb)";
                case "КЩС0120":
                    return "Метгемоглобин (fMtHb)";
                case "КЩС0135":
                    return "Калий (К)";
                case "КЩС0150":
                    return "Натрий (Na)";
                case "КЩС0165":
                    return "Кальций ионизированный";
                case "КЩС0180":
                    return "Хлор (Cl-)";
                case "КЩС0195":
                    return "Глюкоза";
                case "КЩС0210":
                    return "Лактат";
                case "КЩС0240":
                    return "pHt";
                case "КЩС0255":
                    return "Парциальное давление диоксида углерода с Т-поправкой";
                case "КЩС0270":
                    return "Парциальное давление кислорода с Т-поправкой";
                case "КЩС0285":
                    return "Общая концентрация кислорода (tO2)";
                case "КЩС0310":
                    return "Парциальное давление кислорода (р50)";
                case "КЩС0325":
                    return "Избыток оснований стандартный";
                case "КЩС0340":
                    return "Стандартный бикарбонат";
                case "КЩС0355":
                    return "Избыток оснований истинный";
                default:
                    return "";
            }
        }

        // примечание к строке отчета
        public string GetRowComment(string testCodePar)
        {
            return "";
        }

        #region SQL скрипты
        // получаем количество для каждого теста
        //public ReportRow GetDataFromDB(int number, string testCodePar, string validationFrom, string validationTo)
        public ReportRow GetDataFromDB(int number, string testCodePar, DateTime validationFrom, DateTime validationTo)
        {
            string testName = "";
            string testNumber = number.ToString();
            int testhspCount = 0;
            int testexpressCount = 0;
            int testconsultCount = 0;
            int testruspoleCount = 0;

            testName = GetRowName(testCodePar); // есть ли название строки отчета, которое не совпадает с именем теста

            try
            {
                // Build a configuration object from JSON file
                IConfiguration config = new ConfigurationBuilder()
                    .AddJsonFile("AppConfig.json")
                    .Build();

                // Get a configuration section
                IConfigurationSection section = config.GetSection("ConnnectionStrings");

                string? CGMConnectionString = section["CGMConnection"];
                CGMConnectionString = string.Concat(CGMConnectionString, $"User Id = {user}; Password = {password}");

                using (SqlConnection Connection = new SqlConnection(CGMConnectionString))
                {
                    Connection.Open();

                    //string parameters_ = string.Join(",", biochemicalTests);
                    
                    
                    if(testName == "")
                    {
                        // получаем название теста
                        //SqlCommand TestNameCommand = new SqlCommand($"SELECT a.ana_analys FROM KDLPROD..ana a WHERE a.ana_analyskod = '{testCodePar}'", Connection);
                        SqlCommand TestNameCommand = new SqlCommand($"SELECT a.ana_analys FROM KDLPROD..ana a WHERE a.ana_analyskod = @testcode", Connection);
                        // создаем параметр для имени
                        SqlParameter testParam = new SqlParameter("@testcode", testCodePar);
                        // добавляем параметр к команде
                        TestNameCommand.Parameters.Add(testParam);

                        SqlDataReader Reader = TestNameCommand.ExecuteReader();

                        if (Reader.HasRows)
                        {
                            while (Reader.Read())
                            {
                                if (!Reader.IsDBNull(0)) { testName = Reader.GetString(0); };
                            }
                        }
                        else { }
                        Reader.Close();
                    }
                    

                    // получаем количество тестов

                    SqlCommand TestCountCommand = new SqlCommand(
                        "SELECT  t1.hospital, t2.express, t3.consult, t4.ruspole FROM " +  
                            "(SELECT ROW_NUMBER() OVER (ORDER BY (select 1)) AS RowNum, COUNT(*) AS hospital " + 
                            "FROM KDLReportView kv " +
                                "WHERE kv.test_code IN (@testcode) " +
                                       "AND kv.ValidationDate >= @validationDateFrom AND kv.ValidationDate <= @validationDateTo " +
                                       "AND kv.OperatorId NOT IN ('ATA', 'ITA', 'МАВ', 'РЮВ', 'ПИН', 'АНА', 'ТВИ', 'ШМИ', 'ПАХ', 'ГПГ', 'ПАВ', 'ПДМ', 'ОНВ', 'РМВ', 'СЕА', 'ССО', 'RSIG') " +
                                       "AND kv.ClientCode NOT IN ('56', '57', '58', '59', '60', '61', '63')) t1 " +
                            "JOIN " +
                            "(SELECT ROW_NUMBER() OVER (ORDER BY (select 1)) AS RowNum, COUNT(*) AS express " +
                            "FROM KDLReportView kv " +
                                "WHERE kv.test_code IN (@testcode) " +
                                       "AND kv.ValidationDate >=  @validationDateFrom AND kv.ValidationDate <= @validationDateTo " +
                                       "AND kv.OperatorId IN ('РУМ', 'ПОВ', 'БЛН', 'КОВ', 'KOV', 'ККА')) t2 ON t1.RowNum = t2.RowNum " +
                            "JOIN " +
                            "(SELECT ROW_NUMBER() OVER (ORDER BY (select 1)) AS RowNum, COUNT(*) AS consult " +
                            "FROM KDLReportView kv " +
                                "WHERE kv.test_code IN (@testcode) " +
                                       "AND kv.ValidationDate >= @validationDateFrom AND kv.ValidationDate <= @validationDateTo " +
                                       "AND kv.OperatorId NOT IN ('ATA', 'ITA', 'МАВ', 'РЮВ', 'ПИН', 'АНА', 'ТВИ', 'ШМИ', 'ПАХ', 'ГПГ', 'ПАВ', 'ПДМ', 'ОНВ', 'РМВ', 'СЕА', 'ССО', 'RSIG') " +
                                       "AND kv.ClientCode = 42) t3 ON t1.RowNum = t3.RowNum " +
                            "JOIN " +
                            "(SELECT ROW_NUMBER() OVER (ORDER BY (select 1)) AS RowNum, COUNT(*) AS ruspole " +
                            "FROM KDLReportView kv " +
                                "WHERE kv.test_code IN (@testcode) " +
                                "AND kv.ValidationDate >= @validationDateFrom AND kv.ValidationDate <= @validationDateTo " +
                                "AND kv.analyzer NOT IN ('SAPPHIRE', 'FUS100', 'MIN6200', 'SAPPHIR') " +
                                "AND kv.ClientCode IN ('56', '57', '58', '59', '60', '61', '63') " +
                                "AND kv.OperatorId NOT IN ('ATA', 'ITA', 'МАВ', 'РЮВ', 'ПИН', 'АНА', 'ТВИ', 'ШМИ', 'ПАХ', 'ГПГ', 'ПАВ', 'ПДМ', 'ОНВ', 'РМВ', 'СЕА', 'ССО', 'RSIG')) t4 ON t1.RowNum = t4.RowNum ", Connection);       

                    // создаем параметр для имени
                    SqlParameter testcodeParam = new SqlParameter("@testcode", testCodePar);
                    TestCountCommand.Parameters.Add(testcodeParam);
                    // создаем параметры для дат валидации
                    SqlParameter validationFromParam = new SqlParameter("@validationDateFrom", validationFrom);
                    TestCountCommand.Parameters.Add(validationFromParam);
                    SqlParameter validationToParam = new SqlParameter("@validationDateTo", validationTo);
                    TestCountCommand.Parameters.Add(validationToParam);

                    //Console.WriteLine(TestCountCommand.CommandText);

                    SqlDataReader TestCountReader = TestCountCommand.ExecuteReader();

                    if (TestCountReader.HasRows)
                    {
                        while (TestCountReader.Read())
                        {
                            if (!TestCountReader.IsDBNull(0)) { testhspCount = TestCountReader.GetInt32(0); };
                            if (!TestCountReader.IsDBNull(1)) { testexpressCount = TestCountReader.GetInt32(1); };
                            if (!TestCountReader.IsDBNull(2)) { testconsultCount = TestCountReader.GetInt32(2); };
                            if (!TestCountReader.IsDBNull(3)) { testruspoleCount = TestCountReader.GetInt32(3); };
                        }
                    }
                    else { }
                    TestCountReader.Close();

                    Connection.Close();
                }
            }
            catch (Exception ex) 
            {
                Console.WriteLine(ex);
            }
          
            return new ReportRow { rowNumber = testNumber, rowName = testName, hspColumn = testhspCount, expressColumn = testexpressCount, consultColumn = testconsultCount, ruspoleColumn = testruspoleCount };
        }

        // получение общего количество для тестов из массива
        public ReportRow GetDataForArrayFromDB(int number, string[] testCodesArray, DateTime validationFrom, DateTime validationTo)
        {
            string testName = "";
            //string testName = null;
            string testNumber = number.ToString();
            int testhspCount = 0;
            int testexpressCount = 0;
            int testconsultCount = 0;
            int testruspoleCount = 0;

            string testCodePar = testCodesArray[0]; // имя первого теста из массива, по нему устанавливаем название строки
            Console.WriteLine(testCodePar);

            testName = GetRowName(testCodePar); // есть ли название строки отчета, которое не совпадает с именем теста

            try
            {
                // Build a configuration object from JSON file
                IConfiguration config = new ConfigurationBuilder()
                    .AddJsonFile("AppConfig.json")
                    .Build();

                // Get a configuration section
                IConfigurationSection section = config.GetSection("ConnnectionStrings");

                string? CGMConnectionString = section["CGMConnection"];
                CGMConnectionString = string.Concat(CGMConnectionString, $"User Id = {user}; Password = {password}");

                using (SqlConnection Connection = new SqlConnection(CGMConnectionString))
                {
                    Connection.Open();

                    //string parameters_ = string.Join(",", testCodesArray);
                    string parameters_ = string.Join(",", testCodesArray.Select(n => "'" + n + "'"));


                    if (testName == "")
                    {
                        // получаем название теста
                        //SqlCommand TestNameCommand = new SqlCommand($"SELECT a.ana_analys FROM KDLPROD..ana a WHERE a.ana_analyskod = '{testCodePar}'", Connection);
                        SqlCommand TestNameCommand = new SqlCommand($"SELECT a.ana_analys FROM KDLPROD..ana a WHERE a.ana_analyskod = @testcode", Connection);
                        // создаем параметр для имени
                        SqlParameter testParam = new SqlParameter("@testcode", testCodePar);
                        // добавляем параметр к команде
                        TestNameCommand.Parameters.Add(testParam);

                        SqlDataReader Reader = TestNameCommand.ExecuteReader();

                        if (Reader.HasRows)
                        {
                            while (Reader.Read())
                            {
                                if (!Reader.IsDBNull(0)) { testName = Reader.GetString(0); };
                            }
                        }
                        else { }
                        Reader.Close();
                    }


                    // получаем количество тестов

                    SqlCommand TestCountCommand = new SqlCommand(
                        "SELECT  t1.hospital, t2.express, t3.consult, t4.ruspole FROM " +
                            "(SELECT ROW_NUMBER() OVER (ORDER BY (select 1)) AS RowNum, COUNT(*) AS hospital " +
                            "FROM KDLReportView kv " +
                                $"WHERE kv.test_code IN ({@parameters_}) " +
                                       "AND kv.ValidationDate >= @validationDateFrom AND kv.ValidationDate <= @validationDateTo " +
                                       "AND kv.OperatorId NOT IN ('ATA', 'ITA', 'МАВ', 'РЮВ', 'ПИН', 'АНА', 'ТВИ', 'ШМИ', 'ПАХ', 'ГПГ', 'ПАВ', 'ПДМ', 'ОНВ', 'РМВ', 'СЕА', 'ССО', 'RSIG') " +
                                       "AND kv.ClientCode NOT IN ('42','56', '57', '58', '59', '60', '61', '63')) t1 " +
                            "JOIN " +
                            "(SELECT ROW_NUMBER() OVER (ORDER BY (select 1)) AS RowNum, COUNT(*) AS express " +
                            "FROM KDLReportView kv " +
                                $"WHERE kv.test_code IN ({@parameters_}) " +
                                       "AND kv.ValidationDate >=  @validationDateFrom AND kv.ValidationDate <= @validationDateTo " +
                                       "AND kv.OperatorId IN ('РУМ', 'ПОВ', 'БЛН', 'КОВ', 'KOV', 'ККА')) t2 ON t1.RowNum = t2.RowNum " +
                            "JOIN " +
                            "(SELECT ROW_NUMBER() OVER (ORDER BY (select 1)) AS RowNum, COUNT(*) AS consult " +
                            "FROM KDLReportView kv " +
                                $"WHERE kv.test_code IN ({@parameters_})  " +
                                       "AND kv.ValidationDate >= @validationDateFrom AND kv.ValidationDate <= @validationDateTo " +
                                       "AND kv.OperatorId NOT IN ('ATA', 'ITA', 'МАВ', 'РЮВ', 'ПИН', 'АНА', 'ТВИ', 'ШМИ', 'ПАХ', 'ГПГ', 'ПАВ', 'ПДМ', 'ОНВ', 'РМВ', 'СЕА', 'ССО', 'RSIG') " +
                                       "AND kv.ClientCode = 42) t3 ON t1.RowNum = t3.RowNum " +
                            "JOIN " +
                            "(SELECT ROW_NUMBER() OVER (ORDER BY (select 1)) AS RowNum, COUNT(*) AS ruspole " +
                            "FROM KDLReportView kv " +
                                $"WHERE kv.test_code IN ({@parameters_}) " +
                                "AND kv.ValidationDate >= @validationDateFrom AND kv.ValidationDate <= @validationDateTo " +
                                "AND kv.analyzer NOT IN ('SAPPHIRE', 'FUS100', 'MIN6200', 'SAPPHIR') " +
                                "AND kv.ClientCode IN ('56', '57', '58', '59', '60', '61', '63') " +
                                "AND kv.OperatorId NOT IN ('ATA', 'ITA', 'МАВ', 'РЮВ', 'ПИН', 'АНА', 'ТВИ', 'ШМИ', 'ПАХ', 'ГПГ', 'ПАВ', 'ПДМ', 'ОНВ', 'РМВ', 'СЕА', 'ССО', 'RSIG')) t4 ON t1.RowNum = t4.RowNum ", Connection);

                    // создаем параметр для имени
                    //SqlParameter testcodeParam = new SqlParameter("@parameters_", parameters_);
                    //TestCountCommand.Parameters.Add(testcodeParam);
                    // создаем параметры для дат валидации
                    SqlParameter validationFromParam = new SqlParameter("@validationDateFrom", validationFrom);
                    TestCountCommand.Parameters.Add(validationFromParam);
                    SqlParameter validationToParam = new SqlParameter("@validationDateTo", validationTo);
                    TestCountCommand.Parameters.Add(validationToParam);

                    SqlDataReader TestCountReader = TestCountCommand.ExecuteReader();

                    if (TestCountReader.HasRows)
                    {
                        while (TestCountReader.Read())
                        {
                            if (!TestCountReader.IsDBNull(0)) { testhspCount = TestCountReader.GetInt32(0); };
                            if (!TestCountReader.IsDBNull(1)) { testexpressCount = TestCountReader.GetInt32(1); };
                            if (!TestCountReader.IsDBNull(2)) { testconsultCount = TestCountReader.GetInt32(2); };
                            if (!TestCountReader.IsDBNull(3)) { testruspoleCount = TestCountReader.GetInt32(3); };
                        }
                    }
                    else { }
                    TestCountReader.Close();

                    Connection.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }

            Console.WriteLine(testNumber);
            return new ReportRow { rowNumber = testNumber, rowName = testName, hspColumn = testhspCount, expressColumn = testexpressCount, consultColumn = testconsultCount, ruspoleColumn = testruspoleCount };
        }

        // получение общего количество для тестов из массива для Итогов
        public ReportRow GetTotalsCountFromDB(string rowNum, string[] testCodesArray, DateTime validationFrom, DateTime validationTo)
        {
            //string testName = "";
            string testNumber = rowNum;
            int testhspCount = 0;
            int testexpressCount = 0;
            int testconsultCount = 0;
            int testruspoleCount = 0;

            string testName = GetTotalName(rowNum);
            Console.WriteLine(testName);

            try
            {
                // Build a configuration object from JSON file
                IConfiguration config = new ConfigurationBuilder()
                    .AddJsonFile("AppConfig.json")
                    .Build();

                // Get a configuration section
                IConfigurationSection section = config.GetSection("ConnnectionStrings");

                string? CGMConnectionString = section["CGMConnection"];
                CGMConnectionString = string.Concat(CGMConnectionString, $"User Id = {user}; Password = {password}");

                using (SqlConnection Connection = new SqlConnection(CGMConnectionString))
                {
                    Connection.Open();

                    string parametersValue = string.Join(",", testCodesArray.Select(n => "'" + n + "'"));
                    Console.WriteLine($"STRING PARAMS: {parametersValue}");


                    // получаем общее количество тестов
                    SqlCommand TestCountCommand = new SqlCommand(
                        "SELECT  t1.hospital, t2.express, t3.consult, t4.ruspole FROM " +
                            "(SELECT ROW_NUMBER() OVER (ORDER BY (select 1)) AS RowNum, COUNT(*) AS hospital " +
                            "FROM KDLReportView kv " +
                                $"WHERE kv.test_code IN ({@parametersValue}) " +
                                       "AND kv.ValidationDate >= @validationDateFrom AND kv.ValidationDate <= @validationDateTo " +
                                       "AND kv.OperatorId NOT IN ('ATA', 'ITA', 'МАВ', 'РЮВ', 'ПИН', 'АНА', 'ТВИ', 'ШМИ', 'ПАХ', 'ГПГ', 'ПАВ', 'ПДМ', 'ОНВ', 'РМВ', 'СЕА', 'ССО', 'RSIG') " +
                                       "AND kv.ClientCode NOT IN ('42','56', '57', '58', '59', '60', '61', '63')) t1 " +
                            "JOIN " +
                            "(SELECT ROW_NUMBER() OVER (ORDER BY (select 1)) AS RowNum, COUNT(*) AS express " +
                            "FROM KDLReportView kv " +
                                $"WHERE kv.test_code IN ({@parametersValue}) " +
                                       "AND kv.ValidationDate >=  @validationDateFrom AND kv.ValidationDate <= @validationDateTo " +
                                       "AND kv.OperatorId IN ('РУМ', 'ПОВ', 'БЛН', 'КОВ', 'KOV', 'ККА')) t2 ON t1.RowNum = t2.RowNum " +
                            "JOIN " +
                            "(SELECT ROW_NUMBER() OVER (ORDER BY (select 1)) AS RowNum, COUNT(*) AS consult " +
                            "FROM KDLReportView kv " +
                                $"WHERE kv.test_code IN ({@parametersValue}) " +
                                       "AND kv.ValidationDate >= @validationDateFrom AND kv.ValidationDate <= @validationDateTo " +
                                       "AND kv.OperatorId NOT IN ('ATA', 'ITA', 'МАВ', 'РЮВ', 'ПИН', 'АНА', 'ТВИ', 'ШМИ', 'ПАХ', 'ГПГ', 'ПАВ', 'ПДМ', 'ОНВ', 'РМВ', 'СЕА', 'ССО', 'RSIG') " +
                                       "AND kv.ClientCode = 42) t3 ON t1.RowNum = t3.RowNum " +
                            "JOIN " +
                            "(SELECT ROW_NUMBER() OVER (ORDER BY (select 1)) AS RowNum, COUNT(*) AS ruspole " +
                            "FROM KDLReportView kv " +
                                $"WHERE kv.test_code IN ({@parametersValue}) " +
                                "AND kv.ValidationDate >= @validationDateFrom AND kv.ValidationDate <= @validationDateTo " +
                                "AND kv.analyzer NOT IN ('SAPPHIRE', 'FUS100', 'MIN6200', 'SAPPHIR') " +
                                "AND kv.ClientCode IN ('56', '57', '58', '59', '60', '61', '63') " +
                                "AND kv.OperatorId NOT IN ('ATA', 'ITA', 'МАВ', 'РЮВ', 'ПИН', 'АНА', 'ТВИ', 'ШМИ', 'ПАХ', 'ГПГ', 'ПАВ', 'ПДМ', 'ОНВ', 'РМВ', 'СЕА', 'ССО', 'RSIG')) t4 ON t1.RowNum = t4.RowNum ", Connection);

                    // создаем параметры для дат валидации
                    SqlParameter validationFromParam = new SqlParameter("@validationDateFrom", validationFrom);
                    TestCountCommand.Parameters.Add(validationFromParam);
                    SqlParameter validationToParam = new SqlParameter("@validationDateTo", validationTo);
                    TestCountCommand.Parameters.Add(validationToParam);

                    //Console.WriteLine(TestCountCommand.CommandText.ToString());

                    SqlDataReader TestCountReader = TestCountCommand.ExecuteReader();

                    if (TestCountReader.HasRows)
                    {
                        while (TestCountReader.Read())
                        {
                            if (!TestCountReader.IsDBNull(0)) { testhspCount = TestCountReader.GetInt32(0); Console.WriteLine($"COUNT HSP: {testhspCount}"); };
                            if (!TestCountReader.IsDBNull(1)) { testexpressCount = TestCountReader.GetInt32(1); Console.WriteLine($"COUNT EXPR: {testexpressCount}"); };
                            if (!TestCountReader.IsDBNull(2)) { testconsultCount = TestCountReader.GetInt32(2); Console.WriteLine($"COUNT CNSLT: {testconsultCount}"); };
                            if (!TestCountReader.IsDBNull(3)) { testruspoleCount = TestCountReader.GetInt32(3); Console.WriteLine($"COUNT POLE: {testruspoleCount}"); };
                        }
                    }
                    else 
                    {
                        Console.WriteLine("NO ROWS");
                    }
                    TestCountReader.Close();

                    Connection.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }

            return new ReportRow { rowNumber = testNumber, rowName = testName, hspColumn = testhspCount, expressColumn = testexpressCount, consultColumn = testconsultCount, ruspoleColumn = testruspoleCount };
        }

        // получение списка тестов, которые были выполнены за заданный период
        public string[] GetTestsArray(string discipline, DateTime validationFrom, DateTime validationTo) 
        {
            List<string> tests = new List<string>();

            try
            {
                // Build a configuration object from JSON file
                IConfiguration config = new ConfigurationBuilder()
                    .AddJsonFile("AppConfig.json")
                    .Build();

                // Get a configuration section
                IConfigurationSection section = config.GetSection("ConnnectionStrings");

                string? CGMConnectionString = section["CGMConnection"];
                CGMConnectionString = string.Concat(CGMConnectionString, $"User Id = {user}; Password = {password}");

                using (SqlConnection Connection = new SqlConnection(CGMConnectionString))
                {
                    Connection.Open();

                    SqlCommand TestListCommand = new SqlCommand("SELECT DISTINCT(kv.test_code) FROM KDLReportView kv WHERE kv.discipline = @discipline " +
                                                                "AND kv.ValidationDate >= @validationDateFrom AND kv.ValidationDate <= @validationDateTo ", Connection);

                    SqlParameter disciplineParam = new SqlParameter("@discipline", discipline);
                    TestListCommand.Parameters.Add(disciplineParam);
                    SqlParameter validationFromParam = new SqlParameter("@validationDateFrom", validationFrom);
                    TestListCommand.Parameters.Add(validationFromParam);
                    SqlParameter validationToParam = new SqlParameter("@validationDateTo", validationTo);
                    TestListCommand.Parameters.Add(validationToParam);

                    SqlDataReader TestListReader = TestListCommand.ExecuteReader();

                    if (TestListReader.HasRows)
                    {
                        while (TestListReader.Read())
                        {
                            if (!TestListReader.IsDBNull(0)) 
                            {
                                tests.Add(TestListReader.GetString(0)); 
                            }
                        }
                    }
                }
            }
            catch
            {

            }

            string[] testsArray = tests.ToArray();
            return testsArray;
        }

        // получение значения теста Количество препаратов для пунктата ЩЖ (кол-во препаратов в комментарии к заявке)
        public ReportRow GetPreparatCountFromDB()
        {
            try
            {
                // Build a configuration object from JSON file
                IConfiguration config = new ConfigurationBuilder()
                    .AddJsonFile("AppConfig.json")
                    .Build();

                // Get a configuration section
                IConfigurationSection section = config.GetSection("ConnnectionStrings");

                string? CGMConnectionString = section["CGMConnection"];
                CGMConnectionString = string.Concat(CGMConnectionString, $"User Id = {user}; Password = {password}");
            }
            catch
            {

            }

            return 
        }


        // получение количества препаратов
        // public ReportRow GetPreparatsCountFromDB(string test_code, DateTime validationFrom, DateTime validationTo)
        //{
        //
        // }

        /*
        public ReportRow GetSingleTotal(string rowNum, List<string> testCodesArray, DateTime validationFrom, DateTime validationTo)
        {
            //string testName = "";
            string testNumber = rowNum;
            int testhspCount = 0;
            int testexpressCount = 0;
            int testconsultCount = 0;
            int testruspoleCount = 0;

            string testName = GetTotalName(rowNum);
            Console.WriteLine(testName);

            try
            {
                // Build a configuration object from JSON file
                IConfiguration config = new ConfigurationBuilder()
                    .AddJsonFile("AppConfig.json")
                    .Build();

                // Get a configuration section
                IConfigurationSection section = config.GetSection("ConnnectionStrings");

                string? CGMConnectionString = section["CGMConnection"];
                CGMConnectionString = string.Concat(CGMConnectionString, $"User Id = {user}; Password = {password}");

                using (SqlConnection Connection = new SqlConnection(CGMConnectionString))
                {
                    Connection.Open();

                    //string parameters_ = string.Join(",", testCodesArray);
                    //Console.WriteLine($"PARAMS: {parameters_}");

                    string parameterValue = string.Join(",", testCodesArray.Select(n => "'" + n + "'"));
                    Console.WriteLine($"STRING PARAMS: {parameterValue}");

                    //string parameterValue = "'Б0001', 'Б0150'";

                    // получаем общее количество тестов
                    SqlCommand TestCountCommand = new SqlCommand(
                        "SELECT  t1.hospital FROM " +
                            "(SELECT ROW_NUMBER() OVER (ORDER BY (select 1)) AS RowNum, COUNT(*) AS hospital " +
                            "FROM KDLReportView kv " +
                                $"WHERE kv.test_code IN ({parameterValue}) " +
                                       "AND kv.ValidationDate >= @validationDateFrom AND kv.ValidationDate <= @validationDateTo " +
                                       "AND kv.OperatorId NOT IN ('ATA', 'ITA', 'МАВ', 'РЮВ', 'ПИН', 'АНА', 'ТВИ', 'ШМИ', 'ГПГ', 'ПАВ', 'ПДМ', 'ОНВ', 'РМВ', 'СЕА', 'ССО', 'RSIG') " +
                                       "AND kv.ClientCode NOT IN ('56', '57', '58', '59', '60', '61', '63')) t1 " , Connection);

                    // создаем параметр для имени
                    //SqlParameter testcodeParam = new SqlParameter("@testsParam", parameterValue);
                    //SqlParameter testcodeParam = new SqlParameter("@parameters_", parameterValue);
                    //Console.WriteLine($"PAR value {testcodeParam.Value}");
                    //TestCountCommand.Parameters.Add(testcodeParam);


 

                    //SqlParameter parameter = new SqlParameter("@testsParam", SqlDbType.NVarChar);
                    //parameter.Value = parameterValue;
                    //Console.WriteLine($"PARAMETER: {parameter.Value.ToString()}");
                    //TestCountCommand.Parameters.Add(parameter);
                    

                    // создаем параметры для дат валидации
                    SqlParameter validationFromParam = new SqlParameter("@validationDateFrom", validationFrom);
                    TestCountCommand.Parameters.Add(validationFromParam);
                    SqlParameter validationToParam = new SqlParameter("@validationDateTo", validationTo);
                    TestCountCommand.Parameters.Add(validationToParam);

                    Console.WriteLine(TestCountCommand.CommandText.ToString());

                    SqlDataReader TestCountReader = TestCountCommand.ExecuteReader();

                    if (TestCountReader.HasRows)
                    {
                        while (TestCountReader.Read())
                        {
                            if (!TestCountReader.IsDBNull(0)) { testhspCount = TestCountReader.GetInt32(0); Console.WriteLine($"COUNT HSP: {testhspCount}"); };
                            //if (!TestCountReader.IsDBNull(1)) { testexpressCount = TestCountReader.GetInt32(1); Console.WriteLine($"COUNT EXPR: {testexpressCount}"); };
                            //if (!TestCountReader.IsDBNull(2)) { testconsultCount = TestCountReader.GetInt32(2); Console.WriteLine($"COUNT CNSLT: {testconsultCount}"); };
                            //if (!TestCountReader.IsDBNull(3)) { testruspoleCount = TestCountReader.GetInt32(3); Console.WriteLine($"COUNT POLE: {testruspoleCount}"); };
                        }
                    }
                    else
                    {
                        Console.WriteLine("NO ROWS");
                    }
                    TestCountReader.Close();

                    Connection.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }

            return new ReportRow { rowNumber = testNumber, rowName = testName, hspColumn = testhspCount, expressColumn = testexpressCount, consultColumn = testconsultCount, ruspoleColumn = testruspoleCount };

        }
        */

        #endregion
    }
}
