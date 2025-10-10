using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data.SqlClient;

namespace _30Form
{
    public class ReportRow
    {
        public string rowName { get; set; }
        public int hspColumn { get; set; }
        public int expressColumn { get; set; }
        public int consultColumn { get; set; }
        public int ruspoleColumn { get; set; }
    }

    public class Reporter
    {
        string user = "mielogrammauser";
        string password = "Qw123456";

        string[] biochemicalTests = new string[] { "Б0001", "Б0005" };
        string[] gematologyTests = new string[] {  };
        string[] coagulogramTests = new string[] { };

        public List<ReportRow> GetReport()
        {
            var rows = new List<ReportRow>();

            rows.Add(new ReportRow
            {
                rowName = "Общий белок в сыворотке крови",
                hspColumn = 10,
                expressColumn = 20,
                consultColumn = 30,
                ruspoleColumn = 40
            });

            rows.Add(new ReportRow
            {
                rowName = "альбумин в сыворотке крови",
                hspColumn = 1,
                expressColumn = 2,
                consultColumn = 3,
                ruspoleColumn = 4
            });



            return rows;
        }

        #region SQL скрипты
        public void GetDataFromDB()
        {

            try
            {
                string CGMConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings["CGMConnection"].ConnectionString;
                CGMConnectionString = String.Concat(CGMConnectionString, $"User Id = {user}; Password = {password}");

                using (SqlConnection Connection = new SqlConnection(CGMConnectionString))
                {
                    Connection.Open();

                    //var parameters = new List<SqlParameter>();
                    //var parameterNames = new List<string>();

                    string parameters_ = string.Join(",", biochemicalTests);


                }
            }
            catch (Exception ex) 
            {

            }
            
        }

        #endregion
    }
}
