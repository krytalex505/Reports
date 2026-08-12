using Reports.models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Windows;

namespace Reports.service
{
    class DaoService
    {
        private readonly SqlConnection connection;
        public string date;
        public string date1;
        public string date2;
        public string Month;
        public string Year;



        public DaoService()
        {
            connection = new SqlConnection("Data Source=86.57.137.8,1433;Initial Catalog=ap6pay;Persist Security Info=True;User ID=admin;Password=682830;");
            connection.Open();
        }
        private List<Pay> QueryList(string sql)
        {
            List<Pay> pays = new List<Pay>();
            SqlCommand cmd = new SqlCommand(sql, connection);
            SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                pays.Add(Pay.Parse(reader));
            }
            
            reader.Close();
            return pays;
           
        }

        //public List<Pay> GetDataForASpecificDateForShop()
        //{
        //    return QueryList("SELECT Id, tab_no, fio, summa, val, val1, type, date FROM dbo.dobor_" + Month + "_" + Year + " Where date between '" + date1 + "' And '" + date2 + "' And type = '518'");
        //}

        //public List<Pay> GetDataForASpecificDateForKitchen()
        //{
        //    return QueryList("SELECT Id, tab_no, fio, summa, val, val1, type, date FROM dbo.dobor_" + Month + "_" + Year + " Where date between '" + date1 + "' And '" + date2 + "' And type = '521'");
        //}

        //public List<Pay> GetDataForASpecificDateGeneral()
        //{
        //    return QueryList("SELECT Id, tab_no, fio, summa, val, val1, type, date FROM dbo.dobor_" + Month + "_" + Year + " Where date between '" + date1 + "' And '" + date2 + "'");
        //}

        public List<Pay> GetPaysOfMonth2()
        {
            return QueryList("SELECT Id, tab_no, fio, summa, val, val1, type, date FROM dbo.dobor_" + Month + "_" + Year + "  Where date between '" + date1 + "' And '" + date2 + "' ORDER BY type ");
        }

        public List<Pay> GetPaysOfMonth(DateTime dateTime)
        {
          return QueryList("SELECT Id, tab_no, fio, summa, val, val1, type, date FROM dbo.dobor_" + dateTime.Month + "_" + dateTime.Year + " ORDER BY type");
        }

        public void CloseConnection()
        {
            connection.Close();
        }
          

        public void ExportDataToDbf(List<Pay> pays, string filePath, string fileName, Action performStep)
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-EN");
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=dBASE IV;User ID=admin;Password=;";
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            using (OleDbCommand command = connection.CreateCommand())
            {
                connection.Open();
                command.CommandText = string.Format("CREATE TABLE {0} (Tabel integer, SUMMA integer, KOD integer, DAT varchar(21))", fileName); //, string.Join(",", columns)
                command.ExecuteNonQuery();
               DateTime dateTime = DateTime.Now.AddDays(-1);
                foreach (Pay pay in pays)
                {
                    if (/*dateTime.ToString("dd.MM.yyyy")*/date == pay.Date.ToString("dd.MM.yyyy"))
                    {
                        command.CommandText = string.Format("insert into {0} values {1} ",
                                              fileName,
                                              string.Format("({0},{1},{2},'{3}') ", pay.TabNom, pay.Price, pay.Type, pay.Date.ToString("dd.MM.yyyy")));
                        command.ExecuteNonQuery();
                        performStep.Invoke();
                    }

                }
            }
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("ru-RU");
        }
    }
}
