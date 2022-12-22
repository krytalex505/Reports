using Reports.models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace Reports.service
{
    class DaoService
    {
        private readonly SqlConnection connection;

        public DaoService()
        {
            connection = new SqlConnection("Data Source=86.57.137.8,1433;Initial Catalog=ap6pay;Persist Security Info=True;User ID=admin;Password=682830");
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
            return pays;
        }

        public List<Pay> GetPaysOfMonth(DateTime dateTime)
        {
            return QueryList("SELECT Id, tab_no, fio, summa, val, val1, type, date FROM dbo.dobor_" + dateTime.Month + "_" + dateTime.Year);
        }

        public void CloseConnection()
        {
            connection.Close();
        }

        public void ExportDataToDbf(List<Pay> pays, string filePath, string fileName, Action performStep)
        {
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=dBASE IV;User ID=admin;Password=;";
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            using (OleDbCommand command = connection.CreateCommand())
            {
                connection.Open();
                command.CommandText = string.Format("CREATE TABLE {0} (Tabel integer, SUMMA integer, KOD integer, DAT varchar(21))", fileName); //, string.Join(",", columns)
                command.ExecuteNonQuery();

                foreach (Pay pay in pays)
                {
                    command.CommandText = string.Format("insert into {0} values {1}", 
                        fileName, 
                        string.Format("({0},{1},{2},'{3}')", pay.TabNom, pay.Price, pay.Type, pay.Date.ToString("dd.MM.yyyy")));
                    command.ExecuteNonQuery();
                    performStep.Invoke();
                }
            }
        }
    }
}
