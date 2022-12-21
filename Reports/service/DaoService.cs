using Reports.models;
using System;
using System.Collections.Generic;
using System.Data;
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
    }
}
