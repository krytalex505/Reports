using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Text;

namespace Reports.models
{
    class Pay
    {
        //Id, tab_no As Табельный,fio As ФИО,summa As Сумма,val As Собственная,val1 As Готовая,type As Код,dat

        public long Id { get; set; }
        [System.ComponentModel.DisplayName("Табельный номер")]
        public int TabNom { get; set; }
        [System.ComponentModel.DisplayName("ФИО")]
        public string Name { get; set; }
        [System.ComponentModel.DisplayName("Сумма")]
        public double Price { get; set; }
        [System.ComponentModel.DisplayName("Собственная")]
        public double? Val { get; set; }
        [System.ComponentModel.DisplayName("Готовая")]
        public double? Val1 { get; set; }
        [System.ComponentModel.DisplayName("Код")]
        public int Type { get; set; }
        [System.ComponentModel.DisplayName("Дата")]
        public DateTime Date { get; set; }

        public static Pay Parse(SqlDataReader reader)
        {
            Pay pay = new Pay();
            pay.Id = pay.ParseValue(reader["id"], val => long.Parse(val));
            pay.TabNom = pay.ParseValue(reader["tab_no"], val => int.Parse(val));
            pay.Name = reader["fio"].ToString();
            pay.Price = pay.ParseValue(reader["summa"], val => double.Parse(val, new CultureInfo("en-us")));
            pay.Val = pay.ParseValue(reader["val"], val => double.Parse(val, new CultureInfo("en-us")));
            pay.Val1 = pay.ParseValue(reader["val1"], val => double.Parse(val, new CultureInfo("en-us")));
            pay.Type = pay.ParseValue(reader["type"], val => int.Parse(val));
            pay.Date = pay.ParseValue(reader["date"], val => DateTime.Parse(val));
            return pay;
        }
        private T ParseValue<T>(object val, Func<string, T> parser)
        {
            return string.IsNullOrEmpty(val.ToString()) ? default(T) : parser.Invoke(val.ToString());
        }
    }
}
