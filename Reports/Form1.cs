using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.IO;
using Spire.DataExport.DBF;
using System.Data.OleDb;
using Reports.models;
using Reports.service;
using System.ComponentModel;
using System.Threading.Tasks;

namespace Reports
{
    public partial class Form1 : Form
    {
        
        private readonly List<Panel> _addpanels2;
        //private int _progress;
        //private delegate void Delegate();
        private IList<Pay> pays = new SortableBindingList<Pay>();
        private DaoService daoService;
        Excel.Application xlApp;
        Excel.Worksheet xlSheet;
        Excel.Range xlSheetRange;
        public SqlConnection con; 
        string stroka = "Data Source=86.57.137.8,1433;Initial Catalog=ap6pay;Persist Security Info=True;User ID=admin;Password=682830";
        public DataTable dataTable;
        //public string date = DateTime.UtcNow.ToString("dd.MM.yyyy 23:59");
        private DateTime actualDate;
        public Form1()
        {
            
            con = new SqlConnection(stroka);
            //con.Open();
            InitializeComponent();

            //Width = Screen.PrimaryScreen.Bounds.Width;
            //Height = Screen.PrimaryScreen.Bounds.Height;
            daoService = new DaoService();
            _addpanels2 = new List<Panel> { panel1, panel3, panel4 };
            Application.ApplicationExit += new EventHandler(this.OnApplicationExit);
            Ap6Pay();
            //Ap6Pay();
            Sum();

        }
        public void Ap6Pay3()
        {

            dataGridView1.AutoGenerateColumns = true;
            dataGridView1.DataSource = new BindingSource(pays, null);
            dataGridView1.Sort(dataGridView1.Columns[0], ListSortDirection.Ascending);
            foreach (var pay in daoService.GetPaysOfMonth2())
                pays.Add(pay);
            //System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("ru-Ru");
            dataGridView1.Columns["Id"].Visible = false;
        }
        public void Ap6Pay2()
        {
      
            dataGridView1.AutoGenerateColumns = true;
            dataGridView1.DataSource = new BindingSource(pays, null);
            dataGridView1.Sort(dataGridView1.Columns[0], ListSortDirection.Ascending);
            foreach (var pay in daoService.GetPaysOfMonth(DateTime.Now))
                pays.Add(pay);
            //System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("ru-Ru");
            dataGridView1.Columns["Id"].Visible = false;
        }

        private void OnApplicationExit(object sender, EventArgs e)
        {
            daoService.CloseConnection();
        }

        private void AddColumn(string name, string headerText)
        {
            DataGridViewColumn col = new DataGridViewColumn();
            col.HeaderText = headerText;
            col.Name = name;
            dataGridView1.Columns.Add(col);
        }
        public void GetDataTabell78()
        {
            SqlDataAdapter ada = new SqlDataAdapter("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'dobor_" + dateTimePicker2.Value.Month + "_" + dateTimePicker2.Value.Year + "'", con);
            DataTable dt = new DataTable();
            ada.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                string date6 = dateTimePicker2.Value.ToString("dd.MM.yyyy 00:00");
                string date7 = dateTimePicker2.Value.ToString("dd.MM.yyyy 23:59");
                string command = "SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма, val As Собственная,val1 As Готовая,type As Тип,date As Дата FROM dbo.dobor_" + dateTimePicker2.Value.Month + "_" + dateTimePicker2.Value.Year + " Where  tab_no='" + maskedTextBox1.Text + "' And date between '" + date6 + "' And '" + date7 + "'";
                SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command, con);
                DataSet dataSet = new DataSet();
                sqlDataAdapter.Fill(dataSet);
                dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
                dataGridView1.Columns[0].Visible = false;
            }
            else
            {
                MessageBox.Show("Данных за выбранный период не существует!");
                maskedTextBox1.Clear();
            }

        }

        public void GetDataTabell()
        {
            SqlDataAdapter ada = new SqlDataAdapter("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'dobor_" + dateTimePicker2.Value.Month + "_" + dateTimePicker2.Value.Year + "'", con);
            DataTable dt = new DataTable();
            ada.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                string command = "SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма, val As Собственная,val1 As Готовая,type As Код,date As Дата FROM dbo.dobor_" + dateTimePicker2.Value.Month + "_" + dateTimePicker2.Value.Year + " Where  tab_no='" + maskedTextBox1.Text + "'";
                SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command, con);
                DataSet dataSet = new DataSet();
                sqlDataAdapter.Fill(dataSet);
                dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
            }
             else
             {
                    MessageBox.Show("Данных за выбранный период не существует!");
                    maskedTextBox1.Clear();
            }



        }
        public void GetDataForPerDateKitchen()
        {
            string date1 = dateTimePicker1.Value.ToString("dd.MM.yyyy 00:00");
            string date2 = dateTimePicker1.Value.ToString("dd.MM.yyyy 23:59");
            // зачем выгружать из sql, если есть таблица
            string command = "SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма, val As Собственная,val1 As Готовая,type As Код, date as Дата FROM dbo.dobor_" + dateTimePicker1.Value.Month + "_" + dateTimePicker1.Value.Year + " Where date between '" + date1 + "' And '" + date2 + "' And Type = 521";
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command, con);
            DataSet dataSet = new DataSet();
            sqlDataAdapter.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
            dataGridView1.Columns[0].Visible = false;
        }
        public void GetDataForPerDateShop()
        {
            string date1 = dateTimePicker1.Value.ToString("dd.MM.yyyy 00:00");
            string date2 = dateTimePicker1.Value.ToString("dd.MM.yyyy 23:59");
            // зачем выгружать из sql, если есть таблица
            string command = "SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма, val As Собственная,val1 As Готовая, type As Код, date as Дата FROM dbo.dobor_" + dateTimePicker1.Value.Month + "_" + dateTimePicker1.Value.Year + " Where date between '" + date1 + "' And '" + date2 + "' And Type = 518";
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command, con);
            DataSet dataSet = new DataSet();
            sqlDataAdapter.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
            dataGridView1.Columns[0].Visible = false;
        }
        public void GetDataForPerDate()
        {
            string date1 = dateTimePicker1.Value.ToString("dd.MM.yyyy 00:00");
            string date2 = dateTimePicker1.Value.ToString("dd.MM.yyyy 23:59");
            // зачем выгружать из sql, если есть таблица
            string command = "SELECT Id, tab_no As Табельный,fio As ФИО ,summa As Сумма,val As Собственная,val1 As Готовая,type As Код,date As Дата FROM dbo.dobor_" + dateTimePicker1.Value.Month + "_" + dateTimePicker1.Value.Year + " Where date between '" + date1 + "' And '" + date2 + "'";
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command, con);
            DataSet dataSet = new DataSet();
            sqlDataAdapter.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
            dataGridView1.Columns[0].Visible = false;
        }
        public void GetDataForMonthKitchen()
        {

            string command = "SELECT Id, tab_no As Табельный,fio As ФИО ,summa As Сумма,val As Собственная,val1 As Готовая,type As Код , date As Дата FROM dbo.dobor_" + dateTimePicker4.Value.Month + "_" + dateTimePicker4.Value.Year + " Where type = 521";
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command, con);
            DataSet dataSet = new DataSet();
            sqlDataAdapter.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
            dataGridView1.Columns[0].Visible = false;

        }


        public void GetDataForMonthhShop()
        {

            string command = "SELECT Id, tab_no As Табельный,fio As ФИО ,summa As Сумма,val As Собственная,val1 As Готовая,type As Код,date As Дата FROM dbo.dobor_" + dateTimePicker4.Value.Month + "_" + dateTimePicker4.Value.Year + " Where type = 518";
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command, con);
            DataSet dataSet = new DataSet();
            sqlDataAdapter.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
            dataGridView1.Columns[0].Visible = false;
        }
        public void GetDataForMonthh()
        {

            string command = "SELECT Id, tab_no As Табельный,fio As ФИО ,summa As Сумма,val As Собственная,val1 As Готовая,type As Код,date As Дата FROM dbo.dobor_" + dateTimePicker4.Value.Month + "_" + dateTimePicker4.Value.Year + "";
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command, con);
            DataSet dataSet = new DataSet();
            sqlDataAdapter.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
            dataGridView1.Columns[0].Visible = false;
        }


        //public void Ap6Pay78778()
        //{

        //    SqlDataAdapter ada = new SqlDataAdapter("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'dobor_" + DateTime.Now.Month + "_" + DateTime.Now.Year + "'", con);
        //    DataTable dt = new DataTable();
        //    ada.Fill(dt);
        //    if (dt.Rows.Count > 0)
        //    {
        //        string date9 = DateTime.UtcNow.ToString("dd.MM.yyyy 00:00");
        //        string date10 = DateTime.UtcNow.ToString("dd.MM.yyyy 23:59");

        //        string command = "SELECT tab_no ,summa,type,date FROM dbo.dobor_" + DateTime.Now.Month + "_" + DateTime.Now.Year + " where date between '" + date9 + "' and '" + date10 + "'";
        //        SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command, con);
        //        DataSet dataSet = new DataSet();
        //        sqlDataAdapter.Fill(dataSet);
        //        dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
        //    }
        //    else
        //    {
        //        MessageBox.Show("Данных за текущий месяц не существует!");
        //    }

        //}


        public void Ap6Pay()
        {
            actualDate = DateTime.Now;

            //if(actualDate.Year == date.Year && mouth)
            //sql

            SqlDataAdapter ada = new SqlDataAdapter("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'dobor_" + DateTime.Now.Month + "_" + DateTime.Now.Year + "'", con);
            DataTable dt = new DataTable();
            ada.Fill(dt);
            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("Данных за текущий месяц не существует!");
                return;
            }
            // переделать
            string command = "SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма,val As Собственная,val1 As Готовая,type As Код,date As Дата FROM dbo.dobor_" + DateTime.Now.Month + "_" + DateTime.Now.Year + "";
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command, con);
            DataSet dataSet = new DataSet();
            sqlDataAdapter.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
            dataGridView1.Columns[0].Visible = false;
        }

        public void panelhide()
        {
            panel1.Visible = false;
            panel3.Visible = false;
            panel4.Visible = false;
        }

        private DataTable GetDataForMonthkitchenhedproducts()
        {

            DataTable dt = new DataTable();
            try
            {
                // зачем выгружать из sql, если есть таблица
                string query = @"SELECT Id, tab_no As Табельный,fio As ФИО, summa As Сумма, val As Собственная,val1 As Готовая, type As Код FROM dbo.dobor_" + DateTime.Now.Month + "_" + DateTime.Now.Year + " Where type = 521";
                SqlCommand comm = new SqlCommand(query, con);
                con.Open();
                SqlDataAdapter da = new SqlDataAdapter(comm);
                DataSet ds = new DataSet();
                da.Fill(ds);
                dt = ds.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                con.Close();
                //con.Dispose();
            }
            return dt;
        }
        private DataTable GetDataTabel2()
        {
            DataTable dt = new DataTable();
            try
            {
                string date4 = dateTimePicker2.Value.ToString("dd.MM.yyyy 00:00");
                string date5 = dateTimePicker2.Value.ToString("dd.MM.yyyy 23:59");
                // зачем выгружать из sql, если есть таблица
                string query = @"SELECT tab_no As Табельный,fio As ФИО,summa As Сумма, val As Собственная,val1 As Готовая,type As Тип,date As Дата FROM dbo.dobor_" + dateTimePicker2.Value.Month + "_" + dateTimePicker2.Value.Year + " Where  tab_no='" + maskedTextBox1.Text + "' and date between '" + date4 + "' and '" + date5 + "'";
                SqlCommand comm = new SqlCommand(query, con);
                con.Open();
                SqlDataAdapter da = new SqlDataAdapter(comm);
                DataSet ds = new DataSet();
                da.Fill(ds);
                dt = ds.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                con.Close();
                //con.Dispose();
            }
            return dt;
        }
        private DataTable GetDataTabel()
        {
            DataTable dt = new DataTable();
            try
            {
                // зачем выгружать из sql, если есть таблица
                string query = @"SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма, val As Собственная,val1 As Готовая,type As Тип,date As Дата FROM dbo.dobor_" + dateTimePicker2.Value.Month + "_" + dateTimePicker2.Value.Year + " Where  tab_no='" + maskedTextBox1.Text + "'";
                SqlCommand comm = new SqlCommand(query, con);
                con.Open();
                SqlDataAdapter da = new SqlDataAdapter(comm);
                DataSet ds = new DataSet();
                da.Fill(ds);
                dt = ds.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                con.Close();
               //con.Dispose();
            }
            return dt;
        }

        private DataTable GetDataShop()
        {

            DataTable dt = new DataTable();
            try
            {
                // зачем выгружать из sql, если есть таблица
                string query = @"SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма,type As Тип FROM dbo.dobor_" + DateTime.Now.Month + "_" + DateTime.Now.Year + " Where type=518 ";
                SqlCommand comm = new SqlCommand(query, con);
                con.Open();
                SqlDataAdapter da = new SqlDataAdapter(comm);
                DataSet ds = new DataSet();
                da.Fill(ds);
                dt = ds.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                con.Close();
                //con.Dispose();
            }
            return dt;
        }



        private DataTable GetDataForMonthShop()
        {
            DataTable dt = new DataTable();
            try
            {
                // зачем выгружать из sql, если есть таблица
                string query = @"SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма, type As Тип FROM dbo.dobor_" + dateTimePicker4.Value.Month + "_" + dateTimePicker4.Value.Year + " Where type = 518";
                SqlCommand comm = new SqlCommand(query, con);
                con.Open();
                SqlDataAdapter da = new SqlDataAdapter(comm);
                DataSet ds = new DataSet();
                da.Fill(ds);
                dt = ds.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                con.Close();
                //con.Dispose();
            }
            return dt;
        }


        private DataTable GetDataForMonthKitchenDataPicker()
        {
            DataTable dt = new DataTable();
            try
            {
                // зачем выгружать из sql, если есть таблица
                string query = @"SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма,val As Собственная,val1 As Готовая,type As Тип FROM dbo.dobor_" + dateTimePicker4.Value.Month + "_" + dateTimePicker4.Value.Year + " Where type = 521";
                SqlCommand comm = new SqlCommand(query, con);
                con.Open();
                SqlDataAdapter da = new SqlDataAdapter(comm);
                DataSet ds = new DataSet();
                da.Fill(ds);
                dt = ds.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                con.Close();
                //con.Dispose();
            }
            return dt;
        }



        private DataTable GetDataForMonth()
        {

            DataTable dt = new DataTable();
            try
            {
                // зачем выгружать из sql, если есть таблица
                string query = @"SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма,val As Собственная,val1 As Готовая,type As Тип FROM dbo.dobor_" + dateTimePicker4.Value.Month + "_" + dateTimePicker4.Value.Year + "";
                SqlCommand comm = new SqlCommand(query, con);
                con.Open();
                SqlDataAdapter da = new SqlDataAdapter(comm);
                DataSet ds = new DataSet();
                da.Fill(ds);
                dt = ds.Tables[0];
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                con.Close();
                //con.Dispose();
            }
            return dt;
        }

        private DataTable GetDataTodayShop()
        {
            DataTable dt = new DataTable();
            try
            {

                // зачем выгружать из sql, если есть таблица
                string date = DateTime.UtcNow.ToString("dd.MM.yyyy 23:59");
                string query = @"SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма,type As Тип FROM dbo.dobor_" + DateTime.Now.Month + "_" + DateTime.Now.Year + " Where date between '" + DateTime.Today + "' And '" + date + "' And type = 518";
                SqlCommand comm = new SqlCommand(query, con);
                con.Open();
                SqlDataAdapter da = new SqlDataAdapter(comm);
                DataSet ds = new DataSet();
                da.Fill(ds);
                dt = ds.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                con.Close();
                //con.Dispose();
            }
            return dt;
        }

        private DataTable GetDataTodaykitchen()
        {

            DataTable dt = new DataTable();
            try
            {
                string date = DateTime.UtcNow.ToString("dd.MM.yyyy 23:59");
                // зачем выгружать из sql, если есть таблица
                string query = @"SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма,val As Собственная,val1 As Готовая,type As Тип FROM dbo.dobor_" + DateTime.Now.Month + "_" + DateTime.Now.Year + " Where date between '" + DateTime.Today + "' And '" + date + "' And type = 521";
                SqlCommand comm = new SqlCommand(query, con);
                con.Open();
                SqlDataAdapter da = new SqlDataAdapter(comm);
                DataSet ds = new DataSet();
                da.Fill(ds);
                dt = ds.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                con.Close();
                //con.Dispose();
            }
            return dt;
        }
        private DataTable GetDataToday()
        {

            DataTable dt = new DataTable();
            try
            {
                // зачем выгружать из sql, если есть таблица
                string date = DateTime.UtcNow.ToString("dd.MM.yyyy 23:59");
                string query = @"SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма,val As Собственная,val1 As Готовая,type As Тип FROM dbo.dobor_" + DateTime.Now.Month + "_" + DateTime.Now.Year + " Where date between '" + DateTime.Today + "' And '" + date + "'";
                SqlCommand comm = new SqlCommand(query, con);
                con.Open();
                SqlDataAdapter da = new SqlDataAdapter(comm);
                DataSet ds = new DataSet();
                da.Fill(ds);
                dt = ds.Tables[0];

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                con.Close();
                //con.Dispose();
            }
            return dt;
        }

        private DataTable GetDataDayKitchen()
        {
            DataTable dt = new DataTable();
            try
            {
                string date1 = dateTimePicker1.Value.ToString("dd.MM.yyyy 00:00");
                string date2 = dateTimePicker1.Value.ToString("dd.MM.yyyy 23:59");

                // зачем выгружать из sql, если есть таблица
                string query = @"SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма, val As Собственная,val1 As Готовая,type As Тип FROM dbo.dobor_" + dateTimePicker1.Value.Month + "_" + dateTimePicker1.Value.Year + " Where date between '" + date1 + "' And '" + date2 + "' And Type = 521";
                SqlCommand comm = new SqlCommand(query, con);
                con.Open();
                SqlDataAdapter da = new SqlDataAdapter(comm);
                DataSet ds = new DataSet();
                da.Fill(ds);
                dt = ds.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                con.Close();
                //con.Dispose();
            }
            return dt;
        }



        private DataTable GetDataDayShop()
        {

            DataTable dt = new DataTable();
            try
            {
                string date1 = dateTimePicker1.Value.ToString("dd.MM.yyyy 00:00");
                string date2 = dateTimePicker1.Value.ToString("dd.MM.yyyy 23:59");
                // зачем выгружать из sql, если есть таблица
                string query = @"SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма,type As Тип FROM dbo.dobor_" + dateTimePicker1.Value.Month + "_" + dateTimePicker1.Value.Year + " Where date between '" + date1 + "' And '" + date2 + "' And Type = 518";
                SqlCommand comm = new SqlCommand(query, con);
                con.Open();
                SqlDataAdapter da = new SqlDataAdapter(comm);
                DataSet ds = new DataSet();
                da.Fill(ds);
                dt = ds.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                con.Close();
                //con.Dispose();
            }
            return dt;
        }




        private DataTable GetDataDay()
        {

            DataTable dt = new DataTable();
            try
            {
                string date1 = dateTimePicker1.Value.ToString("dd.MM.yyyy 00:00");
                string date2 = dateTimePicker1.Value.ToString("dd.MM.yyyy 23:59");
                // зачем выгружать из sql, если есть таблица
                string query = @"SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма, val As Собственная,val1 As Готовая,type As Тип FROM dbo.dobor_" + dateTimePicker1.Value.Month + "_" + dateTimePicker1.Value.Year + " Where date between '" + date1 + "' And '" + date2 + "'";
                SqlCommand comm = new SqlCommand(query, con);
                con.Open();
                SqlDataAdapter da = new SqlDataAdapter(comm);
                DataSet ds = new DataSet();
                da.Fill(ds);
                dt = ds.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                con.Close();
                //con.Dispose();
            }
            return dt;
        }

        private DataTable GetDataTotal()
        {

            DataTable dt = new DataTable();
            try
            {
                string query = @"SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма,val As Собственная,val1 As Готовая,type As Тип FROM dbo.dobor_" + DateTime.Now.Month + "_" + DateTime.Now.Year + "";
                SqlCommand comm = new SqlCommand(query, con);
                con.Open();
                SqlDataAdapter da = new SqlDataAdapter(comm);
                DataSet ds = new DataSet();
                da.Fill(ds);
                dt = ds.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                con.Close();
                //con.Dispose();
            }
            return dt;
        }

        public void ReportTabel2()
        {

            string date1 = dateTimePicker2.Value.ToString("dd.MM.yyyy 00:00");
            string date2 = dateTimePicker2.Value.ToString("dd.MM.yyyy 23:59");
            DataTable dt1 = new DataTable();
            SqlDataAdapter ada = new SqlDataAdapter("SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма, val As Собственная,val1 As Готовая,type As Тип,date As Дата FROM dbo.dobor_" + dateTimePicker2.Value.Month + "_" + dateTimePicker2.Value.Year + " Where  tab_no='" + maskedTextBox1.Text + "' and date between '" + date1 + "' and '" + date2 + "'", con);
            ada.Fill(dt1);
            if (dt1.Rows.Count > 0)
            {
                MessageBox.Show("Данные экспортированы в Excel");
                xlApp = new Excel.Application();
                xlApp = new Excel.Application();
                try
                {
                    xlApp.Workbooks.Add(Type.Missing);
                    xlApp.Interactive = false;
                    xlApp.EnableEvents = false;
                    xlSheet = (Excel.Worksheet)xlApp.Sheets[1];
                    xlSheet.Name = "Данные";
                    DataTable dt = GetDataTabel2();
                    int collInd = 0;
                    int rowInd = 0;
                    string data = "";

                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        data = dt.Columns[i].ColumnName.ToString();
                        xlSheet.Cells[2, i + 1] = data;
                        xlSheetRange = xlSheet.get_Range("A2:Z2", Type.Missing);
                    }

                    for (rowInd = 0; rowInd < dt.Rows.Count; rowInd++)
                    {

                        for (collInd = 0; collInd < dt.Columns.Count; collInd++)
                        {
                            data = dt.Rows[rowInd].ItemArray[collInd].ToString();
                            xlSheet.Cells[rowInd + 3, collInd + 1] = data;

                        }

                    }
                    xlSheet.Cells[rowInd + 3, collInd - 5] = "Итоговая сумма:";
                    xlSheet.Cells[rowInd + 3, collInd - 4].Formula = "=Sum(" + xlSheet.Cells[1, 4].Address + ":" + xlSheet.Cells[rowInd + 2, 4].Address + ")";

                    // жесть... Если работаешь с range, то зачем к каждой ячейке ставить стили

                    xlSheet.get_Range("A2:A2", Type.Missing);
                    Microsoft.Office.Interop.Excel.Range tRange = xlSheet.UsedRange;
                    tRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    tRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                    xlSheet.get_Range("A1", "G1").Merge();
                    xlSheet.Cells[1, 1] = "Ведомость по  табельному № " + maskedTextBox1.Text + " за " + dateTimePicker2.Value.ToString("dd.MM.yyyy");
                    xlSheetRange = xlSheet.UsedRange;
                    xlSheetRange.Columns.AutoFit();
                    xlSheetRange.Cells.HorizontalAlignment = -4108;
                    xlSheetRange.Cells.VerticalAlignment = -4108;
                    xlSheetRange.Rows.AutoFit();
                    (xlSheetRange.Cells[1, 1] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 1] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 2] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 3] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 4] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 5] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 6] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 7] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 8] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[rowInd + 3, collInd - 5] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[rowInd + 3, collInd - 4] as Excel.Range).Font.Bold = true;

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                finally
                {

                    xlApp.Visible = true;
                    xlApp.Interactive = true;
                    xlApp.ScreenUpdating = true;
                    xlApp.UserControl = true;
                    releaseObject(xlSheetRange);
                    releaseObject(xlSheet);
                    releaseObject(xlApp);

                }
            } else
            {
                MessageBox.Show("Данных за выбранный период не было!");
                maskedTextBox1.Clear();
            }
        }


        public void ReportTabel()
        {
            MessageBox.Show("Данные экспортированы в Excel");
            xlApp = new Excel.Application();
            xlApp = new Excel.Application();
            try
            {
                xlApp.Workbooks.Add(Type.Missing);
                xlApp.Interactive = false;
                xlApp.EnableEvents = false;
                xlSheet = (Excel.Worksheet)xlApp.Sheets[1];
                xlSheet.Name = "Данные";
                DataTable dt = GetDataTabel();
                int collInd = 0;
                int rowInd = 0;
                string data = "";

                for (int i = 0; i < dt.Columns.Count; i++)
                {
       
                    data = dt.Columns[i].ColumnName.ToString();
                    xlSheet.Cells[2, i + 1] = data;
                    xlSheetRange = xlSheet.get_Range("A2:Z2", Type.Missing);
                }

                for (rowInd = 0; rowInd < dt.Rows.Count; rowInd++)
                {
                   
                    for (collInd = 0; collInd < dt.Columns.Count; collInd++)
                    {

                        data = dt.Rows[rowInd].ItemArray[collInd].ToString();
                        xlSheet.Cells[rowInd + 3, collInd + 1] = data;

                    }

                }
              
                xlSheet.Cells[rowInd + 3, collInd - 5] = "Итоговая сумма:";
                xlSheet.Cells[rowInd + 3, collInd - 4].Formula = "=Sum(" + xlSheet.Cells[1, 4].Address + ":" + xlSheet.Cells[rowInd + 2, 4].Address + ")";

                // жесть... Если работаешь с range, то зачем к каждой ячейке ставить стили

                xlSheet.get_Range("A2:A2", Type.Missing);
                Microsoft.Office.Interop.Excel.Range tRange = xlSheet.UsedRange;
                tRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                tRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                xlSheet.get_Range("A1", "G1").Merge();
                xlSheet.Cells[1, 1] = "Ведомость по  табельному № " + maskedTextBox1.Text  + " за " + dateTimePicker2.Value.Month + " месяц  " + dateTimePicker2.Value.Year + " года";
                xlSheetRange = xlSheet.UsedRange;
                xlSheetRange.Columns.AutoFit();
                xlSheetRange.Cells.HorizontalAlignment = -4108;
                xlSheetRange.Cells.VerticalAlignment = -4108;
                xlSheetRange.Rows.AutoFit();
                (xlSheetRange.Cells[1, 1] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 1] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 2] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 3] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 4] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 5] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 6] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 7] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 8] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[rowInd + 3, collInd - 5] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[rowInd + 3, collInd - 4] as Excel.Range).Font.Bold = true;
               
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {

                xlApp.Visible = true;
                xlApp.Interactive = true;
                xlApp.ScreenUpdating = true;
                xlApp.UserControl = true;
                releaseObject(xlSheetRange);
                releaseObject(xlSheet);
                releaseObject(xlApp);

            }

        }

        public void ReportForMonthKitchen()
        {
            //con.Open();
            MessageBox.Show("Данные экспортированы в Excel");
            xlApp = new Excel.Application();
            xlApp = new Excel.Application();
            try
            {
                xlApp.Workbooks.Add(Type.Missing);
                xlApp.Interactive = false;
                xlApp.EnableEvents = false;
                xlSheet = (Excel.Worksheet)xlApp.Sheets[1];
                xlSheet.Name = "Данные";
                DataTable dt = GetDataForMonthKitchenDataPicker();
                int collInd = 0;
                int rowInd = 0;
                string data = "";

                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    data = dt.Columns[i].ColumnName.ToString();
                    xlSheet.Cells[2, i + 1] = data;
                    xlSheetRange = xlSheet.get_Range("A2:Z2", Type.Missing);
                }

                for (rowInd = 0; rowInd < dt.Rows.Count; rowInd++)
                {

                    for (collInd = 0; collInd < dt.Columns.Count; collInd++)
                    {
                        data = dt.Rows[rowInd].ItemArray[collInd].ToString();
                        xlSheet.Cells[rowInd + 3, collInd + 1] = data;

                    }

                }

                xlSheet.Cells[rowInd + 4, collInd - 4] = "Итого:";
                xlSheet.Cells[rowInd + 4, collInd - 3].Formula = "=Sum(" + xlSheet.Cells[1, 4].Address + ":" + xlSheet.Cells[rowInd + 2, 4].Address + ")";
                xlSheet.Cells[rowInd + 4, collInd - 2].Formula = "=Sum(" + xlSheet.Cells[1, 5].Address + ":" + xlSheet.Cells[rowInd + 2, 5].Address + ")";
                xlSheet.Cells[rowInd + 4, collInd - 1].Formula = "=Sum(" + xlSheet.Cells[1, 6].Address + ":" + xlSheet.Cells[rowInd + 2, 6].Address + ")";
                Microsoft.Office.Interop.Excel.Range tRange = xlSheet.UsedRange;
                tRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                tRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                xlSheet.get_Range("A1", "G1").Merge();
                xlSheet.Cells[1, 1] = "Ведомость по столовой за " + dateTimePicker4.Value.Month + " месяц " + dateTimePicker4.Value.Year + " года";
                xlSheetRange = xlSheet.UsedRange;
                xlSheetRange.Columns.AutoFit();
                xlSheetRange.Cells.HorizontalAlignment = -4108;
                xlSheetRange.Cells.VerticalAlignment = -4108;
                xlSheetRange.Rows.AutoFit();

                (xlSheetRange.Cells[1, 1] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 1] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 2] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 3] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 4] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 5] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 6] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 7] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 8] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[rowInd + 4, collInd - 4] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[rowInd + 4, collInd - 3] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[rowInd + 4, collInd - 3] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[rowInd + 4, collInd - 2] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[rowInd + 4, collInd - 1] as Excel.Range).Font.Bold = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {

                xlApp.Visible = true;
                xlApp.Interactive = true;
                xlApp.ScreenUpdating = true;
                xlApp.UserControl = true;
                releaseObject(xlSheetRange);
                releaseObject(xlSheet);
                releaseObject(xlApp);

            }
            //con.Close();
        }



        public void ReportForMonthShop()
        {

            MessageBox.Show("Данные экспортированы в Excel"); // уже?)
            xlApp = new Excel.Application();
            xlApp = new Excel.Application();
            try
            {
                xlApp.Workbooks.Add(Type.Missing);
                xlApp.Interactive = false;
                xlApp.EnableEvents = false;
                xlSheet = (Excel.Worksheet)xlApp.Sheets[1];
                xlSheet.Name = "Данные";
                DataTable dt = GetDataForMonthShop();
                int collInd = 0;
                int rowInd = 0;
                string data = "";

                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    data = dt.Columns[i].ColumnName.ToString();
                    xlSheet.Cells[2, i + 1] = data;
                    xlSheetRange = xlSheet.get_Range("A2:Z2", Type.Missing);
                }

                for (rowInd = 0; rowInd < dt.Rows.Count; rowInd++)
                {

                    for (collInd = 0; collInd < dt.Columns.Count; collInd++)
                    {
                        data = dt.Rows[rowInd].ItemArray[collInd].ToString();
                        xlSheet.Cells[rowInd + 3, collInd + 1] = data;

                    }

                }

                xlSheet.Cells[rowInd + 4, collInd - 2] = "Итоговая сумма:";
                xlSheet.Cells[rowInd + 4, collInd - 1].Formula = "=Sum(" + xlSheet.Cells[3, 4].Address + ":" + xlSheet.Cells[rowInd + 2, 4].Address + ")";
                Microsoft.Office.Interop.Excel.Range tRange = xlSheet.UsedRange;
                tRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                tRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                xlSheet.get_Range("A1", "E1").Merge();
                xlSheet.Cells[1, 1] = "Ведомость по магазину за " + dateTimePicker4.Value.Month + " месяц " + dateTimePicker4.Value.Year + " года";
                xlSheetRange = xlSheet.UsedRange;
                xlSheetRange.Columns.AutoFit();
                xlSheetRange.Cells.HorizontalAlignment = -4108;
                xlSheetRange.Cells.VerticalAlignment = -4108;
                xlSheetRange.Rows.AutoFit();

                (xlSheetRange.Cells[1, 1] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 1] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 2] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 3] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 4] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 5] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 6] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 7] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 8] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[rowInd + 4, collInd - 2] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[rowInd + 4, collInd - 1] as Excel.Range).Font.Bold = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {

                xlApp.Visible = true;
                xlApp.Interactive = true;
                xlApp.ScreenUpdating = true;
                xlApp.UserControl = true;
                releaseObject(xlSheetRange);
                releaseObject(xlSheet);
                releaseObject(xlApp);

            }
            //con.Close();
        }


        public void ReportForMonth()
        {

            //con.Open();
            MessageBox.Show("Данные экспортированы в Excel");
            xlApp = new Excel.Application();
            xlApp = new Excel.Application();
            try
            {
                xlApp.Workbooks.Add(Type.Missing);
                xlApp.Interactive = false;
                xlApp.EnableEvents = false;
                xlSheet = (Excel.Worksheet)xlApp.Sheets[1];
                xlSheet.Name = "Данные";
                DataTable dt = GetDataForMonth();
                int collInd = 0;
                int rowInd = 0;
                string data = "";

                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    data = dt.Columns[i].ColumnName.ToString();
                    xlSheet.Cells[2, i + 1] = data;
                    xlSheetRange = xlSheet.get_Range("A2:Z2", Type.Missing);
                }

                for (rowInd = 0; rowInd < dt.Rows.Count; rowInd++)
                {

                    for (collInd = 0; collInd < dt.Columns.Count; collInd++)
                    {
                        data = dt.Rows[rowInd].ItemArray[collInd].ToString();
                        xlSheet.Cells[rowInd + 3, collInd + 1] = data;

                    }

                }

                xlSheet.Cells[rowInd + 4, collInd - 4] = "Итоговая сумма:";
                xlSheet.Cells[rowInd + 4, collInd - 3].Formula = "=Sum(" + xlSheet.Cells[3, 4].Address + ":" + xlSheet.Cells[rowInd + 2, 4].Address + ")";
                Microsoft.Office.Interop.Excel.Range tRange = xlSheet.UsedRange;
                tRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                tRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                xlSheet.get_Range("A1", "G1").Merge();
                xlSheet.Cells[1, 1] = "Ведомость за " + dateTimePicker4.Value.Month + " месяц " + dateTimePicker4.Value.Year + " года";
                xlSheetRange = xlSheet.UsedRange;
                xlSheetRange.Columns.AutoFit();
                xlSheetRange.Cells.HorizontalAlignment = -4108;
                xlSheetRange.Cells.VerticalAlignment = -4108;
                xlSheetRange.Rows.AutoFit();

                (xlSheetRange.Cells[1, 1] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 1] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 2] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 3] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 4] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 5] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 6] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 7] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[2, 8] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[rowInd + 4, collInd - 4] as Excel.Range).Font.Bold = true;
                (xlSheetRange.Cells[rowInd + 4, collInd - 3] as Excel.Range).Font.Bold = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {

                xlApp.Visible = true;
                xlApp.Interactive = true;
                xlApp.ScreenUpdating = true;
                xlApp.UserControl = true;
                releaseObject(xlSheetRange);
                releaseObject(xlSheet);
                releaseObject(xlApp);

            }
            //con.Close();
        }

        public void ReportDate()
        {

            SqlDataAdapter ada3 = new SqlDataAdapter("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'dobor_" + dateTimePicker1.Value.Month + "_" + dateTimePicker1.Value.Year + "'", con);
            DataTable dt3 = new DataTable();
            ada3.Fill(dt3);
            if (dt3.Rows.Count > 0)
            {

                string date1 = dateTimePicker1.Value.ToString("dd.MM.yyyy 00:00");
                string date2 = dateTimePicker1.Value.ToString("dd.MM.yyyy 23:59");
                string date3 = dateTimePicker1.Value.ToString("dd.MM.yyyy");
                DataTable dt1 = new DataTable();
                SqlDataAdapter ada = new SqlDataAdapter("SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма, val As Собственная,val1 As Готовая,type As Тип,date As Дата FROM dbo.dobor_" + dateTimePicker1.Value.Month + "_" + dateTimePicker1.Value.Year + " Where date between '" + date1 + "' And '" + date2 + "'", con);
                ada.Fill(dt1);
                if (dt1.Rows.Count > 0)
                {
                    MessageBox.Show("Данные экспортированы в Excel");
                    xlApp = new Excel.Application();
                    xlApp = new Excel.Application();
                    try
                    {
                        xlApp.Workbooks.Add(Type.Missing);
                        xlApp.Interactive = false;
                        xlApp.EnableEvents = false;
                        xlSheet = (Excel.Worksheet)xlApp.Sheets[1];
                        xlSheet.Name = "Данные";
                        DataTable dt = GetDataDay();
                        int collInd = 0;
                        int rowInd = 0;
                        string data = "";

                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            data = dt.Columns[i].ColumnName.ToString();
                            xlSheet.Cells[2, i + 1] = data;
                            xlSheetRange = xlSheet.get_Range("A2:Z2", Type.Missing);
                        }

                        for (rowInd = 0; rowInd < dt.Rows.Count; rowInd++)
                        {

                            for (collInd = 0; collInd < dt.Columns.Count; collInd++)
                            {
                                data = dt.Rows[rowInd].ItemArray[collInd].ToString();
                                xlSheet.Cells[rowInd + 3, collInd + 1] = data;

                            }

                        }
                        xlSheet.Cells[rowInd + 4, collInd - 4] = "Итого:";
                        xlSheet.Cells[rowInd + 4, collInd - 3].Formula = "=Sum(" + xlSheet.Cells[1, 4].Address + ":" + xlSheet.Cells[rowInd + 2, 4].Address + ")";
                        xlSheet.Cells[rowInd + 4, collInd - 2].Formula = "=Sum(" + xlSheet.Cells[1, 5].Address + ":" + xlSheet.Cells[rowInd + 2, 5].Address + ")";
                        xlSheet.Cells[rowInd + 4, collInd - 1].Formula = "=Sum(" + xlSheet.Cells[1, 6].Address + ":" + xlSheet.Cells[rowInd + 2, 6].Address + ")";
                        Microsoft.Office.Interop.Excel.Range tRange = xlSheet.UsedRange;
                        tRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        tRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                        xlSheet.get_Range("A1", "G1").Merge();
                        xlSheet.Cells[1, 1] = "Ведомость за " + date3 + "";
                        xlSheetRange = xlSheet.UsedRange;
                        xlSheetRange.Columns.AutoFit();
                        xlSheetRange.Cells.HorizontalAlignment = -4108;
                        xlSheetRange.Cells.VerticalAlignment = -4108;
                        xlSheetRange.Rows.AutoFit();
                        (xlSheetRange.Cells[1, 1] as Excel.Range).Font.Bold = true;
                        (xlSheetRange.Cells[2, 1] as Excel.Range).Font.Bold = true;
                        (xlSheetRange.Cells[2, 2] as Excel.Range).Font.Bold = true;
                        (xlSheetRange.Cells[2, 3] as Excel.Range).Font.Bold = true;
                        (xlSheetRange.Cells[2, 4] as Excel.Range).Font.Bold = true;
                        (xlSheetRange.Cells[2, 5] as Excel.Range).Font.Bold = true;
                        (xlSheetRange.Cells[2, 6] as Excel.Range).Font.Bold = true;
                        (xlSheetRange.Cells[2, 7] as Excel.Range).Font.Bold = true;
                        (xlSheetRange.Cells[2, 8] as Excel.Range).Font.Bold = true;
                        (xlSheetRange.Cells[rowInd + 4, collInd - 4] as Excel.Range).Font.Bold = true;
                        (xlSheetRange.Cells[rowInd + 4, collInd - 3] as Excel.Range).Font.Bold = true;
                        (xlSheetRange.Cells[rowInd + 4, collInd - 2] as Excel.Range).Font.Bold = true;
                        (xlSheetRange.Cells[rowInd + 4, collInd - 1] as Excel.Range).Font.Bold = true;

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                    finally
                    {

                        xlApp.Visible = true;
                        xlApp.Interactive = true;
                        xlApp.ScreenUpdating = true;
                        xlApp.UserControl = true;
                        releaseObject(xlSheetRange);
                        releaseObject(xlSheet);
                        releaseObject(xlApp);

                    }
                }
                else
                {
                    MessageBox.Show("Данных за выбранный период не существует!");
                }
            }
            else
            {
                MessageBox.Show("Данных за выбранный период не существует!");
            }
        }


        public void ReportShop()
        {

            SqlDataAdapter ada2 = new SqlDataAdapter("SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма,type As Тип,date As Дата FROM dbo.dobor_" + DateTime.Now.Month + "_" + DateTime.Now.Year + " Where type=518 ", con);
            DataTable dt2 = new DataTable();
            ada2.Fill(dt2);
            if (dt2.Rows.Count > 0)
            {
                MessageBox.Show("Данные экспортированы в Excel");
                xlApp = new Excel.Application();
                xlApp = new Excel.Application();
                try
                {
                    xlApp.Workbooks.Add(Type.Missing);
                    xlApp.Interactive = false;
                    xlApp.EnableEvents = false;
                    xlSheet = (Excel.Worksheet)xlApp.Sheets[1];
                    xlSheet.Name = "Данные";
                    DataTable dt = GetDataShop();
                    int collInd = 0;
                    int rowInd = 0;
                    string data = "";

                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        data = dt.Columns[i].ColumnName.ToString();
                        xlSheet.Cells[2, i + 1] = data;
                        xlSheetRange = xlSheet.get_Range("A2:Z2", Type.Missing);
                    }

                    for (rowInd = 0; rowInd < dt.Rows.Count; rowInd++)
                    {

                        for (collInd = 0; collInd < dt.Columns.Count; collInd++)
                        {
                            data = dt.Rows[rowInd].ItemArray[collInd].ToString();
                            xlSheet.Cells[rowInd + 3, collInd + 1] = data;

                        }

                    }
                    xlSheet.Cells[rowInd + 4, collInd - 2] = "Итоговая сумма:";
                    xlSheet.Cells[rowInd + 4, collInd - 1].Formula = "=Sum(" + xlSheet.Cells[3, 4].Address + ":" + xlSheet.Cells[rowInd + 2, 4].Address + ")";
                    Microsoft.Office.Interop.Excel.Range tRange = xlSheet.UsedRange;
                    tRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    tRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                    xlSheet.get_Range("A1", "E1").Merge();
                    xlSheet.Cells[1, 1] = "Ведомость по магазину за " + DateTime.Now.Month + " месяц " + DateTime.Now.Year + " года";
                    xlSheetRange = xlSheet.UsedRange;
                    xlSheetRange.Columns.AutoFit();
                    xlSheetRange.Cells.HorizontalAlignment = -4108;
                    xlSheetRange.Cells.VerticalAlignment = -4108;
                    xlSheetRange.Rows.AutoFit();
                    (xlSheetRange.Cells[1, 1] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 1] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 2] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 3] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 4] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 5] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 6] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[rowInd + 4, collInd - 2] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[rowInd + 4, collInd - 1] as Excel.Range).Font.Bold = true;

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                finally
                {

                    xlApp.Visible = true;
                    xlApp.Interactive = true;
                    xlApp.ScreenUpdating = true;
                    xlApp.UserControl = true;
                    releaseObject(xlSheetRange);
                    releaseObject(xlSheet);
                    releaseObject(xlApp);

                }
            }
            else
            {
                MessageBox.Show("Данных по магазину за месяц не было!");
            }
        }

        public void ReportForkitchenproducts()
        {

            SqlDataAdapter ada2 = new SqlDataAdapter("SELECT Id, tab_no As Табельный,fio As ФИО, summa As Сумма, val As Собственная,val1 As Готовая, type As Код,date As Дата FROM dbo.dobor_" + DateTime.Now.Month + "_" + DateTime.Now.Year + " Where type = 521", con);
            DataTable dt2 = new DataTable();
            ada2.Fill(dt2);
            if (dt2.Rows.Count > 0)
            {
                MessageBox.Show("Данные экспортированы в Excel");
                xlApp = new Excel.Application();
                xlApp = new Excel.Application();
                try
                {
                    xlApp.Workbooks.Add(Type.Missing);
                    xlApp.Interactive = false;
                    xlApp.EnableEvents = false;
                    xlSheet = (Excel.Worksheet)xlApp.Sheets[1];
                    xlSheet.Name = "Данные";
                    DataTable dt = GetDataForMonthkitchenhedproducts();
                    int collInd = 0;
                    int rowInd = 0;
                    string data = "";

                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        data = dt.Columns[i].ColumnName.ToString();
                        xlSheet.Cells[2, i + 1] = data;
                        xlSheetRange = xlSheet.get_Range("A2:Z2", Type.Missing);
                    }
                    for (rowInd = 0; rowInd < dt.Rows.Count; rowInd++)
                    {

                        for (collInd = 0; collInd < dt.Columns.Count; collInd++)
                        {
                            data = dt.Rows[rowInd].ItemArray[collInd].ToString();
                            xlSheet.Cells[rowInd + 3, collInd + 1] = data;

                        }
                    }
                    xlSheet.Cells[rowInd + 4, collInd - 4] = "Итого:";
                    xlSheet.Cells[rowInd + 4, collInd - 3].Formula = "=Sum(" + xlSheet.Cells[1, 4].Address + ":" + xlSheet.Cells[rowInd + 2, 4].Address + ")";
                    xlSheet.Cells[rowInd + 4, collInd - 2].Formula = "=Sum(" + xlSheet.Cells[1, 5].Address + ":" + xlSheet.Cells[rowInd + 2, 5].Address + ")";
                    xlSheet.Cells[rowInd + 4, collInd - 1].Formula = "=Sum(" + xlSheet.Cells[1, 6].Address + ":" + xlSheet.Cells[rowInd + 2, 6].Address + ")";
                    Microsoft.Office.Interop.Excel.Range tRange = xlSheet.UsedRange;
                    tRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    tRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                    xlSheet.get_Range("A1", "G1").Merge();
                    xlSheet.Cells[1, 1] = "Ведомость по столовой за " + DateTime.Now.Month + " месяц " + DateTime.Now.Year + " года";
                    xlSheetRange = xlSheet.UsedRange;
                    xlSheetRange.Columns.AutoFit();
                    xlSheetRange.Cells.HorizontalAlignment = -4108;
                    xlSheetRange.Cells.VerticalAlignment = -4108;
                    xlSheetRange.Rows.AutoFit();
                    (xlSheetRange.Cells[2, 1] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[1, 1] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 2] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 3] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 4] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 5] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 6] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 7] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 8] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[rowInd + 4, collInd - 4] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[rowInd + 4, collInd - 3] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[rowInd + 4, collInd - 3] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[rowInd + 4, collInd - 2] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[rowInd + 4, collInd - 1] as Excel.Range).Font.Bold = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                finally
                {

                    xlApp.Visible = true;
                    xlApp.Interactive = true;
                    xlApp.ScreenUpdating = true;
                    xlApp.UserControl = true;
                    releaseObject(xlSheetRange);
                    releaseObject(xlSheet);
                    releaseObject(xlApp);

                }
            }
            else
            {
                MessageBox.Show("Данных по столовой не было!");
            }
        }


        public void ReportKitchenPerDate()
        {

            string date1 = dateTimePicker1.Value.ToString("dd.MM.yyyy 00:00");
            string date2 = dateTimePicker1.Value.ToString("dd.MM.yyyy 23:59");
            SqlDataAdapter ada2 = new SqlDataAdapter("SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма,val As Собственная,val1 As Готовая,type As Тип,date As Дата FROM dbo.dobor_" + dateTimePicker1.Value.Month + "_" + dateTimePicker1.Value.Year + " Where date between '" + date1 + "' And '" + date2 + "' And type = '521'", con);
            DataTable dt2 = new DataTable();
            ada2.Fill(dt2);
            if (dt2.Rows.Count > 0)
            {
                MessageBox.Show("Данные экспортированы в Excel");
                xlApp = new Excel.Application();
                xlApp = new Excel.Application();
                try
                {
                    xlApp.Workbooks.Add(Type.Missing);
                    xlApp.Interactive = false;
                    xlApp.EnableEvents = false;
                    xlSheet = (Excel.Worksheet)xlApp.Sheets[1];
                    xlSheet.Name = "Данные";
                    DataTable dt = GetDataDayKitchen();
                    int collInd = 0;
                    int rowInd = 0;
                    string data = "";

                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        data = dt.Columns[i].ColumnName.ToString();
                        xlSheet.Cells[2, i + 1] = data;
                        xlSheetRange = xlSheet.get_Range("A2:Z2", Type.Missing);
                    }
                    for (rowInd = 0; rowInd < dt.Rows.Count; rowInd++)
                    {

                        for (collInd = 0; collInd < dt.Columns.Count; collInd++)
                        {
                            data = dt.Rows[rowInd].ItemArray[collInd].ToString();
                            xlSheet.Cells[rowInd + 3, collInd + 1] = data;

                        }
                    }
                    xlSheet.Cells[rowInd + 4, collInd - 4] = "Итого:";
                    xlSheet.Cells[rowInd + 4, collInd - 3].Formula = "=Sum(" + xlSheet.Cells[1, 4].Address + ":" + xlSheet.Cells[rowInd + 2, 4].Address + ")";
                    xlSheet.Cells[rowInd + 4, collInd - 2].Formula = "=Sum(" + xlSheet.Cells[1, 5].Address + ":" + xlSheet.Cells[rowInd + 2, 5].Address + ")";
                    xlSheet.Cells[rowInd + 4, collInd - 1].Formula = "=Sum(" + xlSheet.Cells[1, 6].Address + ":" + xlSheet.Cells[rowInd + 2, 6].Address + ")";
                    Microsoft.Office.Interop.Excel.Range tRange = xlSheet.UsedRange;
                    tRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    tRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                    xlSheet.get_Range("A1", "G1").Merge();
                    xlSheet.Cells[1, 1] = "Ведомость по столовой за " + dateTimePicker1.Value.ToString("dd.MM.yyyy") + "";
                    xlSheetRange = xlSheet.UsedRange;
                    xlSheetRange.Columns.AutoFit();
                    xlSheetRange.Cells.HorizontalAlignment = -4108;
                    xlSheetRange.Cells.VerticalAlignment = -4108;
                    xlSheetRange.Rows.AutoFit();
                    (xlSheetRange.Cells[1, 1] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 1] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 2] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 3] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 4] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 5] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 6] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 7] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 8] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[rowInd + 4, collInd - 4] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[rowInd + 4, collInd - 3] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[rowInd + 4, collInd - 2] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[rowInd + 4, collInd - 1] as Excel.Range).Font.Bold = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                finally
                {

                    xlApp.Visible = true;
                    xlApp.Interactive = true;
                    xlApp.ScreenUpdating = true;
                    xlApp.UserControl = true;
                    releaseObject(xlSheetRange);
                    releaseObject(xlSheet);
                    releaseObject(xlApp);

                }
            }
            else
            {
                MessageBox.Show("Данных по столовой за выбранный период не было!");
            }

        }



        public void ReportShopPerDate()
        {

            string date1 = dateTimePicker1.Value.ToString("dd.MM.yyyy 00:00");
            string date2 = dateTimePicker1.Value.ToString("dd.MM.yyyy 23:59");
            SqlDataAdapter ada2 = new SqlDataAdapter("SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма,val As Собственная,val1 As Готовая,type As Тип,date As Дата FROM dbo.dobor_" + dateTimePicker1.Value.Month + "_" + dateTimePicker1.Value.Year + " Where date between '" + date1 + "' And '" + date2 + "' And type = 518", con);
            DataTable dt2 = new DataTable();
            ada2.Fill(dt2);
            if (dt2.Rows.Count > 0)
            {
                MessageBox.Show("Данные экспортированы в Excel");
                xlApp = new Excel.Application();
                xlApp = new Excel.Application();
                try
                {
                    xlApp.Workbooks.Add(Type.Missing);
                    xlApp.Interactive = false;
                    xlApp.EnableEvents = false;
                    xlSheet = (Excel.Worksheet)xlApp.Sheets[1];
                    xlSheet.Name = "Данные";
                    DataTable dt = GetDataDayShop();
                    int collInd = 0;
                    int rowInd = 0;
                    string data = "";

                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        data = dt.Columns[i].ColumnName.ToString();
                        xlSheet.Cells[2, i + 1] = data;
                        xlSheetRange = xlSheet.get_Range("A2:Z2", Type.Missing);
                    }
                    for (rowInd = 0; rowInd < dt.Rows.Count; rowInd++)
                    {

                        for (collInd = 0; collInd < dt.Columns.Count; collInd++)
                        {
                            data = dt.Rows[rowInd].ItemArray[collInd].ToString();
                            xlSheet.Cells[rowInd + 3, collInd + 1] = data;

                        }
                    }
                    xlSheet.Cells[rowInd + 4, collInd - 2] = "Итоговая сумма:";
                    xlSheet.Cells[rowInd + 4, collInd - 1].Formula = "=Sum(" + xlSheet.Cells[3, 4].Address + ":" + xlSheet.Cells[rowInd + 2, 4].Address + ")";
                    Microsoft.Office.Interop.Excel.Range tRange = xlSheet.UsedRange;
                    tRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    tRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                    xlSheet.get_Range("A1", "E1").Merge();
                    xlSheet.Cells[1, 1] = "Ведомость по магазину за " + dateTimePicker1.Value.ToString("dd.MM.yyyy") + "";
                    xlSheetRange = xlSheet.UsedRange;
                    xlSheetRange.Columns.AutoFit();
                    xlSheetRange.Cells.HorizontalAlignment = -4108;
                    xlSheetRange.Cells.VerticalAlignment = -4108;
                    xlSheetRange.Rows.AutoFit();
                    (xlSheetRange.Cells[1, 1] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 1] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 2] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 3] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 4] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 5] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 6] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 7] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 8] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[rowInd + 4, collInd - 2] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[rowInd + 4, collInd - 1] as Excel.Range).Font.Bold = true;

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                finally
                {

                    xlApp.Visible = true;
                    xlApp.Interactive = true;
                    xlApp.ScreenUpdating = true;
                    xlApp.UserControl = true;
                    releaseObject(xlSheetRange);
                    releaseObject(xlSheet);
                    releaseObject(xlApp);

                }
            }
            else
            {
                MessageBox.Show("Данных по магазину за выбранный период не было!");
            }

        }



        public void TodayReportShop()
        {

            string date = DateTime.UtcNow.ToString("dd.MM.yyyy 23:59");
            SqlDataAdapter ada2 = new SqlDataAdapter("SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма,val As Собственная,val1 As Готовая,type As Тип,date As Дата FROM dbo.dobor_" + DateTime.Now.Month + "_" + DateTime.Now.Year + " Where date between '" + DateTime.Today + "' And '" + date + "' And type = 518", con);
            DataTable dt2 = new DataTable();
            ada2.Fill(dt2);
            if (dt2.Rows.Count > 0)
            {
                MessageBox.Show("Данные экспортированы в Excel");
                xlApp = new Excel.Application();
                xlApp = new Excel.Application();
                try
                {
                    xlApp.Workbooks.Add(Type.Missing);
                    xlApp.Interactive = false;
                    xlApp.EnableEvents = false;
                    xlSheet = (Excel.Worksheet)xlApp.Sheets[1];
                    xlSheet.Name = "Данные";
                    DataTable dt = GetDataTodayShop();
                    int collInd = 0;
                    int rowInd = 0;
                    string data = "";

                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        data = dt.Columns[i].ColumnName.ToString();
                        xlSheet.Cells[2, i + 1] = data;
                        xlSheetRange = xlSheet.get_Range("A2:Z2", Type.Missing);
                    }
                    for (rowInd = 0; rowInd < dt.Rows.Count; rowInd++)
                    {

                        for (collInd = 0; collInd < dt.Columns.Count; collInd++)
                        {
                            data = dt.Rows[rowInd].ItemArray[collInd].ToString();
                            xlSheet.Cells[rowInd + 3, collInd + 1] = data;

                        }
                    }
                    xlSheet.Cells[rowInd + 4, collInd - 2] = "Итоговая сумма:";
                    xlSheet.Cells[rowInd + 4, collInd - 1].Formula = "=Sum(" + xlSheet.Cells[3, 4].Address + ":" + xlSheet.Cells[rowInd + 2, 4].Address + ")";
                    Microsoft.Office.Interop.Excel.Range tRange = xlSheet.UsedRange;
                    tRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    tRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                    string date2 = DateTime.UtcNow.ToString("dd.MM.yyyy");
                    xlSheet.get_Range("A1", "E1").Merge();
                    xlSheet.Cells[1, 1] = "Ведомость по столовой за " + date2 + "";
                    xlSheetRange = xlSheet.UsedRange;
                    xlSheetRange.Columns.AutoFit();
                    xlSheetRange.Cells.HorizontalAlignment = -4108;
                    xlSheetRange.Cells.VerticalAlignment = -4108;
                    xlSheetRange.Rows.AutoFit();
                    (xlSheetRange.Cells[1, 1] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 1] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 2] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 3] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 4] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 5] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 6] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 7] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 8] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[rowInd + 4, collInd - 2] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[rowInd + 4, collInd - 1] as Excel.Range).Font.Bold = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                finally
                {

                    xlApp.Visible = true;
                    xlApp.Interactive = true;
                    xlApp.ScreenUpdating = true;
                    xlApp.UserControl = true;
                    releaseObject(xlSheetRange);
                    releaseObject(xlSheet);
                    releaseObject(xlApp);

                }
            }
            else
            {
                MessageBox.Show("Данных по магазину за сегодня не было!");
            }

        }
        public void TodayReportKitchen()
        {

            string date = DateTime.UtcNow.ToString("dd.MM.yyyy 23:59");
            SqlDataAdapter ada2 = new SqlDataAdapter("SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма,val As Собственная,val1 As Готовая,type As Тип,date FROM dbo.dobor_" + DateTime.Now.Month + "_" + DateTime.Now.Year + " Where date between '" + DateTime.Today + "' And '" + date + "' And type = 521", con);
            DataTable dt2 = new DataTable();
            ada2.Fill(dt2);
            if (dt2.Rows.Count > 0)
            {
                MessageBox.Show("Данные экспортированы в Excel");
                xlApp = new Excel.Application();
                xlApp = new Excel.Application();
                try
                {
                    xlApp.Workbooks.Add(Type.Missing);
                    xlApp.Interactive = false;
                    xlApp.EnableEvents = false;
                    xlSheet = (Excel.Worksheet)xlApp.Sheets[1];
                    xlSheet.Name = "Данные";
                    DataTable dt = GetDataTodaykitchen();
                    int collInd = 0;
                    int rowInd = 0;
                    string data = "";

                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        data = dt.Columns[i].ColumnName.ToString();
                        xlSheet.Cells[2, i + 1] = data;
                        xlSheetRange = xlSheet.get_Range("A2:Z2", Type.Missing);
                    }
                    for (rowInd = 0; rowInd < dt.Rows.Count; rowInd++)
                    {

                        for (collInd = 0; collInd < dt.Columns.Count; collInd++)
                        {
                            data = dt.Rows[rowInd].ItemArray[collInd].ToString();
                            xlSheet.Cells[rowInd + 3, collInd + 1] = data;

                        }
                    }
                    xlSheet.Cells[rowInd + 4, collInd - 4] = "Итого:";
                    xlSheet.Cells[rowInd + 4, collInd - 3].Formula = "=Sum(" + xlSheet.Cells[1, 4].Address + ":" + xlSheet.Cells[rowInd + 2, 4].Address + ")";
                    xlSheet.Cells[rowInd + 4, collInd - 2].Formula = "=Sum(" + xlSheet.Cells[1, 5].Address + ":" + xlSheet.Cells[rowInd + 2, 5].Address + ")";
                    xlSheet.Cells[rowInd + 4, collInd - 1].Formula = "=Sum(" + xlSheet.Cells[1, 6].Address + ":" + xlSheet.Cells[rowInd + 2, 6].Address + ")";
                    Microsoft.Office.Interop.Excel.Range tRange = xlSheet.UsedRange;
                    tRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    tRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                    string date2 = DateTime.UtcNow.ToString("dd.MM.yyyy");
                    xlSheet.get_Range("A1", "G1").Merge();
                    xlSheet.Cells[1, 1] = "Ведомость по столовой за " + date2 + "";
                    xlSheetRange = xlSheet.UsedRange;
                    xlSheetRange.Columns.AutoFit();
                    xlSheetRange.Cells.HorizontalAlignment = -4108;
                    xlSheetRange.Cells.VerticalAlignment = -4108;
                    xlSheetRange.Rows.AutoFit();
                    (xlSheetRange.Cells[1, 1] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 2] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 3] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 4] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 5] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 6] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 7] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 8] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[rowInd + 4, collInd - 4] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[rowInd + 4, collInd - 3] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[rowInd + 4, collInd - 2] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[rowInd + 4, collInd - 1] as Excel.Range).Font.Bold = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                finally
                {

                    xlApp.Visible = true;
                    xlApp.Interactive = true;
                    xlApp.ScreenUpdating = true;
                    xlApp.UserControl = true;
                    releaseObject(xlSheetRange);
                    releaseObject(xlSheet);
                    releaseObject(xlApp);

                }
            }
            else
            {
                MessageBox.Show("Данных по столовой за сегодня не было!");
            }

        }
        public void TodayReport()
        {

            string date5 = DateTime.UtcNow.ToString("dd.MM.yyyy 23:59");
            SqlDataAdapter ada2 = new SqlDataAdapter("SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма,val As Собственная,val1 As Готовая,type As Тип,date FROM dbo.dobor_" + DateTime.Now.Month + "_" + DateTime.Now.Year + " Where date between '" + DateTime.Today + "' And '" + date5 + "'", con);
            DataTable dt2 = new DataTable();
            ada2.Fill(dt2);
            if (dt2.Rows.Count > 0)
            {
                MessageBox.Show("Данные экспортированы в Excel");
                xlApp = new Excel.Application();
                xlApp = new Excel.Application();
                try
                {
                    xlApp.Workbooks.Add(Type.Missing);
                    xlApp.Interactive = false;
                    xlApp.EnableEvents = false;
                    xlSheet = (Excel.Worksheet)xlApp.Sheets[1];
                    xlSheet.Name = "Данные";
                    DataTable dt = GetDataToday();
                    int collInd = 0;
                    int rowInd = 0;
                    string data = "";

                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        data = dt.Columns[i].ColumnName.ToString();
                        xlSheet.Cells[2, i + 1] = data;
                        xlSheetRange = xlSheet.get_Range("A2:Z2", Type.Missing);
                    }
                    for (rowInd = 0; rowInd < dt.Rows.Count; rowInd++)
                    {

                        for (collInd = 0; collInd < dt.Columns.Count; collInd++)
                        {
                            data = dt.Rows[rowInd].ItemArray[collInd].ToString();
                            xlSheet.Cells[rowInd + 3, collInd + 1] = data;

                        }
                    }
                    xlSheet.Cells[rowInd + 4, collInd - 4] = "Итоговая сумма:";
                    xlSheet.Cells[rowInd + 4, collInd - 3].Formula = "=Sum(" + xlSheet.Cells[3, 4].Address + ":" + xlSheet.Cells[rowInd + 2, 4].Address + ")";
                    string date2 = DateTime.UtcNow.ToString("dd.MM.yyyy");
                    Microsoft.Office.Interop.Excel.Range tRange = xlSheet.UsedRange;
                    tRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    tRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                    xlSheet.get_Range("A1", "G1").Merge();
                    xlSheet.Cells[1, 1] = "Общая ведомость за " + date2 + "";
                    xlSheetRange = xlSheet.UsedRange;
                    xlSheetRange.Columns.AutoFit();
                    xlSheetRange.Cells.HorizontalAlignment = -4108;
                    xlSheetRange.Cells.VerticalAlignment = -4108;
                    xlSheetRange.Rows.AutoFit();
                    (xlSheetRange.Cells[1, 1] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 1] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 2] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 3] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 4] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 5] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 6] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 7] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 8] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[rowInd + 4, collInd - 4] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[rowInd + 4, collInd - 3] as Excel.Range).Font.Bold = true;

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                finally
                {

                    xlApp.Visible = true;
                    xlApp.Interactive = true;
                    xlApp.ScreenUpdating = true;
                    xlApp.UserControl = true;
                    releaseObject(xlSheetRange);
                    releaseObject(xlSheet);
                    releaseObject(xlApp);

                }
            }
            else
            {
                MessageBox.Show("Данных за сегодня не было!");
            }

        }
        public void FinalyReport()
        {

            SqlDataAdapter ada2 = new SqlDataAdapter("SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма,val As Собственная,val1 As Готовая,type As Тип, date FROM dbo.dobor_" + DateTime.Now.Month + "_" + DateTime.Now.Year + "", con);
            DataTable dt2 = new DataTable();
            ada2.Fill(dt2);
            if (dt2.Rows.Count > 0)
            {
                MessageBox.Show("Данные экспортированы в Excel");
                xlApp = new Excel.Application();
                xlApp = new Excel.Application();
                try
                {
                    xlApp.Workbooks.Add(Type.Missing);
                    xlApp.Interactive = false;
                    xlApp.EnableEvents = false;
                    xlSheet = (Excel.Worksheet)xlApp.Sheets[1];
                    xlSheet.Name = "Данные";
                    DataTable dt = GetDataTotal();
                    int collInd = 0;
                    int rowInd = 0;
                    string data = "";

                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        data = dt.Columns[i].ColumnName.ToString();
                        xlSheet.Cells[2, i + 1] = data;
                        xlSheetRange = xlSheet.get_Range("A2:Z2", Type.Missing);
                    }
                    for (rowInd = 0; rowInd < dt.Rows.Count; rowInd++)
                    {

                        for (collInd = 0; collInd < dt.Columns.Count; collInd++)
                        {
                            data = dt.Rows[rowInd].ItemArray[collInd].ToString();
                            xlSheet.Cells[rowInd + 3, collInd + 1] = data;

                        }
                    }
                    xlSheet.Cells[rowInd + 4, collInd - 4] = "Итоговая сумма:";
                    xlSheet.Cells[rowInd + 4, collInd - 3].Formula = "=Sum(" + xlSheet.Cells[1, 4].Address + ":" + xlSheet.Cells[rowInd + 2, 4].Address + ")";
                    Microsoft.Office.Interop.Excel.Range tRange = xlSheet.UsedRange;
                    tRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    tRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                    xlSheet.get_Range("A1", "G1").Merge();
                    xlSheet.Cells[1, 1] = "Ведомость за " + DateTime.Now.Month + " месяц " + DateTime.Now.Year + " года";
                    xlSheetRange = xlSheet.UsedRange;
                    xlSheetRange.Columns.AutoFit();
                    xlSheetRange.Cells.HorizontalAlignment = -4108;
                    xlSheetRange.Cells.VerticalAlignment = -4108;
                    xlSheetRange.Rows.AutoFit();
                    (xlSheetRange.Cells[1, 1] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 1] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 2] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 3] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 4] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 5] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 6] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 7] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[2, 8] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[rowInd + 4, collInd - 4] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[rowInd + 4, collInd - 3] as Excel.Range).Font.Bold = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                finally
                {

                    xlApp.Visible = true;
                    xlApp.Interactive = true;
                    xlApp.ScreenUpdating = true;
                    xlApp.UserControl = true;
                    releaseObject(xlSheetRange);
                    releaseObject(xlSheet);
                    releaseObject(xlApp);

                }
            }
            else
            {
                MessageBox.Show("Данных за этот месяц нет!");
            }
        }
        void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show(ex.ToString(), "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
                GC.Collect();
            }
        }



        private void button8_Click(object sender, EventArgs e)
        {
            panel3.Visible = true;
            this.OpenAddPanel2(panel3, _addpanels2);
            panel1.Visible = false;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Ap6Pay();
            Sum();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (radioButton4.Checked)
            {
                SqlDataAdapter ada = new SqlDataAdapter("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'dobor_" + dateTimePicker4.Value.Month + "_" + dateTimePicker4.Value.Year + "'", con);
                DataTable dt = new DataTable();
                ada.Fill(dt);
                if (dt.Rows.Count > 0)
                {
                    GetDataForMonthhShop();
                    Sum();
                }
                else
                {
                    MessageBox.Show("Данных за выбранный период не существует!");
                }
                radioButton4.Checked = false;
            }
            else if (radioButton3.Checked)
            {
                SqlDataAdapter ada = new SqlDataAdapter("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'dobor_" + dateTimePicker4.Value.Month + "_" + dateTimePicker4.Value.Year + "'", con);
                DataTable dt = new DataTable();
                ada.Fill(dt);
                if (dt.Rows.Count > 0)
                {
                    GetDataForMonthKitchen();
                    Sum();
                }
                else
                {
                    MessageBox.Show("Данных за выбранный период не существует!");
                }
                radioButton3.Checked = false;
            }

            else if (MessageBox.Show("Вы хотите вывести общие данные?", "Внимание", MessageBoxButtons.YesNo,
                       MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
            {
                SqlDataAdapter ada = new SqlDataAdapter("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'dobor_" + dateTimePicker4.Value.Month + "_" + dateTimePicker4.Value.Year + "'", con);
                DataTable dt = new DataTable();
                ada.Fill(dt);
                if (dt.Rows.Count > 0)
                {
                    GetDataForMonthh();
                    Sum();
                    //if (GetDataForMonth() is IDisposable)
                    //GetDataForMonth().Dispose();
                }
                else
                {
                    MessageBox.Show("Данных за выбранный период не существует!");
                }
                radioButton3.Checked = false;
                radioButton4.Checked = false;
            }
            else
            {
                MessageBox.Show("Для вывода конкретных данных, выберите код!");
                radioButton3.Checked = false;
                radioButton4.Checked = false;
            }

        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (radioButton4.Checked)
            {
                SqlDataAdapter ada = new SqlDataAdapter("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'dobor_" + dateTimePicker4.Value.Month + "_" + dateTimePicker4.Value.Year + "'", con);
                DataTable dt = new DataTable();
                ada.Fill(dt);
                if (dt.Rows.Count > 0)
                {
                    ReportForMonthShop();
                    radioButton4.Checked = false;
                }
                else
                {
                    MessageBox.Show("Данных за выбранный период не существует!");
                }
                    radioButton3.Checked = false;
                    radioButton4.Checked = false;
                }
            else if (radioButton3.Checked)
            {
                SqlDataAdapter ada = new SqlDataAdapter("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'dobor_" + dateTimePicker4.Value.Month + "_" + dateTimePicker4.Value.Year + "'", con);
                DataTable dt = new DataTable();
                ada.Fill(dt);
                if (dt.Rows.Count > 0)
                {
                    ReportForMonthKitchen();
                    radioButton3.Checked = false;
                }
                else
                {
                    MessageBox.Show("Данных за выбранный период не существует!");
                }
                radioButton3.Checked = false;
                radioButton4.Checked = false;
            }
            else if (MessageBox.Show("Вы хотите напечатать общие данные?", "Внимание", MessageBoxButtons.YesNo,
                         MessageBoxIcon.Question) == DialogResult.Yes)
            {

                SqlDataAdapter ada = new SqlDataAdapter("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'dobor_" + dateTimePicker4.Value.Month + "_" + dateTimePicker4.Value.Year + "'", con);
                DataTable dt = new DataTable();
                ada.Fill(dt);
                if (dt.Rows.Count > 0)
                {
                    ReportForMonth();
                }
                else
                {
                    MessageBox.Show("Данных за выбранный период не существует!");
                }
                radioButton3.Checked = false;
                radioButton4.Checked = false;
            }
            else
            {
                MessageBox.Show("Для печати конкретных данных, выберите код!");
                radioButton3.Checked = false;
                radioButton4.Checked = false;
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            panel4.Visible = true;
            this.OpenAddPanel2(panel4, _addpanels2);
            panel3.Visible = false;
            maskedTextBox1.Focus();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (maskedTextBox1.Text == "")
            {
                MessageBox.Show("Вы не ввели табельный!");
            }
            else
            {
                switch (MessageBox.Show("Вы хотите вывести данные за месяц?",
              "Внимание",
              MessageBoxButtons.YesNoCancel,
              MessageBoxIcon.Question))
                {
                    case DialogResult.Yes:
                        GetDataTabell();
                        Sum();
                        maskedTextBox1.Clear();
                        maskedTextBox1.Focus();
                        break;

                    case DialogResult.No:
                        GetDataTabell78();
                        Sum();
                        if (dataGridView1.Rows == null || dataGridView1.Rows.Count == 0)
                        {
                            MessageBox.Show("По данному табельному данных за выбранный день не существует!");
                            Ap6Pay();
                            Sum();
                            maskedTextBox1.Clear();
                            maskedTextBox1.Focus();
                        }
                        break;

                    case DialogResult.Cancel:

                        break;
                }

            }

        }


        private void button10_Click(object sender, EventArgs e)
        {
            Ap6Pay();
            Sum();
            maskedTextBox1.Clear();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (maskedTextBox1.Text == "") {

                MessageBox.Show("Вы не ввели табельный номер!");            
            } else {

                switch (MessageBox.Show("Вы хотите экспортировать данные за месяц?",
                     "Внимание",
                     MessageBoxButtons.YesNoCancel,
                     MessageBoxIcon.Question))
                {
                    case DialogResult.Yes:
                        SqlDataAdapter ada = new SqlDataAdapter("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'dobor_" + dateTimePicker2.Value.Month + "_" + dateTimePicker2.Value.Year + "'", con);
                        DataTable dt = new DataTable();
                        ada.Fill(dt);
                        if (dt.Rows.Count > 0)
                        {
                          
                            //progressExportToDbf.PerformStep();
                            ReportTabel();
                            maskedTextBox1.Clear();
                            maskedTextBox1.Focus();
                        }
                        else
                        {
                            MessageBox.Show("Данных за выбранный период не существует!");
                            maskedTextBox1.Clear();
                            maskedTextBox1.Focus();
                        }

                        break;

                    case DialogResult.No:
                        SqlDataAdapter ada2 = new SqlDataAdapter("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'dobor_" + dateTimePicker2.Value.Month + "_" + dateTimePicker2.Value.Year + "'", con);
                        DataTable dt2 = new DataTable();
                        ada2.Fill(dt2);
                        if (dt2.Rows.Count > 0)
                        {
                            ReportTabel2();
                            maskedTextBox1.Clear();
                            maskedTextBox1.Focus();
                        }
                        else
                        {
                            MessageBox.Show("Данных за выбранный период не существует!");
                            maskedTextBox1.Clear();
                                        maskedTextBox1.Focus();
                        }
                        break;

                    case DialogResult.Cancel:

                        break;
                }
            }
        
        }


        public void Sum()
        {

            var foundRows = this.dataGridView1.Rows.Cast<DataGridViewRow>().Where(row => row.Cells["Код"].Value.ToString() == "518");

            double summkitchen = 0;
            foreach (DataGridViewRow dataGrid in foundRows)
            {
                double incom5;
                double.TryParse((dataGrid.Cells[3].Value ?? "0").ToString().Replace(".", ","), out incom5);
                summkitchen += incom5;
            }
            
            

            double summa = 0;
            double val1 = 0;
            double val = 0;          

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                
                double incom;
                double incom1;
                double incom2;
                double.TryParse((row.Cells[3].Value ?? "0").ToString().Replace(".", ","), out incom);
                double.TryParse((row.Cells[5].Value ?? "0").ToString().Replace(".", ","), out incom1);
                double.TryParse((row.Cells[4].Value ?? "0").ToString().Replace(".", ","), out incom2);
                summa += incom; 
                val1 += incom1;
                val += incom2;

            }
            
            label7.Text = Math.Round(summa,2).ToString() + " руб.";
            label9.Text = Math.Round(val1, 2).ToString() + " руб.";
            label8.Text = Math.Round(val,2).ToString() + " руб.";
            label13.Text = Math.Round(summkitchen, 2).ToString() + " руб.";


        }


        private void button13_Click(object sender, EventArgs e)
        {
            
            daoService.Month = dateTimePicker1.Value.Month.ToString();
            daoService.Year = dateTimePicker1.Value.Year.ToString();
            SqlDataAdapter ada2 = new SqlDataAdapter("SELECT Id, tab_no, fio, summa, val, val1, type, date FROM dbo.dobor_" + dateTimePicker1.Value.Month + "_" + dateTimePicker1.Value.Year + " Where date between '" + dateTimePicker1.Value.ToString("dd.MM.yyyy 00:00") + "' And '" + dateTimePicker1.Value.ToString("dd.MM.yyyy 23:59") + "' And (Type = '518' OR  Type = '521')", con);
            DataTable dt2 = new DataTable();
            ada2.Fill(dt2);

            if (dt2.Rows.Count > 0)
            {
                daoService.date = dateTimePicker1.Value.ToString("dd.MM.yyyy");
                daoService.date1 = dateTimePicker1.Value.ToString("dd.MM.yyyy 00:00");
                daoService.date2 = dateTimePicker1.Value.ToString("dd.MM.yyyy 23:59");
                Ap6Pay3();
                if (pays.Count == 0)
                {
                    MessageBox.Show("Список для экспорта пуст!");
                    return;
                }        
                string date3 =  dateTimePicker1.Value.ToString("ddMMyy");
                string filePath = "D:\\Отчёты";
                string fileName = "DB" + date3; /*DateTime.UtcNow.ToString("ddMMyy");*/
                if (File.Exists(Path.Combine(filePath, fileName + ".dbf")))
                {
                    if (MessageBox.Show("Файл уже существует! Удалить старый файл?", "Внимание", MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question) == DialogResult.No)
                    {
                        return;
                    }
                    try
                    {
                        File.Delete(Path.Combine(filePath, fileName + ".dbf"));
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString(), "Ошибка удаления файла!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                }
                progressExportToDbf.Value = 0;
                progressExportToDbf.Visible = true;
                progressExportToDbf.Maximum = pays.Count;
                progressExportToDbf.Step = 1;
                progressExportToDbf.Minimum = 0;

                try
                {
                    daoService.ExportDataToDbf(pays.ToList(), filePath, fileName, () => progressExportToDbf.PerformStep());
                    MessageBox.Show("Экспорт выполнен успешно!");


                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), "Ошибка экспорта!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                finally
                {
                    progressExportToDbf.Visible = false;
                }
            } else
            {
                MessageBox.Show("Данных для выгрузки за выбранный день не существует!");
            }
            Ap6Pay();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
            this.WindowState = FormWindowState.Normal;
            this.Focus(); this.Show();
        }

        private void MonthlyReportGeneralButton_Click(object sender, EventArgs e)
        {
            FinalyReport();
        }

        private void MonthlyReportKitchenButton_Click(object sender, EventArgs e)
        {
            ReportForkitchenproducts();
        }

        private void MonthlyReportShopButton_Click(object sender, EventArgs e)
        {
            ReportShop();
        }

        private void ReportForTodayGeneralButton_Click(object sender, EventArgs e)
        {
            TodayReport();
        }

        private void ReportForTodayKitchenButton_Click(object sender, EventArgs e)
        {
            TodayReportKitchen();
        }

        private void ReportForTodayShopButton_Click(object sender, EventArgs e)
        {
            TodayReportShop();
        }

        private void CloseAllTabsButton_Click(object sender, EventArgs e)
        {
            panelhide();
        }

        private void UpdateDataForSecondButton_Click(object sender, EventArgs e)
        {
            Ap6Pay();
            Sum();
        }

        private void SelectDataForASpecificDateButton_Click(object sender, EventArgs e)
        {
            panel1.Visible = true;
            this.OpenAddPanel2(panel1, _addpanels2);
            panel3.Visible = false;
        }

        private void ViewingDataInATableForDateButton_Click(object sender, EventArgs e)
        {
            //daoService.CloseConnection();
            //dataGridView1.Rows.Clear();

            SqlDataAdapter ada = new SqlDataAdapter("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'dobor_" + dateTimePicker1.Value.Month + "_" + dateTimePicker1.Value.Year + "'", con);
            DataTable dt = new DataTable();
            ada.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                SqlDataAdapter ada2 = new SqlDataAdapter("SELECT Id, tab_no, fio, summa, val, val1, type, date FROM dbo.dobor_" + dateTimePicker1.Value.Month + "_" + dateTimePicker1.Value.Year + " Where date between '" + dateTimePicker1.Value.ToString("dd.MM.yyyy 00:00") + "' And '" + dateTimePicker1.Value.ToString("dd.MM.yyyy 23:59") + "' And (Type = '518'  OR  Type = '521')", con);
                DataTable dt2 = new DataTable();
                ada2.Fill(dt2);

                if (dt2.Rows.Count > 0)
                {

                    // 3 одниковых sql запроса в одном методе - 3 инфаркта у Курочки
                    if (radioButton1.Checked)
                    {
                        GetDataForPerDateShop();
                        Sum();
                        if (dataGridView1.Rows == null || dataGridView1.Rows.Count == 0)
                        {
                            MessageBox.Show("Данных за выбранный день не существует!");
                            Ap6Pay();
                            Sum();
                        }
                        //dataGridView1.AutoGenerateColumns = true;
                        //dataGridView1.DataSource = new BindingSource(pays, null);
                        //dataGridView1.Sort(dataGridView1.Columns[0], ListSortDirection.Ascending);
                        //foreach (var pay in daoService.GetDataForASpecificDateForShop())
                        //{
                        //    pays.Add(pay);
                        //    dataGridView1.Columns["Id"].Visible = false;
                        //}

                    }
                    else if (radioButton2.Checked)
                    {
                        GetDataForPerDateKitchen();
                        Sum();
                        //dataGridView1.AutoGenerateColumns = true;
                        //dataGridView1.DataSource = new BindingSource(pays, null);
                        //dataGridView1.Sort(dataGridView1.Columns[0], ListSortDirection.Ascending);
                        //foreach (var pay in daoService.GetDataForASpecificDateForKitchen())
                        //{
                        //    pays.Add(pay);
                        //    dataGridView1.Columns["Id"].Visible = false;
                        //}
                        if (dataGridView1.Rows == null || dataGridView1.Rows.Count == 0)
                        {
                            MessageBox.Show("Данных за выбранный день не существует!");
                            Ap6Pay();
                            Sum();
                        }

                    }
                    else 
                    {
                        GetDataForPerDate();
                        Sum();
                        if (dataGridView1.Rows == null || dataGridView1.Rows.Count == 0)
                        {
                            MessageBox.Show("Данных за выбранный день не существует!");
                            Ap6Pay();
                            Sum();
                        }
                    }
                   
                }
                else
                {
                    MessageBox.Show("Данных за выбранный период не существует!");
                }

            }
            else
            {
                MessageBox.Show("Данных за выбранный период не существует!");
            }
        }


        private void CancelViewingDataForDateButton_Click(object sender, EventArgs e)
        {
            //dataGridView1.Rows.Clear();
            //Ap6Pay2();
            Ap6Pay();
            Sum();
            radioButton1.Checked = false;
            radioButton2.Checked = false;
        }

        private void ReportDataInExcelForASpecificDateButton_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                ReportShopPerDate();
                radioButton1.Checked = false;
            }
            else if (radioButton2.Checked)
            {
                ReportKitchenPerDate();
                radioButton2.Checked = false;
            }
            else if (MessageBox.Show("Вы хотите распечатать общие данные?", "Внимание", MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
            {
                ReportDate();
                radioButton1.Checked = false;
                radioButton2.Checked = false;
            }
            else
            {
                MessageBox.Show("Для печати конкретных данных, выберите код!");
            }
        }

        private void Form1_DoubleClick(object sender, EventArgs e)
        {
           //fullScreen.Toggle();
        }

    }
}



