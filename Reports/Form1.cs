using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.IO;

namespace Reports
{
    public partial class Form1 : Form
    {
        private readonly List<Panel> _addpanels2;
        Excel.Application xlApp;
        Excel.Worksheet xlSheet;
        Excel.Range xlSheetRange;
        public SqlConnection con;
        public string date = DateTime.UtcNow.ToString("dd.MM.yyyy 23:59");

        public Form1()
        {
                con = new SqlConnection("Data Source=86.57.137.8,1433;Initial Catalog=ap6pay;Persist Security Info=True;User ID=admin;Password=682830");
                InitializeComponent();
                _addpanels2 = new List<Panel> { panel1, panel3, panel4 };
                Ap6Pay();
                Sum();

        }

        public void GetDataTabell78()
        {

                string date6 = dateTimePicker2.Value.ToString("dd.MM.yyyy 00:00");
                string date7 = dateTimePicker2.Value.ToString("dd.MM.yyyy 23:59");
                string command = "SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма, val As Собственная,val1 As Готовая,type As Тип,date As Дата FROM dbo.dobor_" + dateTimePicker2.Value.Month + "_" + dateTimePicker2.Value.Year + " Where  tab_no='" + maskedTextBox1.Text + "' And date between '" + date6 + "' And '" + date7 + "'";
                SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command, con );
                DataSet dataSet = new DataSet();
                sqlDataAdapter.Fill(dataSet);
                dataGridView1.DataSource = dataSet.Tables[0].DefaultView;

        }

        public void GetDataTabell()
        {

                string command = "SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма, val As Собственная,val1 As Готовая,type As Тип,date As Дата FROM dbo.dobor_" + DateTime.Now.Month + "_" + DateTime.Now.Year + " Where  tab_no='" + maskedTextBox1.Text + "'";
                SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command, con);
                DataSet dataSet = new DataSet();
                sqlDataAdapter.Fill(dataSet);
                dataGridView1.DataSource = dataSet.Tables[0].DefaultView;


        }

        public void GetDataForMonthKitchen()
        {

                string command = "SELECT Id, tab_no As Табельный,fio As ФИО ,summa As Сумма,val As Собственная,val1 As Готовая,type As Код,date As Дата FROM dbo.dobor_" + dateTimePicker4.Value.Month + "_" + dateTimePicker4.Value.Year + " Where type = 521";
                SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command, con);
                DataSet dataSet = new DataSet();
                sqlDataAdapter.Fill(dataSet);
                dataGridView1.DataSource = dataSet.Tables[0].DefaultView;


        }


        public void GetDataForMonthhShop()
        {

                string command = "SELECT Id, tab_no As Табельный,fio As ФИО ,summa As Сумма,val As Собственная,val1 As Готовая,type As Код,date As Дата FROM dbo.dobor_" + dateTimePicker4.Value.Month + "_" + dateTimePicker4.Value.Year + " Where type = 518";
                SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command, con);
                DataSet dataSet = new DataSet();
                sqlDataAdapter.Fill(dataSet);
                dataGridView1.DataSource = dataSet.Tables[0].DefaultView;

        }
        public void GetDataForMonthh()
        {

                string command = "SELECT Id, tab_no As Табельный,fio As ФИО ,summa As Сумма,val As Собственная,val1 As Готовая,type As Код,date As Дата FROM dbo.dobor_" + dateTimePicker4.Value.Month + "_" + dateTimePicker4.Value.Year + "";
                SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command, con);
                DataSet dataSet = new DataSet();
                sqlDataAdapter.Fill(dataSet);
                dataGridView1.DataSource = dataSet.Tables[0].DefaultView;

        }


        public void Ap6Pay78778()
        {

                SqlDataAdapter ada = new SqlDataAdapter("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'dobor_" + DateTime.Now.Month + "_" + DateTime.Now.Year + "'", con);
                DataTable dt = new DataTable();
                ada.Fill(dt);
                if (dt.Rows.Count > 0)
                {
                    string date9 = DateTime.UtcNow.ToString("dd.MM.yyyy 00:00");
                    string date10 = DateTime.UtcNow.ToString("dd.MM.yyyy 23:59");

                    string command = "SELECT tab_no ,summa,type,date FROM dbo.dobor_" + DateTime.Now.Month + "_" + DateTime.Now.Year + " where date between '" + date9 + "' and '" + date10 + "'";
                    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command, con);
                    DataSet dataSet = new DataSet();
                    sqlDataAdapter.Fill(dataSet);
                    dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
                }
                else
                {
                    MessageBox.Show("Данных за текущий месяц не существует!");
                }

        }


        public void Ap6Pay()
        {

                SqlDataAdapter ada = new SqlDataAdapter("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'dobor_" + DateTime.Now.Month + "_" + DateTime.Now.Year + "'", con);
                DataTable dt = new DataTable();
                ada.Fill(dt);
                if (dt.Rows.Count > 0)
                {
                    string command = "SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма,val As Собственная,val1 As Готовая,type As Код,date As Дата FROM dbo.dobor_" + DateTime.Now.Month + "_" + DateTime.Now.Year + "";
                    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command, con);
                    DataSet dataSet = new DataSet();
                    sqlDataAdapter.Fill(dataSet);
                    dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
                }
                else
                {
                    MessageBox.Show("Данных за текущий месяц не существует!");
                }

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
                    con.Dispose();
                }
                return dt;
        }
        private DataTable GetDataTabel()
        {
                DataTable dt = new DataTable();
                try
                {
                    string query = @"SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма, val As Собственная,val1 As Готовая,type As Тип,date As Дата FROM dbo.dobor_" + DateTime.Now.Month + "_" + DateTime.Now.Year + " Where  tab_no='" + maskedTextBox1.Text + "'";
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
                    con.Dispose();
                }
                return dt;
        }

        private DataTable GetDataShop()
        {

                DataTable dt = new DataTable();
                try
                {
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
                    con.Dispose();
                }
                return dt;
        }



        private DataTable GetDataForMonthShop()
        {
                DataTable dt = new DataTable();
                try
                {
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
                    con.Dispose();
                }
                return dt;
        }


        private DataTable GetDataForMonthKitchenDataPicker()
        {
                DataTable dt = new DataTable();
                try
                {
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
                    con.Dispose();
                }
                return dt;
        }



        private DataTable GetDataForMonth()
        {

                DataTable dt = new DataTable();
                try
                {
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
                    con.Dispose();
                }
                return dt;
        }

        private DataTable GetDataTodayShop()
        {
                DataTable dt = new DataTable();
                try
                {

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
                    con.Dispose();
                }
                return dt;
        }

        private DataTable GetDataTodaykitchen()
        {

                DataTable dt = new DataTable();
                try
                {
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
                    con.Dispose();
                }
                return dt;
        }
        private DataTable GetDataToday()
        {

                DataTable dt = new DataTable();
                try
                {
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
                    con.Dispose();
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
                    con.Dispose();
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
                    con.Dispose();
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
                    con.Dispose();
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
                    con.Dispose();
                }
                return dt;
        }
        private void button3_Click(object sender, EventArgs e)
        {
            panel1.Visible = true;
            this.OpenAddPanel2(panel1, _addpanels2);
            panel3.Visible = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {


                if (radioButton1.Checked)
                {

                    SqlDataAdapter ada = new SqlDataAdapter("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'dobor_" + dateTimePicker1.Value.Month + "_" + dateTimePicker1.Value.Year + "'", con);
                    DataTable dt = new DataTable();
                    ada.Fill(dt);
                    if (dt.Rows.Count > 0)
                    {
                        string date1 = dateTimePicker1.Value.ToString("dd.MM.yyyy 00:00");
                        string date2 = dateTimePicker1.Value.ToString("dd.MM.yyyy 23:59");
                        SqlDataAdapter ada2 = new SqlDataAdapter("SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма, val As Собственная,val1 As Готовая,type As Тип,date As Дата FROM dbo.dobor_" + dateTimePicker1.Value.Month + "_" + dateTimePicker1.Value.Year + " Where date between '" + date1 + "' And '" + date2 + "'", con);
                        DataTable dt2 = new DataTable();
                        ada2.Fill(dt2);
                        if (dt2.Rows.Count > 0)
                        {
                            string command = "SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма, val As Собственная,val1 As Готовая,type As Тип,date As Дата FROM dbo.dobor_" + dateTimePicker1.Value.Month + "_" + dateTimePicker1.Value.Year + " Where date between '" + date1 + "' And '" + date2 + "' And Type = 518";
                            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command, con);
                            DataSet dataSet = new DataSet();
                            sqlDataAdapter.Fill(dataSet);
                            dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
                            Sum();
                            radioButton1.Checked = false;
                        }
                        else
                        {
                            MessageBox.Show("Данных за выбранный период не существует!");
                            radioButton1.Checked = false;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Данных за выбранный период не существует!");
                        radioButton1.Checked = false;
                    }
                }
                else if (radioButton2.Checked)
                {
                    SqlDataAdapter ada = new SqlDataAdapter("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'dobor_" + dateTimePicker1.Value.Month + "_" + dateTimePicker1.Value.Year + "'", con);
                    DataTable dt = new DataTable();
                    ada.Fill(dt);
                    if (dt.Rows.Count > 0)
                    {
                        string date1 = dateTimePicker1.Value.ToString("dd.MM.yyyy 00:00");
                        string date2 = dateTimePicker1.Value.ToString("dd.MM.yyyy 23:59");
                        SqlDataAdapter ada2 = new SqlDataAdapter("SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма, val As Собственная,val1 As Готовая,type As Тип,date As Дата FROM dbo.dobor_" + dateTimePicker1.Value.Month + "_" + dateTimePicker1.Value.Year + " Where date between '" + date1 + "' And '" + date2 + "'", con);
                        DataTable dt2 = new DataTable();
                        ada2.Fill(dt2);
                        if (dt2.Rows.Count > 0)
                        {
                            string command = "SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма, val As Собственная,val1 As Готовая,type As Тип,date As Дата FROM dbo.dobor_" + dateTimePicker1.Value.Month + "_" + dateTimePicker1.Value.Year + " Where date between '" + date1 + "' And '" + date2 + "' And Type = 521";
                            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command, con);
                            DataSet dataSet = new DataSet();
                            sqlDataAdapter.Fill(dataSet);
                            dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
                            Sum();
                            radioButton2.Checked = false;
                        }
                        else
                        {
                            MessageBox.Show("Данных за выбранный период не существует!");
                            radioButton2.Checked = false;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Данных за выбранный период не существует!");
                        radioButton2.Checked = false;
                    }
                }
                else
                if (MessageBox.Show("Вы хотите вывести общие данные?", "Внимание", MessageBoxButtons.YesNo,
                          MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
                {
                    SqlDataAdapter ada = new SqlDataAdapter("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'dobor_" + dateTimePicker1.Value.Month + "_" + dateTimePicker1.Value.Year + "'", con);
                    DataTable dt = new DataTable();
                    ada.Fill(dt);
                    if (dt.Rows.Count > 0)
                    {
                        string date1 = dateTimePicker1.Value.ToString("dd.MM.yyyy 00:00");
                        string date2 = dateTimePicker1.Value.ToString("dd.MM.yyyy 23:59");
                        SqlDataAdapter ada2 = new SqlDataAdapter("SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма, val As Собственная,val1 As Готовая,type As Тип,date As Дата FROM dbo.dobor_" + dateTimePicker1.Value.Month + "_" + dateTimePicker1.Value.Year + " Where date between '" + date1 + "' And '" + date2 + "'", con);
                        DataTable dt2 = new DataTable();
                        ada2.Fill(dt2);
                        if (dt2.Rows.Count > 0)
                        {
                            string command = "SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма, val As Собственная,val1 As Готовая,type As Тип,date As Дата FROM dbo.dobor_" + dateTimePicker1.Value.Month + "_" + dateTimePicker1.Value.Year + " Where date between '" + date1 + "' And '" + date2 + "'";
                            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command, con);
                            DataSet dataSet = new DataSet();
                            sqlDataAdapter.Fill(dataSet);
                            dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
                            Sum();
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
                else
                {
                    MessageBox.Show("Чтобы выводить определённые данные, выберите код!!");
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
                        xlSheet.Cells[1, i + 1] = data;
                        xlSheetRange = xlSheet.get_Range("A2:Z2", Type.Missing);
                    }

                    for (rowInd = 0; rowInd < dt.Rows.Count; rowInd++)
                    {

                        for (collInd = 0; collInd < dt.Columns.Count; collInd++)
                        {
                            data = dt.Rows[rowInd].ItemArray[collInd].ToString();
                            xlSheet.Cells[rowInd + 2, collInd + 1] = data;

                        }

                    }
                    xlSheet.Cells[rowInd + 3, collInd - 5] = "Итоговая сумма:";
                    xlSheet.Cells[rowInd + 3, collInd - 4].Formula = "=Sum(" + xlSheet.Cells[1, 4].Address + ":" + xlSheet.Cells[rowInd + 2, 4].Address + ")";
                    Microsoft.Office.Interop.Excel.Range tRange = xlSheet.UsedRange;
                    tRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    tRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                    xlSheetRange = xlSheet.UsedRange;
                    xlSheetRange.Columns.AutoFit();
                    xlSheetRange.Cells.HorizontalAlignment = -4108;
                    xlSheetRange.Cells.VerticalAlignment = -4108;
                    xlSheetRange.Rows.AutoFit();
                    (xlSheetRange.Cells[1, 1] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[1, 2] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[1, 3] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[1, 4] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[1, 5] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[1, 6] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[1, 7] as Excel.Range).Font.Bold = true;
                    (xlSheetRange.Cells[1, 8] as Excel.Range).Font.Bold = true;
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

        }



        public void ReportForMonthShop()
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
        }


        public void ReportForMonth()
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
                    MessageBox.Show("Данных по магвзину за месяц не было!");
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
                SqlDataAdapter ada2 = new SqlDataAdapter("SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма,val As Собственная,val1 As Готовая,type As Тип,date As Дата FROM dbo.dobor_" + DateTime.Now.Month + "_" + DateTime.Now.Year + " Where date between '" + date1 + "' And '" + date2 + "' And type = 518", con);
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
                SqlDataAdapter ada2 = new SqlDataAdapter("SELECT Id, tab_no As Табельный,fio As ФИО,summa As Сумма,val As Собственная,val1 As Готовая,type As Тип,date As Дата FROM dbo.dobor_" + DateTime.Now.Month + "_" + DateTime.Now.Year + " Where date between '" + date1 + "' And '" + date2 + "' And type = 518", con);
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
        private void отчётЗаМесяцToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FinalyReport();
        }

        private void отчётЗаМесяцПоСтоловойToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            button3.Visible = true;
            panel1.Visible = false;
        }



        private void button1_Click_1(object sender, EventArgs e)
        {
            Ap6Pay();
            Sum();
            radioButton1.Checked = false;
            radioButton2.Checked = false;
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
                ReportForMonthShop();
                radioButton4.Checked = false;
            }
            else if (radioButton3.Checked)
            {
                ReportForMonthKitchen();
                radioButton3.Checked = false;
            }
            else if (MessageBox.Show("Вы хотите напечатать общие данные?", "Внимание", MessageBoxButtons.YesNo,
                         MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
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

        private void отчётЗаМесяцПоГотовойПродукцииToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ReportShop();

        }

        private void button9_Click(object sender, EventArgs e)
        {
            panel4.Visible = true;
            this.OpenAddPanel2(panel4, _addpanels2);
            panel3.Visible = false;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (maskedTextBox1.Text == "")
            {
                MessageBox.Show("Вы не ввели табель!");
            }
            else
            {
                if (MessageBox.Show("Вы хотите вывести данные за месяц?", "Внимание", MessageBoxButtons.YesNo,
                          MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
                {
                    GetDataTabell();
                    Sum();
                    maskedTextBox1.Clear();
                }
                else
                {
                    GetDataTabell78();
                    Sum();
                    if (dataGridView1.Rows == null || dataGridView1.Rows.Count == 0)
                    {
                        MessageBox.Show("По данному табельному данных не существует!");
                        Ap6Pay();
                        Sum();
                        maskedTextBox1.Clear();
                    }
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
            ReportTabel();
            maskedTextBox1.Clear();
        }

        private void закрытьВкладкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panelhide();
        }

        private void отчётЗаСегодняобщийToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TodayReport();
        }

        private void отчётЗаСегодняПоПродукциистоловаяToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TodayReportKitchen();
        }

        private void button4_Click(object sender, EventArgs e)
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

        private void отчётЗаСегоднямагазинToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TodayReportShop();
        }


        private void отчётЗаМесяцПоСтоловойToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            ReportForkitchenproducts();
        }


        public void Sum()
        {
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
            label7.Text = summa.ToString() + " руб.";
            label9.Text = val1.ToString() + " руб.";
            label8.Text = val.ToString() + " руб.";

        }

        private void обновитьДанныеНаТекущуюСекундуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Ap6Pay();
            Sum();
        }
  

        private void button13_Click(object sender, EventArgs e)
        {

            //SqlCommand command = new SqlCommand();
            //con.Open();
            //command.CommandText = "SELECT Id, tab_no,summa,date FROM dbo.dobor_" + DateTime.Now.Month + "_" + DateTime.Now.Year + "";
            //command.Connection = con;
            //SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command);
            //DataTable dt = new DataTable();
            //sqlDataAdapter.Fill(dt);
            //string date2 = DateTime.UtcNow.ToString("ddMMyy");
            //Spire.DataExport.DBF.DBFExport DBFExport = new Spire.DataExport.DBF.DBFExport();
            //DBFExport.DataSource = Spire.DataExport.Common.ExportSource.DataTable;
            //DBFExport.DataTable = dt;
            //DBFExport.ActionAfterExport = Spire.DataExport.Common.ActionType.OpenView;
            //DBFExport.FileName = "D:/Reports/DM" + date2 + ".dbf";
            //DBFExport.SaveToFile();
            //con.Close();




            System.Windows.Forms.MessageBox.Show("Export Complete.", "Program Info", MessageBoxButtons.OK, MessageBoxIcon.Information);

            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = "SELECT tab_no ,summa,type FROM dbo.dobor_" + 11 + "_" + 2022 + " Where type = 518";
                con.Open();
                using (var reader = cmd.ExecuteReader())
                {
                    string date2 = DateTime.UtcNow.ToString("ddMMyy");
                    string date3 = DateTime.UtcNow.ToString("MM");
                    string date4 = DateTime.UtcNow.ToString("yy");
                    using (var writer = new StreamWriter(@"D:\Reports\DM" + date2 + ".txt"))
                    {
                        while (reader.Read())
                        {
                            writer.WriteLine(reader[0].ToString() + "   " + reader[1].ToString() + "\t  " + reader[2].ToString() + "   " + date3 + "   " + date4);
                        }
                    }
                }
                cmd.CommandText = "SELECT tab_no ,summa,type FROM dbo.dobor_" + 11 + "_" + 2022 + " Where type = 521";

                using (var reader2 = cmd.ExecuteReader())
                {
                    string date2 = DateTime.UtcNow.ToString("ddMMyy");
                    string date3 = DateTime.UtcNow.ToString("MM");
                    string date4 = DateTime.UtcNow.ToString("yy");
                    using (var writer = new StreamWriter(@"D:\Reports\DS" + date2 + ".txt"))
                    {
                        while (reader2.Read())
                        {

                            writer.WriteLine(reader2[0].ToString() + "     " + reader2[1].ToString() + "\t  " + reader2[2].ToString() + "    " + date3 + "    " + date4);
                        }
                    }
                    con.Close();
                }

            }
        }
    }
}
    public static class PanelExtensions2
    {
        public static void OpenPanel2(this Panel panel)
        {
            panel.Dock = DockStyle.Fill;
            panel.Visible = true;
        }
        public static void ClosePanel2(this Panel panel)
        {
            panel.Dock = DockStyle.None;
            panel.Visible = false;
        }
        public static void OpenAddPanel2(this Form form, Panel target, List<Panel> addPanels)
        {
            addPanels.Except(new List<Panel> { target }).ToList().ForEach(x =>
            {
                x.ClosePanel2();
            });

            target.OpenPanel2();
        }
    }
    

