namespace Reports
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.отчётыToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.MonthlyReportGeneralButton = new System.Windows.Forms.ToolStripMenuItem();
            this.MonthlyReportKitchenButton = new System.Windows.Forms.ToolStripMenuItem();
            this.MonthlyReportShopButton = new System.Windows.Forms.ToolStripMenuItem();
            this.отчётыПоДнямToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ReportForTodayGeneralButton = new System.Windows.Forms.ToolStripMenuItem();
            this.ReportForTodayKitchenButton = new System.Windows.Forms.ToolStripMenuItem();
            this.ReportForTodayShopButton = new System.Windows.Forms.ToolStripMenuItem();
            this.CloseAllTabsButton = new System.Windows.Forms.ToolStripMenuItem();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.panel1 = new System.Windows.Forms.Panel();
            this.radioButton2 = new System.Windows.Forms.RadioButton();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            this.label1 = new System.Windows.Forms.Label();
            this.ViewingDataInATableForDateButton = new System.Windows.Forms.Button();
            this.CancelViewingDataForDateButton = new System.Windows.Forms.Button();
            this.ReportDataInExcelForASpecificDateButton = new System.Windows.Forms.Button();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.panel4 = new System.Windows.Forms.Panel();
            this.dateTimePicker2 = new System.Windows.Forms.DateTimePicker();
            this.label11 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.maskedTextBox1 = new System.Windows.Forms.MaskedTextBox();
            this.button10 = new System.Windows.Forms.Button();
            this.button11 = new System.Windows.Forms.Button();
            this.button12 = new System.Windows.Forms.Button();
            this.panel3 = new System.Windows.Forms.Panel();
            this.radioButton3 = new System.Windows.Forms.RadioButton();
            this.radioButton4 = new System.Windows.Forms.RadioButton();
            this.label2 = new System.Windows.Forms.Label();
            this.button5 = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            this.button7 = new System.Windows.Forms.Button();
            this.dateTimePicker4 = new System.Windows.Forms.DateTimePicker();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label13 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.SelectDataForASpecificDateButton = new System.Windows.Forms.Button();
            this.button8 = new System.Windows.Forms.Button();
            this.button9 = new System.Windows.Forms.Button();
            this.button13 = new System.Windows.Forms.Button();
            this.progressExportToDbf = new System.Windows.Forms.ProgressBar();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel4.SuspendLayout();
            this.panel3.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.отчётыToolStripMenuItem,
            this.отчётыПоДнямToolStripMenuItem,
            this.CloseAllTabsButton});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(977, 24);
            this.menuStrip1.TabIndex = 6;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // отчётыToolStripMenuItem
            // 
            this.отчётыToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.MonthlyReportGeneralButton,
            this.MonthlyReportKitchenButton,
            this.MonthlyReportShopButton});
            this.отчётыToolStripMenuItem.Name = "отчётыToolStripMenuItem";
            this.отчётыToolStripMenuItem.Size = new System.Drawing.Size(111, 20);
            this.отчётыToolStripMenuItem.Text = "Отчёты за месяц";
            // 
            // MonthlyReportGeneralButton
            // 
            this.MonthlyReportGeneralButton.Name = "MonthlyReportGeneralButton";
            this.MonthlyReportGeneralButton.Size = new System.Drawing.Size(218, 22);
            this.MonthlyReportGeneralButton.Text = "Отчёт за месяц (общий)";
            this.MonthlyReportGeneralButton.Click += new System.EventHandler(this.MonthlyReportGeneralButton_Click);
            // 
            // MonthlyReportKitchenButton
            // 
            this.MonthlyReportKitchenButton.Name = "MonthlyReportKitchenButton";
            this.MonthlyReportKitchenButton.Size = new System.Drawing.Size(218, 22);
            this.MonthlyReportKitchenButton.Text = "Отчёт за месяц (столовая)";
            this.MonthlyReportKitchenButton.Click += new System.EventHandler(this.MonthlyReportKitchenButton_Click);
            // 
            // MonthlyReportShopButton
            // 
            this.MonthlyReportShopButton.Name = "MonthlyReportShopButton";
            this.MonthlyReportShopButton.Size = new System.Drawing.Size(218, 22);
            this.MonthlyReportShopButton.Text = "Отчёт за месяц  (магазин)";
            this.MonthlyReportShopButton.Click += new System.EventHandler(this.MonthlyReportShopButton_Click);
            // 
            // отчётыПоДнямToolStripMenuItem
            // 
            this.отчётыПоДнямToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ReportForTodayGeneralButton,
            this.ReportForTodayKitchenButton,
            this.ReportForTodayShopButton});
            this.отчётыПоДнямToolStripMenuItem.Name = "отчётыПоДнямToolStripMenuItem";
            this.отчётыПоДнямToolStripMenuItem.Size = new System.Drawing.Size(108, 20);
            this.отчётыПоДнямToolStripMenuItem.Text = "Отчёты по дням";
            // 
            // ReportForTodayGeneralButton
            // 
            this.ReportForTodayGeneralButton.Name = "ReportForTodayGeneralButton";
            this.ReportForTodayGeneralButton.Size = new System.Drawing.Size(227, 22);
            this.ReportForTodayGeneralButton.Text = "Отчёт за сегодня (общий)";
            this.ReportForTodayGeneralButton.Click += new System.EventHandler(this.ReportForTodayGeneralButton_Click);
            // 
            // ReportForTodayKitchenButton
            // 
            this.ReportForTodayKitchenButton.Name = "ReportForTodayKitchenButton";
            this.ReportForTodayKitchenButton.Size = new System.Drawing.Size(227, 22);
            this.ReportForTodayKitchenButton.Text = "Отчёт за сегодня (столовая)";
            this.ReportForTodayKitchenButton.Click += new System.EventHandler(this.ReportForTodayKitchenButton_Click);
            // 
            // ReportForTodayShopButton
            // 
            this.ReportForTodayShopButton.Name = "ReportForTodayShopButton";
            this.ReportForTodayShopButton.Size = new System.Drawing.Size(227, 22);
            this.ReportForTodayShopButton.Text = "Отчёт за сегодня (магазин)";
            this.ReportForTodayShopButton.Click += new System.EventHandler(this.ReportForTodayShopButton_Click);
            // 
            // CloseAllTabsButton
            // 
            this.CloseAllTabsButton.Name = "CloseAllTabsButton";
            this.CloseAllTabsButton.Size = new System.Drawing.Size(112, 20);
            this.CloseAllTabsButton.Text = "Закрыть вкладки";
            this.CloseAllTabsButton.Click += new System.EventHandler(this.CloseAllTabsButton_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.SingleHorizontal;
            dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle7.Padding = new System.Windows.Forms.Padding(18, 0, 0, 0);
            dataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle7;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView1.DefaultCellStyle = dataGridViewCellStyle8;
            this.dataGridView1.Location = new System.Drawing.Point(12, 32);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.Size = new System.Drawing.Size(957, 321);
            this.dataGridView1.TabIndex = 0;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.radioButton2);
            this.panel1.Controls.Add(this.radioButton1);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.ViewingDataInATableForDateButton);
            this.panel1.Controls.Add(this.CancelViewingDataForDateButton);
            this.panel1.Controls.Add(this.ReportDataInExcelForASpecificDateButton);
            this.panel1.Controls.Add(this.dateTimePicker1);
            this.panel1.Location = new System.Drawing.Point(96, 151);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(25, 13);
            this.panel1.TabIndex = 9;
            this.panel1.Visible = false;
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.radioButton2.Location = new System.Drawing.Point(179, 34);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(154, 26);
            this.radioButton2.TabIndex = 12;
            this.radioButton2.TabStop = true;
            this.radioButton2.Text = "521 (столовая)";
            this.radioButton2.UseVisualStyleBackColor = true;
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.radioButton1.Location = new System.Drawing.Point(19, 34);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(147, 26);
            this.radioButton1.TabIndex = 11;
            this.radioButton1.TabStop = true;
            this.radioButton1.Text = "518 (магазин)";
            this.radioButton1.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(16, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(136, 22);
            this.label1.TabIndex = 9;
            this.label1.Text = "Выберите код:";
            // 
            // ViewingDataInATableForDateButton
            // 
            this.ViewingDataInATableForDateButton.Location = new System.Drawing.Point(19, 95);
            this.ViewingDataInATableForDateButton.Name = "ViewingDataInATableForDateButton";
            this.ViewingDataInATableForDateButton.Size = new System.Drawing.Size(306, 35);
            this.ViewingDataInATableForDateButton.TabIndex = 8;
            this.ViewingDataInATableForDateButton.Text = "Просмотр данных";
            this.ViewingDataInATableForDateButton.UseVisualStyleBackColor = true;
            this.ViewingDataInATableForDateButton.Click += new System.EventHandler(this.ViewingDataInATableForDateButton_Click);
            // 
            // CancelViewingDataForDateButton
            // 
            this.CancelViewingDataForDateButton.Location = new System.Drawing.Point(20, 133);
            this.CancelViewingDataForDateButton.Name = "CancelViewingDataForDateButton";
            this.CancelViewingDataForDateButton.Size = new System.Drawing.Size(306, 35);
            this.CancelViewingDataForDateButton.TabIndex = 7;
            this.CancelViewingDataForDateButton.Text = "Отмена";
            this.CancelViewingDataForDateButton.UseVisualStyleBackColor = true;
            this.CancelViewingDataForDateButton.Click += new System.EventHandler(this.CancelViewingDataForDateButton_Click);
            // 
            // ReportDataInExcelForASpecificDateButton
            // 
            this.ReportDataInExcelForASpecificDateButton.Location = new System.Drawing.Point(332, 133);
            this.ReportDataInExcelForASpecificDateButton.Name = "ReportDataInExcelForASpecificDateButton";
            this.ReportDataInExcelForASpecificDateButton.Size = new System.Drawing.Size(207, 35);
            this.ReportDataInExcelForASpecificDateButton.TabIndex = 6;
            this.ReportDataInExcelForASpecificDateButton.Text = "Печать в Excel";
            this.ReportDataInExcelForASpecificDateButton.UseVisualStyleBackColor = true;
            this.ReportDataInExcelForASpecificDateButton.Click += new System.EventHandler(this.ReportDataInExcelForASpecificDateButton_Click);
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.CustomFormat = "";
            this.dateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePicker1.Location = new System.Drawing.Point(19, 65);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.ShowUpDown = true;
            this.dateTimePicker1.Size = new System.Drawing.Size(306, 20);
            this.dateTimePicker1.TabIndex = 2;
            // 
            // panel4
            // 
            this.panel4.Controls.Add(this.dateTimePicker2);
            this.panel4.Controls.Add(this.label11);
            this.panel4.Controls.Add(this.label10);
            this.panel4.Controls.Add(this.maskedTextBox1);
            this.panel4.Controls.Add(this.button10);
            this.panel4.Controls.Add(this.button11);
            this.panel4.Controls.Add(this.button12);
            this.panel4.Location = new System.Drawing.Point(104, 125);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(36, 20);
            this.panel4.TabIndex = 11;
            this.panel4.Visible = false;
            // 
            // dateTimePicker2
            // 
            this.dateTimePicker2.CustomFormat = "";
            this.dateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePicker2.Location = new System.Drawing.Point(19, 25);
            this.dateTimePicker2.Name = "dateTimePicker2";
            this.dateTimePicker2.ShowUpDown = true;
            this.dateTimePicker2.Size = new System.Drawing.Size(306, 20);
            this.dateTimePicker2.TabIndex = 18;
            // 
            // label11
            // 
            this.label11.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label11.Location = new System.Drawing.Point(16, 3);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(190, 22);
            this.label11.TabIndex = 17;
            this.label11.Text = "Выберите дату:";
            // 
            // label10
            // 
            this.label10.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label10.Location = new System.Drawing.Point(16, 42);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(190, 22);
            this.label10.TabIndex = 16;
            this.label10.Text = "Введите табельный:";
            // 
            // maskedTextBox1
            // 
            this.maskedTextBox1.Location = new System.Drawing.Point(17, 67);
            this.maskedTextBox1.Mask = "00000";
            this.maskedTextBox1.Name = "maskedTextBox1";
            this.maskedTextBox1.Size = new System.Drawing.Size(306, 20);
            this.maskedTextBox1.TabIndex = 8;
            // 
            // button10
            // 
            this.button10.Location = new System.Drawing.Point(19, 134);
            this.button10.Name = "button10";
            this.button10.Size = new System.Drawing.Size(306, 35);
            this.button10.TabIndex = 7;
            this.button10.Text = "Отмена";
            this.button10.UseVisualStyleBackColor = true;
            this.button10.Click += new System.EventHandler(this.button10_Click);
            // 
            // button11
            // 
            this.button11.Location = new System.Drawing.Point(331, 134);
            this.button11.Name = "button11";
            this.button11.Size = new System.Drawing.Size(197, 35);
            this.button11.TabIndex = 6;
            this.button11.Text = "Печать в Excel";
            this.button11.UseVisualStyleBackColor = true;
            this.button11.Click += new System.EventHandler(this.button11_Click);
            // 
            // button12
            // 
            this.button12.Location = new System.Drawing.Point(19, 93);
            this.button12.Name = "button12";
            this.button12.Size = new System.Drawing.Size(306, 35);
            this.button12.TabIndex = 4;
            this.button12.Text = "Просмотр данных";
            this.button12.UseVisualStyleBackColor = true;
            this.button12.Click += new System.EventHandler(this.button12_Click);
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.radioButton3);
            this.panel3.Controls.Add(this.radioButton4);
            this.panel3.Controls.Add(this.label2);
            this.panel3.Controls.Add(this.button5);
            this.panel3.Controls.Add(this.button6);
            this.panel3.Controls.Add(this.button7);
            this.panel3.Controls.Add(this.dateTimePicker4);
            this.panel3.Location = new System.Drawing.Point(52, 132);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(26, 13);
            this.panel3.TabIndex = 10;
            this.panel3.Visible = false;
            // 
            // radioButton3
            // 
            this.radioButton3.AutoSize = true;
            this.radioButton3.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.radioButton3.Location = new System.Drawing.Point(178, 33);
            this.radioButton3.Name = "radioButton3";
            this.radioButton3.Size = new System.Drawing.Size(154, 26);
            this.radioButton3.TabIndex = 15;
            this.radioButton3.TabStop = true;
            this.radioButton3.Text = "521 (столовая)";
            this.radioButton3.UseVisualStyleBackColor = true;
            // 
            // radioButton4
            // 
            this.radioButton4.AutoSize = true;
            this.radioButton4.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.radioButton4.Location = new System.Drawing.Point(18, 33);
            this.radioButton4.Name = "radioButton4";
            this.radioButton4.Size = new System.Drawing.Size(147, 26);
            this.radioButton4.TabIndex = 14;
            this.radioButton4.TabStop = true;
            this.radioButton4.Text = "518 (магазин)";
            this.radioButton4.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(15, 8);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(136, 22);
            this.label2.TabIndex = 13;
            this.label2.Text = "Выберите код:";
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(19, 135);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(306, 35);
            this.button5.TabIndex = 7;
            this.button5.Text = "Отмена";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(331, 135);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(196, 35);
            this.button6.TabIndex = 6;
            this.button6.Text = "Печать в Excel";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // button7
            // 
            this.button7.Location = new System.Drawing.Point(19, 94);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(306, 35);
            this.button7.TabIndex = 4;
            this.button7.Text = "Просмотр данных";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // dateTimePicker4
            // 
            this.dateTimePicker4.CustomFormat = "MM.yyyy";
            this.dateTimePicker4.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePicker4.Location = new System.Drawing.Point(19, 64);
            this.dateTimePicker4.Name = "dateTimePicker4";
            this.dateTimePicker4.ShowUpDown = true;
            this.dateTimePicker4.Size = new System.Drawing.Size(306, 20);
            this.dateTimePicker4.TabIndex = 2;
            // 
            // label3
            // 
            this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label3.Location = new System.Drawing.Point(465, 3);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(62, 22);
            this.label3.TabIndex = 14;
            this.label3.Text = "Итого";
            // 
            // label4
            // 
            this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label4.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label4.Location = new System.Drawing.Point(365, 26);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(128, 22);
            this.label4.TabIndex = 15;
            this.label4.Text = "Собственная:";
            // 
            // label5
            // 
            this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label5.Location = new System.Drawing.Point(410, 49);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(84, 22);
            this.label5.TabIndex = 16;
            this.label5.Text = "Готовая:";
            // 
            // label8
            // 
            this.label8.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label8.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label8.Location = new System.Drawing.Point(491, 27);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(143, 21);
            this.label8.TabIndex = 18;
            // 
            // label9
            // 
            this.label9.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label9.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label9.Location = new System.Drawing.Point(491, 50);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(145, 21);
            this.label9.TabIndex = 19;
            // 
            // panel2
            // 
            this.panel2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.panel2.Controls.Add(this.label13);
            this.panel2.Controls.Add(this.label12);
            this.panel2.Controls.Add(this.label7);
            this.panel2.Controls.Add(this.label6);
            this.panel2.Controls.Add(this.label9);
            this.panel2.Controls.Add(this.label8);
            this.panel2.Controls.Add(this.label5);
            this.panel2.Controls.Add(this.label4);
            this.panel2.Controls.Add(this.label3);
            this.panel2.Controls.Add(this.panel3);
            this.panel2.Controls.Add(this.panel4);
            this.panel2.Controls.Add(this.panel1);
            this.panel2.Location = new System.Drawing.Point(12, 359);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(641, 175);
            this.panel2.TabIndex = 10;
            // 
            // label13
            // 
            this.label13.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label13.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label13.Location = new System.Drawing.Point(493, 75);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(145, 22);
            this.label13.TabIndex = 23;
            // 
            // label12
            // 
            this.label12.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label12.Location = new System.Drawing.Point(369, 73);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(125, 22);
            this.label12.TabIndex = 22;
            this.label12.Text = "По магазину:";
            // 
            // label7
            // 
            this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label7.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label7.Location = new System.Drawing.Point(491, 98);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(145, 22);
            this.label7.TabIndex = 20;
            // 
            // label6
            // 
            this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label6.Location = new System.Drawing.Point(419, 97);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(75, 22);
            this.label6.TabIndex = 21;
            this.label6.Text = "Общая:";
            // 
            // SelectDataForASpecificDateButton
            // 
            this.SelectDataForASpecificDateButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.SelectDataForASpecificDateButton.Location = new System.Drawing.Point(659, 421);
            this.SelectDataForASpecificDateButton.Name = "SelectDataForASpecificDateButton";
            this.SelectDataForASpecificDateButton.Size = new System.Drawing.Size(310, 35);
            this.SelectDataForASpecificDateButton.TabIndex = 5;
            this.SelectDataForASpecificDateButton.Text = "Выбрать данные за определённое число";
            this.SelectDataForASpecificDateButton.UseVisualStyleBackColor = true;
            this.SelectDataForASpecificDateButton.Click += new System.EventHandler(this.SelectDataForASpecificDateButton_Click);
            // 
            // button8
            // 
            this.button8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button8.Location = new System.Drawing.Point(660, 491);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(310, 35);
            this.button8.TabIndex = 11;
            this.button8.Text = "Выбрать данные за определённый месяц";
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Click += new System.EventHandler(this.button8_Click);
            // 
            // button9
            // 
            this.button9.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button9.Location = new System.Drawing.Point(659, 456);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(310, 35);
            this.button9.TabIndex = 12;
            this.button9.Text = "Выбрать данные по табельному";
            this.button9.UseVisualStyleBackColor = true;
            this.button9.Click += new System.EventHandler(this.button9_Click);
            // 
            // button13
            // 
            this.button13.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button13.Location = new System.Drawing.Point(659, 389);
            this.button13.Name = "button13";
            this.button13.Size = new System.Drawing.Size(310, 32);
            this.button13.TabIndex = 13;
            this.button13.Text = "Экспорт данных";
            this.button13.UseVisualStyleBackColor = true;
            this.button13.Click += new System.EventHandler(this.button13_Click);
            // 
            // progressExportToDbf
            // 
            this.progressExportToDbf.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.progressExportToDbf.Location = new System.Drawing.Point(659, 359);
            this.progressExportToDbf.Name = "progressExportToDbf";
            this.progressExportToDbf.Size = new System.Drawing.Size(311, 23);
            this.progressExportToDbf.TabIndex = 14;
            this.progressExportToDbf.Visible = false;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(977, 546);
            this.Controls.Add(this.progressExportToDbf);
            this.Controls.Add(this.button13);
            this.Controls.Add(this.button9);
            this.Controls.Add(this.button8);
            this.Controls.Add(this.SelectDataForASpecificDateButton);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.menuStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Отчёты";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.DoubleClick += new System.EventHandler(this.Form1_DoubleClick);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel4.ResumeLayout(false);
            this.panel4.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem отчётыToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem MonthlyReportGeneralButton;
        private System.Windows.Forms.ToolStripMenuItem MonthlyReportKitchenButton;
        private System.Windows.Forms.ToolStripMenuItem MonthlyReportShopButton;
        private System.Windows.Forms.ToolStripMenuItem отчётыПоДнямToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem CloseAllTabsButton;
        private System.Windows.Forms.ToolStripMenuItem ReportForTodayGeneralButton;
        private System.Windows.Forms.ToolStripMenuItem ReportForTodayKitchenButton;
        private System.Windows.Forms.ToolStripMenuItem ReportForTodayShopButton;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Panel panel1;
        public System.Windows.Forms.RadioButton radioButton2;
        public System.Windows.Forms.RadioButton radioButton1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button ViewingDataInATableForDateButton;
        private System.Windows.Forms.Button CancelViewingDataForDateButton;
        private System.Windows.Forms.Button ReportDataInExcelForASpecificDateButton;
        public System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.DateTimePicker dateTimePicker2;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.MaskedTextBox maskedTextBox1;
        private System.Windows.Forms.Button button10;
        private System.Windows.Forms.Button button11;
        private System.Windows.Forms.Button button12;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.RadioButton radioButton3;
        private System.Windows.Forms.RadioButton radioButton4;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.DateTimePicker dateTimePicker4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button SelectDataForASpecificDateButton;
        private System.Windows.Forms.Button button8;
        private System.Windows.Forms.Button button9;
        private System.Windows.Forms.Button button13;
        private System.Windows.Forms.ProgressBar progressExportToDbf;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label12;
    }
}

