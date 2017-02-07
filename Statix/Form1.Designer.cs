namespace Statix
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            this.metroTabControl1 = new MetroFramework.Controls.MetroTabControl();
            this.metroTabPage1 = new MetroFramework.Controls.MetroTabPage();
            this.metroButton700 = new MetroFramework.Controls.MetroButton();
            this.metroGrid1 = new MetroFramework.Controls.MetroGrid();
            this.metroTabPage2 = new MetroFramework.Controls.MetroTabPage();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.metroButton8 = new MetroFramework.Controls.MetroButton();
            this.metroButton7 = new MetroFramework.Controls.MetroButton();
            this.metroButton6 = new MetroFramework.Controls.MetroButton();
            this.metroButton5 = new MetroFramework.Controls.MetroButton();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.metroButton10 = new MetroFramework.Controls.MetroButton();
            this.metroButton9 = new MetroFramework.Controls.MetroButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.metroButton4 = new MetroFramework.Controls.MetroButton();
            this.metroButton3 = new MetroFramework.Controls.MetroButton();
            this.metroButton2 = new MetroFramework.Controls.MetroButton();
            this.metroButton1 = new MetroFramework.Controls.MetroButton();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.metroTabPage3 = new MetroFramework.Controls.MetroTabPage();
            this.metroButton16 = new MetroFramework.Controls.MetroButton();
            this.metroButton15 = new MetroFramework.Controls.MetroButton();
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.metroButton14 = new MetroFramework.Controls.MetroButton();
            this.metroButton13 = new MetroFramework.Controls.MetroButton();
            this.metroButton12 = new MetroFramework.Controls.MetroButton();
            this.metroButton11 = new MetroFramework.Controls.MetroButton();
            this.groupBox9 = new System.Windows.Forms.GroupBox();
            this.groupBox8 = new System.Windows.Forms.GroupBox();
            this.metroTabPage4 = new MetroFramework.Controls.MetroTabPage();
            this.metroButton19 = new MetroFramework.Controls.MetroButton();
            this.metroButton20 = new MetroFramework.Controls.MetroButton();
            this.groupBox19 = new System.Windows.Forms.GroupBox();
            this.groupBox20 = new System.Windows.Forms.GroupBox();
            this.metroTabControl1.SuspendLayout();
            this.metroTabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.metroGrid1)).BeginInit();
            this.metroTabPage2.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.metroTabPage3.SuspendLayout();
            this.groupBox7.SuspendLayout();
            this.metroTabPage4.SuspendLayout();
            this.SuspendLayout();
            // 
            // metroTabControl1
            // 
            this.metroTabControl1.Controls.Add(this.metroTabPage1);
            this.metroTabControl1.Controls.Add(this.metroTabPage2);
            this.metroTabControl1.Controls.Add(this.metroTabPage3);
            this.metroTabControl1.Controls.Add(this.metroTabPage4);
            this.metroTabControl1.Location = new System.Drawing.Point(23, 63);
            this.metroTabControl1.Name = "metroTabControl1";
            this.metroTabControl1.SelectedIndex = 2;
            this.metroTabControl1.Size = new System.Drawing.Size(978, 529);
            this.metroTabControl1.TabIndex = 0;
            this.metroTabControl1.UseSelectable = true;
            this.metroTabControl1.Selected += new System.Windows.Forms.TabControlEventHandler(this.metroTabControl1_Selected);
            // 
            // metroTabPage1
            // 
            this.metroTabPage1.Controls.Add(this.metroButton700);
            this.metroTabPage1.Controls.Add(this.metroGrid1);
            this.metroTabPage1.HorizontalScrollbarBarColor = true;
            this.metroTabPage1.HorizontalScrollbarHighlightOnWheel = false;
            this.metroTabPage1.HorizontalScrollbarSize = 10;
            this.metroTabPage1.Location = new System.Drawing.Point(4, 35);
            this.metroTabPage1.Name = "metroTabPage1";
            this.metroTabPage1.Size = new System.Drawing.Size(970, 490);
            this.metroTabPage1.TabIndex = 0;
            this.metroTabPage1.Text = "Считанные данные";
            this.metroTabPage1.VerticalScrollbarBarColor = true;
            this.metroTabPage1.VerticalScrollbarHighlightOnWheel = false;
            this.metroTabPage1.VerticalScrollbarSize = 10;
            // 
            // metroButton700
            // 
            this.metroButton700.Location = new System.Drawing.Point(422, 446);
            this.metroButton700.Name = "metroButton700";
            this.metroButton700.Size = new System.Drawing.Size(144, 23);
            this.metroButton700.TabIndex = 3;
            this.metroButton700.Text = "Загрузка данных";
            this.metroButton700.UseSelectable = true;
            this.metroButton700.Click += new System.EventHandler(this.metroButton700_Click);
            // 
            // metroGrid1
            // 
            this.metroGrid1.AllowUserToResizeRows = false;
            this.metroGrid1.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.metroGrid1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.metroGrid1.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.metroGrid1.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(174)))), ((int)(((byte)(219)))));
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Segoe UI", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            dataGridViewCellStyle4.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(198)))), ((int)(((byte)(247)))));
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(17)))), ((int)(((byte)(17)))), ((int)(((byte)(17)))));
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.metroGrid1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle4;
            this.metroGrid1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            dataGridViewCellStyle5.Font = new System.Drawing.Font("Segoe UI", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            dataGridViewCellStyle5.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(136)))), ((int)(((byte)(136)))), ((int)(((byte)(136)))));
            dataGridViewCellStyle5.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(198)))), ((int)(((byte)(247)))));
            dataGridViewCellStyle5.SelectionForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(17)))), ((int)(((byte)(17)))), ((int)(((byte)(17)))));
            dataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.metroGrid1.DefaultCellStyle = dataGridViewCellStyle5;
            this.metroGrid1.EnableHeadersVisualStyles = false;
            this.metroGrid1.Font = new System.Drawing.Font("Segoe UI", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.metroGrid1.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.metroGrid1.Location = new System.Drawing.Point(3, 3);
            this.metroGrid1.Name = "metroGrid1";
            this.metroGrid1.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle6.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(174)))), ((int)(((byte)(219)))));
            dataGridViewCellStyle6.Font = new System.Drawing.Font("Segoe UI", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            dataGridViewCellStyle6.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            dataGridViewCellStyle6.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(198)))), ((int)(((byte)(247)))));
            dataGridViewCellStyle6.SelectionForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(17)))), ((int)(((byte)(17)))), ((int)(((byte)(17)))));
            dataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.metroGrid1.RowHeadersDefaultCellStyle = dataGridViewCellStyle6;
            this.metroGrid1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.metroGrid1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.metroGrid1.Size = new System.Drawing.Size(964, 411);
            this.metroGrid1.TabIndex = 2;
            // 
            // metroTabPage2
            // 
            this.metroTabPage2.Controls.Add(this.groupBox4);
            this.metroTabPage2.Controls.Add(this.metroButton10);
            this.metroTabPage2.Controls.Add(this.metroButton9);
            this.metroTabPage2.Controls.Add(this.groupBox1);
            this.metroTabPage2.HorizontalScrollbarBarColor = true;
            this.metroTabPage2.HorizontalScrollbarHighlightOnWheel = false;
            this.metroTabPage2.HorizontalScrollbarSize = 10;
            this.metroTabPage2.Location = new System.Drawing.Point(4, 35);
            this.metroTabPage2.Name = "metroTabPage2";
            this.metroTabPage2.Size = new System.Drawing.Size(970, 490);
            this.metroTabPage2.TabIndex = 1;
            this.metroTabPage2.Text = "Сравнение независимых групп";
            this.metroTabPage2.VerticalScrollbarBarColor = true;
            this.metroTabPage2.VerticalScrollbarHighlightOnWheel = false;
            this.metroTabPage2.VerticalScrollbarSize = 10;
            // 
            // groupBox4
            // 
            this.groupBox4.AutoSize = true;
            this.groupBox4.BackColor = System.Drawing.SystemColors.HighlightText;
            this.groupBox4.Controls.Add(this.metroButton8);
            this.groupBox4.Controls.Add(this.metroButton7);
            this.groupBox4.Controls.Add(this.metroButton6);
            this.groupBox4.Controls.Add(this.metroButton5);
            this.groupBox4.Controls.Add(this.groupBox6);
            this.groupBox4.Controls.Add(this.groupBox5);
            this.groupBox4.Location = new System.Drawing.Point(487, 3);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(480, 262);
            this.groupBox4.TabIndex = 8;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Выберите признаки";
            this.groupBox4.Visible = false;
            // 
            // metroButton8
            // 
            this.metroButton8.Location = new System.Drawing.Point(369, 220);
            this.metroButton8.Name = "metroButton8";
            this.metroButton8.Size = new System.Drawing.Size(75, 23);
            this.metroButton8.TabIndex = 5;
            this.metroButton8.Tag = "8";
            this.metroButton8.Text = "Очистить";
            this.metroButton8.UseSelectable = true;
            this.metroButton8.Click += new System.EventHandler(this.CleanAll_Click);
            // 
            // metroButton7
            // 
            this.metroButton7.Location = new System.Drawing.Point(274, 220);
            this.metroButton7.Name = "metroButton7";
            this.metroButton7.Size = new System.Drawing.Size(75, 23);
            this.metroButton7.TabIndex = 4;
            this.metroButton7.Tag = "7";
            this.metroButton7.Text = "Выбрать все";
            this.metroButton7.UseSelectable = true;
            this.metroButton7.Click += new System.EventHandler(this.SelectAll_Click);
            // 
            // metroButton6
            // 
            this.metroButton6.Location = new System.Drawing.Point(131, 220);
            this.metroButton6.Name = "metroButton6";
            this.metroButton6.Size = new System.Drawing.Size(75, 23);
            this.metroButton6.TabIndex = 3;
            this.metroButton6.Tag = "6";
            this.metroButton6.Text = "Очистить";
            this.metroButton6.UseSelectable = true;
            this.metroButton6.Click += new System.EventHandler(this.CleanAll_Click);
            // 
            // metroButton5
            // 
            this.metroButton5.Location = new System.Drawing.Point(36, 220);
            this.metroButton5.Name = "metroButton5";
            this.metroButton5.Size = new System.Drawing.Size(75, 23);
            this.metroButton5.TabIndex = 2;
            this.metroButton5.Tag = "5";
            this.metroButton5.Text = "Выбрать все";
            this.metroButton5.UseSelectable = true;
            this.metroButton5.Click += new System.EventHandler(this.SelectAll_Click);
            // 
            // groupBox6
            // 
            this.groupBox6.AutoSize = true;
            this.groupBox6.Location = new System.Drawing.Point(244, 19);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(230, 195);
            this.groupBox6.TabIndex = 1;
            this.groupBox6.TabStop = false;
            this.groupBox6.Tag = "1";
            this.groupBox6.Text = "Порядковые переменные";
            // 
            // groupBox5
            // 
            this.groupBox5.AutoSize = true;
            this.groupBox5.Location = new System.Drawing.Point(6, 19);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(230, 195);
            this.groupBox5.TabIndex = 0;
            this.groupBox5.TabStop = false;
            this.groupBox5.Tag = "1";
            this.groupBox5.Text = "Количественные переменные";
            // 
            // metroButton10
            // 
            this.metroButton10.Location = new System.Drawing.Point(487, 441);
            this.metroButton10.Name = "metroButton10";
            this.metroButton10.Size = new System.Drawing.Size(159, 43);
            this.metroButton10.TabIndex = 6;
            this.metroButton10.Text = "Вывести результат в Word";
            this.metroButton10.UseSelectable = true;
            this.metroButton10.Visible = false;
            this.metroButton10.Click += new System.EventHandler(this.metroButton10_Click);
            // 
            // metroButton9
            // 
            this.metroButton9.Location = new System.Drawing.Point(324, 441);
            this.metroButton9.Name = "metroButton9";
            this.metroButton9.Size = new System.Drawing.Size(159, 43);
            this.metroButton9.TabIndex = 5;
            this.metroButton9.Text = "Выполнить сравнение";
            this.metroButton9.UseSelectable = true;
            this.metroButton9.Visible = false;
            this.metroButton9.Click += new System.EventHandler(this.metroButton9_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.AutoSize = true;
            this.groupBox1.BackColor = System.Drawing.SystemColors.HighlightText;
            this.groupBox1.Controls.Add(this.metroButton4);
            this.groupBox1.Controls.Add(this.metroButton3);
            this.groupBox1.Controls.Add(this.metroButton2);
            this.groupBox1.Controls.Add(this.metroButton1);
            this.groupBox1.Controls.Add(this.groupBox3);
            this.groupBox1.Controls.Add(this.groupBox2);
            this.groupBox1.Location = new System.Drawing.Point(3, 3);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(480, 262);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Выберете группирующие факторы";
            // 
            // metroButton4
            // 
            this.metroButton4.Location = new System.Drawing.Point(369, 220);
            this.metroButton4.Name = "metroButton4";
            this.metroButton4.Size = new System.Drawing.Size(75, 23);
            this.metroButton4.TabIndex = 5;
            this.metroButton4.Tag = "4";
            this.metroButton4.Text = "Очистить";
            this.metroButton4.UseSelectable = true;
            this.metroButton4.Click += new System.EventHandler(this.CleanAll_Click);
            // 
            // metroButton3
            // 
            this.metroButton3.Location = new System.Drawing.Point(274, 220);
            this.metroButton3.Name = "metroButton3";
            this.metroButton3.Size = new System.Drawing.Size(75, 23);
            this.metroButton3.TabIndex = 4;
            this.metroButton3.Tag = "3";
            this.metroButton3.Text = "Выбрать все";
            this.metroButton3.UseSelectable = true;
            this.metroButton3.Click += new System.EventHandler(this.SelectAll_Click);
            // 
            // metroButton2
            // 
            this.metroButton2.Location = new System.Drawing.Point(131, 220);
            this.metroButton2.Name = "metroButton2";
            this.metroButton2.Size = new System.Drawing.Size(75, 23);
            this.metroButton2.TabIndex = 3;
            this.metroButton2.Tag = "2";
            this.metroButton2.Text = "Очистить";
            this.metroButton2.UseSelectable = true;
            this.metroButton2.Click += new System.EventHandler(this.CleanAll_Click);
            // 
            // metroButton1
            // 
            this.metroButton1.Location = new System.Drawing.Point(36, 220);
            this.metroButton1.Name = "metroButton1";
            this.metroButton1.Size = new System.Drawing.Size(75, 23);
            this.metroButton1.TabIndex = 2;
            this.metroButton1.Tag = "1";
            this.metroButton1.Text = "Выбрать все";
            this.metroButton1.UseSelectable = true;
            this.metroButton1.Click += new System.EventHandler(this.SelectAll_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.AutoSize = true;
            this.groupBox3.Location = new System.Drawing.Point(244, 19);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(230, 195);
            this.groupBox3.TabIndex = 1;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Номинальные переменные";
            // 
            // groupBox2
            // 
            this.groupBox2.AutoSize = true;
            this.groupBox2.Location = new System.Drawing.Point(6, 19);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(230, 195);
            this.groupBox2.TabIndex = 0;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Бинарные переменные";
            // 
            // metroTabPage3
            // 
            this.metroTabPage3.Controls.Add(this.metroButton16);
            this.metroTabPage3.Controls.Add(this.metroButton15);
            this.metroTabPage3.Controls.Add(this.groupBox7);
            this.metroTabPage3.HorizontalScrollbarBarColor = true;
            this.metroTabPage3.HorizontalScrollbarHighlightOnWheel = false;
            this.metroTabPage3.HorizontalScrollbarSize = 10;
            this.metroTabPage3.Location = new System.Drawing.Point(4, 38);
            this.metroTabPage3.Name = "metroTabPage3";
            this.metroTabPage3.Size = new System.Drawing.Size(970, 487);
            this.metroTabPage3.TabIndex = 2;
            this.metroTabPage3.Text = "Сравнение зависимых групп";
            this.metroTabPage3.VerticalScrollbarBarColor = true;
            this.metroTabPage3.VerticalScrollbarHighlightOnWheel = false;
            this.metroTabPage3.VerticalScrollbarSize = 10;
            // 
            // metroButton16
            // 
            this.metroButton16.Location = new System.Drawing.Point(487, 441);
            this.metroButton16.Name = "metroButton16";
            this.metroButton16.Size = new System.Drawing.Size(159, 43);
            this.metroButton16.TabIndex = 11;
            this.metroButton16.Text = "Вывести результат в Word";
            this.metroButton16.UseSelectable = true;
            this.metroButton16.Visible = false;
            this.metroButton16.Click += new System.EventHandler(this.metroButton16_Click);
            // 
            // metroButton15
            // 
            this.metroButton15.Location = new System.Drawing.Point(324, 441);
            this.metroButton15.Name = "metroButton15";
            this.metroButton15.Size = new System.Drawing.Size(159, 43);
            this.metroButton15.TabIndex = 10;
            this.metroButton15.Text = "Выполнить сравнение";
            this.metroButton15.UseSelectable = true;
            this.metroButton15.Visible = false;
            this.metroButton15.Click += new System.EventHandler(this.metroButton15_Click);
            // 
            // groupBox7
            // 
            this.groupBox7.AutoSize = true;
            this.groupBox7.BackColor = System.Drawing.SystemColors.HighlightText;
            this.groupBox7.Controls.Add(this.metroButton14);
            this.groupBox7.Controls.Add(this.metroButton13);
            this.groupBox7.Controls.Add(this.metroButton12);
            this.groupBox7.Controls.Add(this.metroButton11);
            this.groupBox7.Controls.Add(this.groupBox9);
            this.groupBox7.Controls.Add(this.groupBox8);
            this.groupBox7.Location = new System.Drawing.Point(3, 3);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Size = new System.Drawing.Size(480, 262);
            this.groupBox7.TabIndex = 9;
            this.groupBox7.TabStop = false;
            this.groupBox7.Text = "Выберите признаки";
            // 
            // metroButton14
            // 
            this.metroButton14.Location = new System.Drawing.Point(369, 220);
            this.metroButton14.Name = "metroButton14";
            this.metroButton14.Size = new System.Drawing.Size(75, 23);
            this.metroButton14.TabIndex = 5;
            this.metroButton14.Tag = "12";
            this.metroButton14.Text = "Очистить";
            this.metroButton14.UseSelectable = true;
            this.metroButton14.Click += new System.EventHandler(this.CleanAll_Click);
            // 
            // metroButton13
            // 
            this.metroButton13.Location = new System.Drawing.Point(274, 220);
            this.metroButton13.Name = "metroButton13";
            this.metroButton13.Size = new System.Drawing.Size(75, 23);
            this.metroButton13.TabIndex = 4;
            this.metroButton13.Tag = "11";
            this.metroButton13.Text = "Выбрать все";
            this.metroButton13.UseSelectable = true;
            this.metroButton13.Click += new System.EventHandler(this.SelectAll_Click);
            // 
            // metroButton12
            // 
            this.metroButton12.Location = new System.Drawing.Point(131, 220);
            this.metroButton12.Name = "metroButton12";
            this.metroButton12.Size = new System.Drawing.Size(75, 23);
            this.metroButton12.TabIndex = 3;
            this.metroButton12.Tag = "10";
            this.metroButton12.Text = "Очистить";
            this.metroButton12.UseSelectable = true;
            this.metroButton12.Click += new System.EventHandler(this.CleanAll_Click);
            // 
            // metroButton11
            // 
            this.metroButton11.Location = new System.Drawing.Point(36, 220);
            this.metroButton11.Name = "metroButton11";
            this.metroButton11.Size = new System.Drawing.Size(75, 23);
            this.metroButton11.TabIndex = 2;
            this.metroButton11.Tag = "9";
            this.metroButton11.Text = "Выбрать все";
            this.metroButton11.UseSelectable = true;
            this.metroButton11.Click += new System.EventHandler(this.SelectAll_Click);
            // 
            // groupBox9
            // 
            this.groupBox9.AutoSize = true;
            this.groupBox9.Location = new System.Drawing.Point(244, 19);
            this.groupBox9.Name = "groupBox9";
            this.groupBox9.Size = new System.Drawing.Size(230, 195);
            this.groupBox9.TabIndex = 1;
            this.groupBox9.TabStop = false;
            this.groupBox9.Tag = "2";
            this.groupBox9.Text = "Порядковые переменные";
            // 
            // groupBox8
            // 
            this.groupBox8.AutoSize = true;
            this.groupBox8.Location = new System.Drawing.Point(6, 19);
            this.groupBox8.Name = "groupBox8";
            this.groupBox8.Size = new System.Drawing.Size(230, 195);
            this.groupBox8.TabIndex = 0;
            this.groupBox8.TabStop = false;
            this.groupBox8.Tag = "2";
            this.groupBox8.Text = "Количественные переменные";
            // 
            // metroTabPage4
            // 
            this.metroTabPage4.Controls.Add(this.metroButton19);
            this.metroTabPage4.Controls.Add(this.metroButton20);
            this.metroTabPage4.Controls.Add(this.groupBox19);
            this.metroTabPage4.Controls.Add(this.groupBox20);
            this.metroTabPage4.HorizontalScrollbarBarColor = true;
            this.metroTabPage4.HorizontalScrollbarHighlightOnWheel = false;
            this.metroTabPage4.HorizontalScrollbarSize = 10;
            this.metroTabPage4.Location = new System.Drawing.Point(4, 35);
            this.metroTabPage4.Name = "metroTabPage4";
            this.metroTabPage4.Size = new System.Drawing.Size(970, 490);
            this.metroTabPage4.TabIndex = 3;
            this.metroTabPage4.Text = "Корреляционный анализ";
            this.metroTabPage4.VerticalScrollbarBarColor = true;
            this.metroTabPage4.VerticalScrollbarHighlightOnWheel = false;
            this.metroTabPage4.VerticalScrollbarSize = 10;
            // 
            // metroButton19
            // 
            this.metroButton19.Location = new System.Drawing.Point(458, 128);
            this.metroButton19.Name = "metroButton19";
            this.metroButton19.Size = new System.Drawing.Size(159, 43);
            this.metroButton19.TabIndex = 10;
            this.metroButton19.Text = "Вывести результат в Word";
            this.metroButton19.UseSelectable = true;
            this.metroButton19.Visible = false;
            // 
            // metroButton20
            // 
            this.metroButton20.Location = new System.Drawing.Point(458, 60);
            this.metroButton20.Name = "metroButton20";
            this.metroButton20.Size = new System.Drawing.Size(159, 43);
            this.metroButton20.TabIndex = 9;
            this.metroButton20.Text = "Выполнить анализ";
            this.metroButton20.UseSelectable = true;
            this.metroButton20.Visible = false;
            // 
            // groupBox19
            // 
            this.groupBox19.AutoSize = true;
            this.groupBox19.Location = new System.Drawing.Point(3, 3);
            this.groupBox19.Name = "groupBox19";
            this.groupBox19.Size = new System.Drawing.Size(200, 235);
            this.groupBox19.TabIndex = 7;
            this.groupBox19.TabStop = false;
            this.groupBox19.Text = "Количественные переменные";
            // 
            // groupBox20
            // 
            this.groupBox20.AutoSize = true;
            this.groupBox20.Location = new System.Drawing.Point(209, 3);
            this.groupBox20.Name = "groupBox20";
            this.groupBox20.Size = new System.Drawing.Size(200, 235);
            this.groupBox20.TabIndex = 6;
            this.groupBox20.TabStop = false;
            this.groupBox20.Text = "Порядковые переменные";
            // 
            // Form1
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(1024, 600);
            this.Controls.Add(this.metroTabControl1);
            this.Name = "Form1";
            this.Resizable = false;
            this.Text = "StatiX";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.metroTabControl1.ResumeLayout(false);
            this.metroTabPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.metroGrid1)).EndInit();
            this.metroTabPage2.ResumeLayout(false);
            this.metroTabPage2.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.metroTabPage3.ResumeLayout(false);
            this.metroTabPage3.PerformLayout();
            this.groupBox7.ResumeLayout(false);
            this.groupBox7.PerformLayout();
            this.metroTabPage4.ResumeLayout(false);
            this.metroTabPage4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private MetroFramework.Controls.MetroTabControl metroTabControl1;
        private MetroFramework.Controls.MetroTabPage metroTabPage1;
        private MetroFramework.Controls.MetroTabPage metroTabPage2;
        private MetroFramework.Controls.MetroTabPage metroTabPage3;
        private MetroFramework.Controls.MetroTabPage metroTabPage4;
        private System.Windows.Forms.GroupBox groupBox1;
        private MetroFramework.Controls.MetroButton metroButton10;
        private MetroFramework.Controls.MetroButton metroButton9;
        private MetroFramework.Controls.MetroButton metroButton19;
        private MetroFramework.Controls.MetroButton metroButton20;
        private System.Windows.Forms.GroupBox groupBox19;
        private System.Windows.Forms.GroupBox groupBox20;
        private MetroFramework.Controls.MetroGrid metroGrid1;
        private MetroFramework.Controls.MetroButton metroButton700;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox groupBox2;
        private MetroFramework.Controls.MetroButton metroButton1;
        private MetroFramework.Controls.MetroButton metroButton2;
        private MetroFramework.Controls.MetroButton metroButton4;
        private MetroFramework.Controls.MetroButton metroButton3;
        private System.Windows.Forms.GroupBox groupBox4;
        private MetroFramework.Controls.MetroButton metroButton8;
        private MetroFramework.Controls.MetroButton metroButton7;
        private MetroFramework.Controls.MetroButton metroButton6;
        private MetroFramework.Controls.MetroButton metroButton5;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.GroupBox groupBox5;
        private MetroFramework.Controls.MetroButton metroButton16;
        private MetroFramework.Controls.MetroButton metroButton15;
        private System.Windows.Forms.GroupBox groupBox7;
        private MetroFramework.Controls.MetroButton metroButton14;
        private MetroFramework.Controls.MetroButton metroButton13;
        private MetroFramework.Controls.MetroButton metroButton12;
        private MetroFramework.Controls.MetroButton metroButton11;
        private System.Windows.Forms.GroupBox groupBox9;
        private System.Windows.Forms.GroupBox groupBox8;
    }
}

