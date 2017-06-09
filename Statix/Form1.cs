using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

//Пакет R
using RDotNet;

//Библиотека для работы с Word
using Gios.Word;

//Внешнее оформление
using MetroFramework;
using MetroFramework.Forms;
using MetroFramework.Controls;

namespace Statix
{
    public partial class Form1 : MetroForm
    {
        #region Поля
        
        /// <summary>
        /// Дирректории для C# и R
        /// </summary>
        private string pathC = "..\\..\\graphics";
        private string pathR = "../../graphics/";

        /// <summary>
        /// Пространство языка R
        /// </summary>
        private REngine engine;

        /// <summary>
        /// Ностройки приложения.
        /// </summary>
        Properties.Settings settings;

        /// <summary>
        /// Считанные данные
        /// </summary>
        private Data data;

        /// <summary>
        /// Получаем список индексов с Бин и Ном шкалами
        /// </summary>
        private List<int> binList;

        /// <summary>
        /// Получаем список индексов с Ном шкалами
        /// </summary>
        private List<int> nomList;

        /// <summary>
        /// Получаем список индексов с Кол шкалой
        /// </summary>
        private List<int> colList;

        /// <summary>
        /// Получаем список индексов с Пор шкалой
        /// </summary>
        private List<int> porList;

        /// <summary>
        /// Список индексов группирующих факторов
        /// </summary>
        private List<int> groupFactList = new List<int>();

        /// <summary>
        /// Список выбранных признаков
        /// </summary>
        private List<int> signsList = new List<int>();

        /// <summary>
        /// Сравнение независимых групп. Тест Манна-Уитни
        /// </summary>
        private List<Sample> resIndMannaWhitney = new List<Sample>();

        /// <summary>
        /// Сравнение независимых групп. Тест Краскела-Уоллиса
        /// </summary>
        private List<Sample> resIndKruskalWallis = new List<Sample>();

        /// <summary>
        /// Сравнение зависимых групп. Тест Вилкоксона
        /// </summary>
        private List<Sample> resDepWilcoxon = new List<Sample>();

        /// <summary>
        /// Сравнение зависимых групп. Тест Фридмана
        /// </summary>
        private List<Sample> resDepFridman = new List<Sample>();

        /// <summary>
        /// Структура для хранения результата корреляционного анализа
        /// </summary>
        private struct СorrelationResult
        {
            /// <summary>
            /// Коэффициент корреляции 
            /// </summary>
            public double r;

             /// <summary>
             /// p-value
             /// </summary>
            public double p;

            /// <summary>
            /// Индекс что сравниваем
            /// </summary>
            public int thatCompare;

            /// <summary>
            /// Индекс с чем сравниваем
            /// </summary>
            public int toCompare;
        }

        /// <summary>
        /// Корреляционный анализ. Метод Пирсона
        /// </summary>
        private List<СorrelationResult> resCorPearson = new List<СorrelationResult>();

        /// <summary>
        /// Корреляционный анализ. Метод Спирмена
        /// </summary>
        private List<СorrelationResult> resCorSpearman = new List<СorrelationResult>();

        /// <summary>
        /// Таблицы сопряженности. Список таблиц сопряженности
        /// </summary>
        private List<ContingencyTableResult> resContingencyTables = new List<ContingencyTableResult>();

        /// <summary>
        /// Структура для хранения результата сравнения таблиц методом сопряженных таблиц
        /// </summary>
        private struct ContingencyTableResult
        {
            /// <summary>
            /// p-значение
            /// </summary>
            public double pvalue;

            /// <summary>
            /// значение статистики
            /// </summary>
            public double stat;

            /// <summary>
            /// Таблица сопряженности
            /// </summary>
            public ContingencyTable table;
        }

        #endregion

        public Form1()
        {
            InitializeComponent();
            settings = new Properties.Settings();
        }

        /// <summary>
        /// Загрузка формы. Начальные настройки формы
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            REngine.SetEnvironmentVariables();
            metroTabControl1.SelectedIndex = 0;

            binList = new List<int>();
            nomList = new List<int>();
            colList = new List<int>();
            porList = new List<int>();
        }

        /// <summary>
        /// Считывание данных из файла. Вывод данных для проверки. Формирование списков по типу шкал.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void metroButton700_Click(object sender, EventArgs e)
        {
            //Считывание данных из файла
            //Далее будет выбор файла
            data = new Data("Exam.csv", Encoding.Default);

            //Вывод считанной информации для проверки
            OutRedingInformation();

            //Получение ошибок, найденных при считывании
            Errors errors = data.ErrorsList();

            //Вывод сообщения о найденных ошибках
            //if (errors.CountErrors != 0)
            //{
            //    string title = "Во время считывания данных из файла были найдены ошибки!";
            //    string message = "Желаете посмотреть список ошибок?\n" + "\"Yes\" - Да, \"No\" - Нет.\n";
            //    if (MetroMessageBox.Show(this, message, title, MessageBoxButtons.YesNo, MessageBoxIcon.Error) == DialogResult.Yes)
            //    {
            //        //Вывод списка ошибок в новом окне
            //        FormForErrors Form2 = new FormForErrors(data.ErrorsList().ErrorsList);
            //        Form2.Show();
            //    }
            //}

            //Формирование списков по типу шкал
            CreateListsByTypeScale();
        }

        /// <summary>
        /// Вывод считанных данных в таблицу для проверки
        /// </summary>
        private void OutRedingInformation()
        {
            int records = data.Records;
            metroGrid1.ColumnCount = records;
            metroGrid1.RowCount = data.PatientsCount + 3;

            //Строка "Уникальные значения"
            metroGrid1.Rows[0].DefaultCellStyle.BackColor = settings.ColorTableCaption;
            for (int i = 0; i < records; i++)
                metroGrid1.Rows[0].Cells[i].Value = data.TakeUniqueValuesAtIndex(i);

            //Строка "Тип шкалы"
            metroGrid1.Rows[1].DefaultCellStyle.BackColor = settings.ColorTableCaption;
            for (int i = 0; i < records; i++)
                metroGrid1.Rows[1].Cells[i].Value = data.TakeScaleTypeAtIndex(i);

            //Строка "Название переменной"
            metroGrid1.Rows[2].DefaultCellStyle.BackColor = settings.ColorTableCaption;
            for (int i = 0; i < records; i++)
                metroGrid1.Rows[2].Cells[i].Value = data.TakeVariableNameAtIndex(i);

            for (int j = 0; j < data.PatientsCount; j++)
            {
                List<string> person = data.TakePatientAtIndex(j);
                //Цвет строки в таблице
                if (person.IndexOf("NA") != -1)
                    metroGrid1.Rows[j + 3].DefaultCellStyle.BackColor = settings.ColorTableError;
                else
                    metroGrid1.Rows[j + 3].DefaultCellStyle.BackColor = settings.ColorTableOK;

                for (int i = 0; i < records; i++)
                    metroGrid1.Rows[j + 3].Cells[i].Value = person[i];
            }
        }

        /// <summary>
        /// Формирование списков по типу шкал
        /// </summary>
        private void CreateListsByTypeScale()
        {
            colList = new List<int>();
            porList = new List<int>();
            binList = new List<int>();
            nomList = new List<int>();

            //Получаем списки индексов шкал
            for (int i = 0; i < data.Records; i++)
            {
                string tp = data.TakeScaleTypeAtIndex(i);
                switch (tp)
                {
                    case "бин": binList.Add(i); break;
                    case "ном": nomList.Add(i); break;
                    case "пор": porList.Add(i); break;
                    default: colList.Add(i); break;
                }
            }
        }
        
        /// <summary>
        /// Обработчик выбранной вкладки
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void metroTabControl1_Selected(object sender, TabControlEventArgs e)
        {
            //При переходе между вкладками, данные очищаются
            groupFactList.Clear();
            signsList.Clear();
            
            #region "Сравнение независимых групп"
            if (e.TabPageIndex == 1)
            {
                //Очистим хранилища textBox'ов
                groupBox2.Controls.Clear();
                groupBox3.Controls.Clear();
                groupBox5.Controls.Clear();
                groupBox6.Controls.Clear();
                MetroCheckBox[] rB;
                
                //Вывод списка группирующих переменных
                //Вывод бинарных переменных
                rB = new MetroCheckBox[binList.Count];
                for (int i = 0; i < binList.Count; i++)
                {
                    rB[i] = new MetroCheckBox();
                    rB[i].Text = data.TakeVariableNameAtIndex(binList[i]);
                    rB[i].Checked = false;
                    rB[i].Tag = binList[i];
                    rB[i].AutoSize = true;
                    rB[i].Location = new Point(settings.StandartPadding * 2, settings.PaddingBetweenCheckBoxes * (i + 1));
                    rB[i].Padding = new Padding(settings.StandartPadding * 2);
                    rB[i].Margin = new Padding(settings.StandartPadding * 2);
                    rB[i].CheckedChanged += CheckedChangedForGroupFac;
                    groupBox2.Controls.Add(rB[i]);
                }
                
                //Вывод номинальных переменных
                rB = new MetroCheckBox[nomList.Count];
                for (int i = 0; i < nomList.Count; i++)
                {
                    rB[i] = new MetroCheckBox();
                    rB[i].Text = data.TakeVariableNameAtIndex(nomList[i]);
                    rB[i].Checked = false;
                    rB[i].Tag = nomList[i];
                    rB[i].AutoSize = true;
                    rB[i].Location = new Point(settings.StandartPadding * 2, settings.PaddingBetweenCheckBoxes * (i + 1));
                    rB[i].Padding = new Padding(settings.StandartPadding * 2);
                    rB[i].Margin = new Padding(settings.StandartPadding * 2);
                    rB[i].CheckedChanged += CheckedChangedForGroupFac;
                    groupBox3.Controls.Add(rB[i]);
                }

                //Вывод признаков
                //Вывод списка количественных переменных
                rB = new MetroCheckBox[colList.Count];
                for (int i = 0; i < colList.Count; i++)
                {
                    rB[i] = new MetroCheckBox();
                    rB[i].Text = data.TakeVariableNameAtIndex(colList[i]);
                    rB[i].Checked = false;
                    rB[i].Tag = colList[i];
                    rB[i].AutoSize = true;
                    rB[i].Location = new Point(settings.StandartPadding * 2, settings.PaddingBetweenCheckBoxes * (i + 1));
                    rB[i].Padding = new Padding(settings.StandartPadding * 2);
                    rB[i].Margin = new Padding(settings.StandartPadding * 2);
                    rB[i].CheckedChanged += CheckedChangedForSigns;
                    //rB[i].Size = new Size(220, 17);
                    groupBox5.Controls.Add(rB[i]);
                }
                
                //Вывод списка порядковых переменных
                rB = new MetroCheckBox[porList.Count];
                for (int i = 0; i < porList.Count; i++)
                {
                    rB[i] = new MetroCheckBox();
                    rB[i].Text = data.TakeVariableNameAtIndex(porList[i]);
                    rB[i].Checked = false;
                    rB[i].Tag = porList[i];
                    rB[i].AutoSize = true;
                    rB[i].Location = new Point(settings.StandartPadding * 2, settings.PaddingBetweenCheckBoxes * (i + 1));
                    rB[i].Padding = new Padding(settings.StandartPadding * 2);
                    rB[i].Margin = new Padding(settings.StandartPadding * 2);
                    rB[i].CheckedChanged += CheckedChangedForSigns;
                    groupBox6.Controls.Add(rB[i]);
                }
                
                //Скроем кнопки с других вкладок
                //Кнопки выполнения анализа
                metroButton104.Visible = false;
                metroButton204.Visible = false;
                metroButton304.Visible = false;

                //Кнопки вывод в Word
                metroButton10.Visible = false;
                metroButton105.Visible = false;
                metroButton205.Visible = false;
                metroButton305.Visible = false;

                //Скроем вкладку с признаками. Это для корректного отображения при возврате на данную вкладку
                groupBox4.Visible = false;

                //Настроим высоту groupBox'ов, которые хранят в себе CheckBox'ы
                int height = groupBox2.Size.Height;
                int width = groupBox2.Size.Width;
                if (groupBox2.Controls.Count * settings.PaddingBetweenCheckBoxes > height)
                    height = groupBox2.Controls.Count * settings.PaddingBetweenCheckBoxes;
                groupBox2.Size = new Size(width, height);

                height = groupBox3.Size.Height;
                width = groupBox3.Size.Width;
                if (groupBox3.Controls.Count * settings.PaddingBetweenCheckBoxes > height)
                    height = groupBox3.Controls.Count * settings.PaddingBetweenCheckBoxes;
                groupBox3.Size = new Size(width, height);

                height = groupBox5.Size.Height;
                width = groupBox5.Size.Width;
                if (groupBox5.Controls.Count * settings.PaddingBetweenCheckBoxes > height)
                    height = groupBox5.Controls.Count * settings.PaddingBetweenCheckBoxes;
                groupBox5.Size = new Size(width, height);

                height = groupBox6.Size.Height;
                width = groupBox6.Size.Width;
                if (groupBox6.Controls.Count * settings.PaddingBetweenCheckBoxes > height)
                    height = groupBox6.Controls.Count * settings.PaddingBetweenCheckBoxes;
                groupBox6.Size = new Size(width, height);

                //Выровним groupBox'ы для Груп. факторов и признаков
                //64 - для отображения кнопок под grpBox. Иначе их не видно! ПОФИКСИТЬ!
                if (groupBox2.Size.Height > groupBox3.Size.Height)
                    groupBox1.Size = new Size(groupBox1.Size.Width, groupBox2.Size.Height + 64);
                if (groupBox5.Size.Height > groupBox6.Size.Height)
                    groupBox4.Size = new Size(groupBox4.Size.Width, groupBox5.Size.Height + 64);

                if (groupBox1.Height > groupBox4.Height)
                    groupBox4.Height = groupBox1.Height;
                else
                    groupBox1.Height = groupBox4.Height;

                //Настроим внешний вид страницы
                metroTabPage2_SizeChanged(sender, e);
            }
            #endregion

            #region "Сравнение зависимых групп"

            if (e.TabPageIndex == 2)
            {
                groupBox8.Controls.Clear();
                groupBox9.Controls.Clear();
                MetroCheckBox[] rB;

                //Скроем кнопки с других вкладок
                metroButton9.Visible = false;
                metroButton204.Visible = false;
                metroButton304.Visible = false;

                //Кнопки вывод в Word
                metroButton10.Visible = false;
                metroButton105.Visible = false;
                metroButton205.Visible = false;
                metroButton305.Visible = false;

                //Вывод признаков
                //Вывод списка количественных переменных
                rB = new MetroCheckBox[colList.Count];
                for (int i = 0; i < colList.Count; i++)
                {
                    rB[i] = new MetroCheckBox();
                    rB[i].Text = data.TakeVariableNameAtIndex(colList[i]);
                    rB[i].Checked = false;
                    rB[i].Tag = colList[i];
                    rB[i].AutoSize = true;
                    rB[i].Location = new Point(settings.StandartPadding * 2, settings.PaddingBetweenCheckBoxes * (i + 1));
                    rB[i].Padding = new Padding(settings.StandartPadding * 2);
                    rB[i].Margin = new Padding(settings.StandartPadding * 2);
                    rB[i].CheckedChanged += CheckedChangedForSigns;
                    groupBox8.Controls.Add(rB[i]);
                }

                //Изменение положения кнопок
                Point point = new Point(metroButton100.Location.X, metroButton100.Location.Y);
                point.Y = groupBox8.Size.Height + 25;
                metroButton100.Location = point;
                point = new Point(metroButton101.Location.X, metroButton101.Location.Y);
                point.Y = groupBox8.Size.Height + 25;
                metroButton101.Location = point;

                //Вывод списка порядковых переменных
                rB = new MetroCheckBox[porList.Count];
                for (int i = 0; i < porList.Count; i++)
                {
                    rB[i] = new MetroCheckBox();
                    rB[i].Text = data.TakeVariableNameAtIndex(porList[i]);
                    rB[i].Checked = false;
                    rB[i].Tag = porList[i];
                    rB[i].AutoSize = true;
                    rB[i].Location = new Point(settings.StandartPadding * 2, settings.PaddingBetweenCheckBoxes * (i + 1));
                    rB[i].Padding = new Padding(settings.StandartPadding * 2);
                    rB[i].Margin = new Padding(settings.StandartPadding * 2);
                    rB[i].CheckedChanged += CheckedChangedForSigns;
                    groupBox9.Controls.Add(rB[i]);
                }

                //Изменение положения кнопок
                point = new Point(metroButton102.Location.X, metroButton102.Location.Y);
                point.Y = groupBox9.Size.Height + 25;
                metroButton102.Location = point;
                point = new Point(metroButton103.Location.X, metroButton103.Location.Y);
                point.Y = groupBox9.Size.Height + 25;
                metroButton103.Location = point;

                //Выровняем кнопки "Выполнить сравнение" и "Вывести в Word"
                point = new Point(metroButton104.Location.X, groupBox7.Size.Height);
                point.Y += 20;
                metroButton104.Location = point;
                point = new Point(metroButton105.Location.X, groupBox7.Size.Height);
                point.Y += 20;
                metroButton105.Location = point;

                //Настроим высоту groupBox'ов, которые хранят в себе CheckBox'ы
                int height = groupBox8.Size.Height;
                int width = groupBox8.Size.Width;
                if (groupBox8.Controls.Count * settings.PaddingBetweenCheckBoxes > height)
                    height = groupBox8.Controls.Count * settings.PaddingBetweenCheckBoxes;
                groupBox8.Size = new Size(width, height);

                height = groupBox9.Size.Height;
                width = groupBox9.Size.Width;
                if (groupBox9.Controls.Count * settings.PaddingBetweenCheckBoxes > height)
                    height = groupBox9.Controls.Count * settings.PaddingBetweenCheckBoxes;
                groupBox9.Size = new Size(width, height);


                if (groupBox8.Size.Height > groupBox9.Size.Height)
                    groupBox7.Size = new Size(groupBox7.Size.Width, groupBox8.Size.Height + 64);
                else
                    groupBox7.Size = new Size(groupBox7.Size.Width, groupBox9.Size.Height + 64);

                //Настроим внешний вид страницы
                metroTabPage3_SizeChanged(sender, e);
            }
            #endregion

            #region Корреляционный анализ

            if (e.TabPageIndex == 3)
            {
                groupBox11.Controls.Clear();
                groupBox12.Controls.Clear();
                MetroCheckBox[] rB;

                //Скроем кнопки с других вкладок
                metroButton9.Visible = false;
                metroButton104.Visible = false;
                metroButton304.Visible = false;

                //Кнопки вывод в Word
                metroButton10.Visible = false;
                metroButton105.Visible = false;
                metroButton205.Visible = false;
                metroButton305.Visible = false;

                //Вывод признаков
                //Вывод списка количественных переменных
                rB = new MetroCheckBox[colList.Count];
                for (int i = 0; i < colList.Count; i++)
                {
                    rB[i] = new MetroCheckBox();
                    rB[i].Text = data.TakeVariableNameAtIndex(colList[i]);
                    rB[i].Checked = false;
                    rB[i].Tag = colList[i];
                    rB[i].AutoSize = true;
                    rB[i].Location = new Point(settings.StandartPadding * 2, settings.PaddingBetweenCheckBoxes * (i + 1));
                    rB[i].Padding = new Padding(settings.StandartPadding * 2);
                    rB[i].Margin = new Padding(settings.StandartPadding * 2);
                    rB[i].CheckedChanged += CheckedChangedForSigns;
                    groupBox11.Controls.Add(rB[i]);
                }

                //Изменение положения кнопок
                Point point = new Point(metroButton200.Location.X, metroButton200.Location.Y);
                point.Y = groupBox11.Size.Height + 25;
                metroButton200.Location = point;
                point = new Point(metroButton201.Location.X, metroButton201.Location.Y);
                point.Y = groupBox11.Size.Height + 25;
                metroButton201.Location = point;

                //Вывод списка порядковых переменных
                rB = new MetroCheckBox[porList.Count];
                for (int i = 0; i < porList.Count; i++)
                {
                    rB[i] = new MetroCheckBox();
                    rB[i].Text = data.TakeVariableNameAtIndex(porList[i]);
                    rB[i].Checked = false;
                    rB[i].Tag = porList[i];
                    rB[i].AutoSize = true;
                    rB[i].Location = new Point(settings.StandartPadding * 2, settings.PaddingBetweenCheckBoxes * (i + 1));
                    rB[i].Padding = new Padding(settings.StandartPadding * 2);
                    rB[i].Margin = new Padding(settings.StandartPadding * 2);
                    rB[i].CheckedChanged += CheckedChangedForSigns;
                    groupBox12.Controls.Add(rB[i]);
                }

                //Изменение положения кнопок
                point = new Point(metroButton202.Location.X, metroButton202.Location.Y);
                point.Y = groupBox12.Size.Height + 25;
                metroButton202.Location = point;
                point = new Point(metroButton203.Location.X, metroButton203.Location.Y);
                point.Y = groupBox12.Size.Height + 25;
                metroButton203.Location = point;

                //Выровняем кнопки "Выполнить сравнение" и "Вывести в Word"
                point = new Point(metroButton204.Location.X, groupBox10.Size.Height);
                point.Y += 20;
                metroButton204.Location = point;
                point = new Point(metroButton205.Location.X, groupBox10.Size.Height);
                point.Y += 20;
                metroButton205.Location = point;

                //Настроим высоту groupBox'ов, которые хранят в себе CheckBox'ы
                int height = groupBox11.Size.Height;
                int width = groupBox11.Size.Width;
                if (groupBox11.Controls.Count * settings.PaddingBetweenCheckBoxes > height)
                    height = groupBox11.Controls.Count * settings.PaddingBetweenCheckBoxes;
                groupBox11.Size = new Size(width, height);

                height = groupBox12.Size.Height;
                width = groupBox12.Size.Width;
                if (groupBox12.Controls.Count * settings.PaddingBetweenCheckBoxes > height)
                    height = groupBox12.Controls.Count * settings.PaddingBetweenCheckBoxes;
                groupBox12.Size = new Size(width, height);

                if (groupBox11.Size.Height > groupBox12.Size.Height)
                    groupBox10.Size = new Size(groupBox10.Size.Width, groupBox11.Size.Height + 64);
                else
                    groupBox10.Size = new Size(groupBox10.Size.Width, groupBox12.Size.Height + 64);

                //Настроим внешний вид страницы
                metroTabPage4_SizeChanged(sender, e);
            }

            #endregion

            #region Таблицы сопряженности

            if (e.TabPageIndex == 4)
            {
                groupBox14.Controls.Clear();
                groupBox15.Controls.Clear();
                MetroCheckBox[] rB;

                //Скроем кнопки с других вкладок
                metroButton9.Visible = false;
                metroButton104.Visible = false;
                metroButton204.Visible = false;

                //Кнопки вывод в Word
                metroButton10.Visible = false;
                metroButton105.Visible = false;
                metroButton205.Visible = false;
                metroButton305.Visible = false;

                //Вывод признаков
                //Вывод номинальных переменных
                rB = new MetroCheckBox[nomList.Count];
                for (int i = 0; i < nomList.Count; i++)
                {
                    rB[i] = new MetroCheckBox();
                    rB[i].Text = data.TakeVariableNameAtIndex(nomList[i]);
                    rB[i].Checked = false;
                    rB[i].Tag = nomList[i];
                    rB[i].AutoSize = true;
                    rB[i].Location = new Point(settings.StandartPadding * 2, settings.PaddingBetweenCheckBoxes * (i + 1));
                    rB[i].Padding = new Padding(settings.StandartPadding * 2);
                    rB[i].Margin = new Padding(settings.StandartPadding * 2);
                    rB[i].CheckedChanged += CheckedChangedForGroupFac;
                    groupBox14.Controls.Add(rB[i]);
                }

                //Изменение положения кнопок
                Point point = new Point(metroButton300.Location.X, metroButton300.Location.Y);
                point.Y = groupBox14.Size.Height + 25;
                metroButton300.Location = point;
                point = new Point(metroButton301.Location.X, metroButton301.Location.Y);
                point.Y = groupBox14.Size.Height + 25;
                metroButton301.Location = point;

                //Вывод бинарных переменных
                rB = new MetroCheckBox[binList.Count];
                for (int i = 0; i < binList.Count; i++)
                {
                    rB[i] = new MetroCheckBox();
                    rB[i].Text = data.TakeVariableNameAtIndex(binList[i]);
                    rB[i].Checked = false;
                    rB[i].Tag = binList[i];
                    rB[i].AutoSize = true;
                    rB[i].Location = new Point(settings.StandartPadding * 2, settings.PaddingBetweenCheckBoxes * (i + 1));
                    rB[i].Padding = new Padding(settings.StandartPadding * 2);
                    rB[i].Margin = new Padding(settings.StandartPadding * 2);
                    rB[i].CheckedChanged += CheckedChangedForGroupFac;
                    groupBox15.Controls.Add(rB[i]);
                }

                //Изменение положения кнопок
                point = new Point(metroButton302.Location.X, metroButton302.Location.Y);
                point.Y = groupBox15.Size.Height + 25;
                metroButton302.Location = point;
                point = new Point(metroButton303.Location.X, metroButton303.Location.Y);
                point.Y = groupBox15.Size.Height + 25;
                metroButton303.Location = point;

                //Выровняем кнопки "Выполнить сравнение" и "Вывести в Word"
                point = new Point(metroButton304.Location.X, groupBox13.Size.Height);
                point.Y += 20;
                metroButton304.Location = point;
                point = new Point(metroButton305.Location.X, groupBox13.Size.Height);
                point.Y += 20;
                metroButton305.Location = point;

                //Настроим высоту groupBox'ов, которые хранят в себе CheckBox'ы
                int height = groupBox14.Size.Height;
                int width = groupBox14.Size.Width;
                if (groupBox14.Controls.Count * settings.PaddingBetweenCheckBoxes > height)
                    height = groupBox14.Controls.Count * settings.PaddingBetweenCheckBoxes;
                groupBox14.Size = new Size(width, height);

                height = groupBox15.Size.Height;
                width = groupBox15.Size.Width;
                if (groupBox15.Controls.Count * settings.PaddingBetweenCheckBoxes > height)
                    height = groupBox15.Controls.Count * settings.PaddingBetweenCheckBoxes;
                groupBox15.Size = new Size(width, height);

                if (groupBox14.Size.Height > groupBox15.Size.Height)
                    groupBox13.Size = new Size(groupBox13.Size.Width, groupBox14.Size.Height + 64);
                else
                    groupBox13.Size = new Size(groupBox13.Size.Width, groupBox15.Size.Height + 64);

                //Настроим внешний вид страницы
                metroTabPage5_SizeChanged(sender, e);
            }

            #endregion
            
        }

        /// <summary>
        /// Выбран группирующий фактор
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CheckedChangedForGroupFac(object sender, EventArgs e)
        {
            MetroCheckBox mCB = (MetroCheckBox)sender;
            //В поле Tag лежит индекс элемента
            int index = (int)mCB.Tag;
            //Поставили галочку - добавить в список выбранных элементов и отобразить список признаков
            //Убрали - убрать из списка и, если сняли все галочки, то скрыть признаки
            if (mCB.Checked == true)
            {
                if (groupFactList.IndexOf(index) == -1)
                {
                    groupFactList.Add(index);
                    groupBox4.Visible = true;
                }
                if (Convert.ToInt32(mCB.Parent.Tag) == 4 && groupFactList.Count >= 2)
                    metroButton304.Visible = true;
            }
            else
            {
                groupFactList.Remove(index);
                if (groupFactList.Count == 0)
                    groupBox4.Visible = false;
                if (Convert.ToInt32(mCB.Parent.Tag) == 4 && groupFactList.Count < 2)
                    metroButton304.Visible = false;
            }
        }

        /// <summary>
        /// Выбран признак
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CheckedChangedForSigns(object sender, EventArgs e)
        {
            MetroCheckBox mCB = (MetroCheckBox)sender;
            //В поле Tag лежит индекс элемента
            int index = (int)mCB.Tag;
            //Тэг родительского элемента, в котором лежит чекбокс признака
            int parentTag = Convert.ToInt32(mCB.Parent.Tag);
            //Поставили галочку - добавить в список выбранных элементов и отобразить кнопку для сравнения
            //Убрали - убрать из списка и, если сняли все галочки, то скрыть кнопку для сравнения
            if (mCB.Checked == true)
            {
                signsList.Add(index);
                //Отобразим соответствующую кнопку обработки
                switch(parentTag)
                {
                    case 1: metroButton9.Visible = true; break;
                    case 2: metroButton104.Visible = true; break;
                    case 3: metroButton204.Visible = true; break;
                    //case 4: if (signsList.Count == 2) metroButton302.Visible = true; break;
                }
            }
            else
            {
                signsList.Remove(index);
                if (signsList.Count == 0)
                    //Отобразим соответствующую кнопку обработки
                    switch (parentTag)
                    {
                        case 1: metroButton9.Visible = false; metroButton10.Visible = false; break;
                        case 2: metroButton104.Visible = false; metroButton105.Visible = false; break;
                        case 3: metroButton204.Visible = false; metroButton205.Visible = false; break;
                        //case 4: metroButton302.Visible = false; metroButton303.Visible = false; break;
                    }
                //if (signsList.Count == 1)
                //    metroButton302.Visible = false; metroButton303.Visible = false;
            }
        }

        /// <summary>
        /// Нажата кнопка "Выбрать все"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SelectAll_Click(object sender, EventArgs e)
        {
            //Получим Тэг кнопки
            MetroButton mB = (MetroButton)sender;
            int tag = Convert.ToInt32(mB.Tag);

            //По Тэгу производим действия: добавляем в список выбранных элементов и проставляем галочки
            switch(tag)
            {
                //"Сравнение независимых групп"
                //Бинарные переменные
                case 1: Check(groupBox2); break;
                //Номинальные переменные
                case 3: Check(groupBox3); break;
                //Количественные переменные
                case 5: Check(groupBox5); break;
                //Порядковые переменные
                case 7: Check(groupBox6); break;

                //"Сравнение зависимых групп"
                //Количественные переменные
                case 9: Check(groupBox8); break;
                //Порядковые переменные
                case 11: Check(groupBox9); break;

                //"Корреляционный анализ"
                //Количественные переменные
                case 13: Check(groupBox11); break;
                //Порядковые переменные
                case 15: Check(groupBox12); break;

                //"Таблицы сопряженности"
                //Номинальные переменные
                case 17: Check(groupBox14); break;
                //Бинарные переменные
                case 19: Check(groupBox15); break;
            }
        }

        /// <summary>
        /// Нажата кнопка "Очистить"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CleanAll_Click(object sender, EventArgs e)
        {
            //Получим Тэг кнопки
            MetroButton mB = (MetroButton)sender;
            int tag = Convert.ToInt32(mB.Tag);

            //По Тэгу производим действия: удаляем элементы из списка и снимаем все галочки
            switch (tag)
            {
                //"Сравнение независимых групп"
                //Бинарные переменные
                case 2: Uncheck(groupBox2); break;
                //Номинальные переменные
                case 4: Uncheck(groupBox3); break;
                //Количественные переменные
                case 6: Uncheck(groupBox5); break;
                //Порядковые переменные
                case 8: Uncheck(groupBox6); break;
                    
                //"Сравнение зависимых групп"
                //Количественные переменные
                case 10: Uncheck(groupBox8); break;
                //Порядковые переменные
                case 12: Uncheck(groupBox9); break;

                //"Корреляционный анализ"
                //Количественные переменные
                case 14: Uncheck(groupBox11); break;
                //Порядковые переменные
                case 16: Uncheck(groupBox12); break;

                //"Таблицы сопряженности"
                //Номинальные переменные
                case 18: Uncheck(groupBox14); break;
                //Бинарные переменные
                case 20: Uncheck(groupBox15); break;
            }

            //Если сняли все галочки на группирующих переменных, то
            //Очистим список выбранных признаков, скроем поле с их выбором и снимем галочки
            if (groupFactList.Count == 0)
            {
                Uncheck(groupBox5);
                Uncheck(groupBox6);
                groupBox4.Visible = false;
            }
        }

        /// <summary>
        /// Проставить галочки на выбранных элементах
        /// </summary>
        /// <param name="_groupBox"></param>
        private void Check(GroupBox _groupBox)
        {
            Control.ControlCollection controls = _groupBox.Controls;
            foreach(MetroCheckBox control in controls)
                control.Checked = true;       
        }

        /// <summary>
        /// Снять галочки
        /// </summary>
        /// <param name="_groupBox"></param>
        private void Uncheck(GroupBox _groupBox)
        {
            Control.ControlCollection controls = _groupBox.Controls;
            foreach (MetroCheckBox control in controls)
                control.Checked = false;
        }
        
        /// <summary>
        /// Заполнение результата дополнительными данными (медиана, квантили и отклонение)
        /// </summary>
        /// <param name="_s"></param>
        /// <param name="_t"></param>
        /// <returns></returns>
        private Sample FillingResults(List<Sample.SubSample> _s, Sample _t)
        {
            NumericVector X;
            for (int i = 0; i < _s.Count; i++)
            {
                X = engine.CreateNumericVector(_s[i].SampleList);
                engine.SetSymbol("X", X);

                var median = engine.Evaluate("median(X)").AsNumeric();
                var sD = engine.Evaluate("sd(X)").AsNumeric();
                var quantile = engine.Evaluate("quantile(X)").AsNumeric();
                var tmp = _t.SubSampleList[i];
                tmp.LowerQuintile = quantile[1]; //0
                tmp.TopQuintile = quantile[3]; //4
                tmp.Median = median[0];
                tmp.StandartDeviation = sD[0];
                _t.SubSampleList[i] = tmp;
            }
            return _t;
        }

        #region Сравнение независимых групп

        /// <summary>
        /// Выоплнить "Сравнение независимых групп"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void metroButton9_Click(object sender, EventArgs e)
        {
            Sample testRes;
            GenericVector tmp;
            engine = REngine.GetInstance();
            List<Sample> samples = new List<Sample>();
            samples = Sample.GetSample(data, groupFactList, signsList);

            //Очистим списки результатов
            resIndMannaWhitney.Clear();
            resIndKruskalWallis.Clear();
            
            for (int i = 0; i < samples.Count; i++)
            {
                testRes = new Sample();
                testRes = samples[i];
                if (samples[i].SubSampleList.Count == 2)
                {
                    //Огриничение на количество элементов в выборке. В каждой выборке их должно быть не менее 3
                    if (samples[i].SubSampleList[0].SampleList.Count >= 3 && samples[i].SubSampleList[1].SampleList.Count >= 3)
                    {
                        //Вызвать тест Манна-Уитни
                        NumericVector group1 = engine.CreateNumericVector(samples[i].SubSampleList[0].SampleList);
                        engine.SetSymbol("Rgroup1", group1);
                        NumericVector group2 = engine.CreateNumericVector(samples[i].SubSampleList[1].SampleList);
                        engine.SetSymbol("Rgroup2", group2);
                        tmp = engine.Evaluate("wilcox.test(Rgroup1, Rgroup2, paired = FALSE)").AsList();
                        double p = tmp["p.value"].AsNumeric().First();
                        testRes.PValue = p;

                        //Посчитаем медиану и стандартное отклонение
                        testRes = FillingResults(samples[i].SubSampleList, testRes);
                        resIndMannaWhitney.Add(testRes);

                        //Нарисовать график для выборки
                        CreateGraphic(samples[i], "MW");
                    }
                    //else
                        //TODO: сохранить информацию о срабатывании ограничения на выборку и вывести ее в отчет 
                }
                else
                {
                    try
                    {
                        //Вызвать тест Краскела-Уоллиса
                        //Создадим список выборок для отправки в тест Краскела-Уоллиса
                        GenericVector gV = new GenericVector(engine, samples[i].SubSampleList.Count);
                        NumericVector nV;
                        for (int j = 0; j < samples[i].SubSampleList.Count; j++)
                        {
                            nV = new NumericVector(engine, samples[i].SubSampleList[j].SampleList);
                            if (nV.Length < 3) throw new Exception();
                            gV[j] = nV;
                        }

                        engine.SetSymbol("sample", gV);
                        tmp = engine.Evaluate("kruskal.test(sample)").AsList();
                        double p = tmp["p.value"].AsNumeric().First();
                        testRes.PValue = p;

                        //Посчитаем медиану и стандартное отклонение
                        testRes = FillingResults(samples[i].SubSampleList, testRes);
                        resIndKruskalWallis.Add(testRes);

                        //Нарисовать график для выборки
                        CreateGraphic(samples[i], "KW");
                    }
                    catch(Exception)
                    {
                        //TODO: сохранить информацию о срабатывании ограничения на выборку и вывести ее в отчет 
                    }
                }
            }
            //Отобразим кнопку для вывода в Word
            metroButton10.Visible = true;
        }

        /// <summary>
        /// Вывод результатов проверки независимых выборок в Word
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void metroButton10_Click(object sender, EventArgs e)
        {
            WordDocument report = new WordDocument(WordDocumentFormat.InCentimeters(21, 29.7, 2.5, 1, 2, 2));
            report.WriteControlWord(@"sl360\slmult1");
            report.SetFont(settings.FontStandart);
            report.SetTextAlign(WordTextAlign.Center);
            report.WriteLine("Сравнение средних по 2 независимым выборкам ");

            report = OutResultInTableIndependent(report, "Манна-Уитни", resIndMannaWhitney);
            report = OutResultInTableIndependent(report, "Краскела-Уоллиса", resIndKruskalWallis);

            report.SaveToFile("..\\..\\ResultIndep.doc");
            System.Diagnostics.Process.Start("..\\..\\ResultIndep.doc");

            //Удалим все созданные графики
            var files = Directory.GetFiles(pathC);
            foreach(string file in files)
                File.Delete(file);
        }

        /// <summary>
        /// Вывод информации в таблицы
        /// </summary>
        /// <param name="_wordDocument">Документ</param>
        /// <param name="_methodName">Название метода</param>
        /// <param name="_resList">Список с результатами проверки</param>
        /// <returns></returns>
        private WordDocument OutResultInTableIndependent(WordDocument _wordDocument, string _methodName, List<Sample> _resList)
        {
            //Получим пути созданных графиков для текущего метода
            string[] dirs;
            string methodName = "";
            bool pval = false;
            if (_methodName.Contains("Манна"))
            {
                dirs = Directory.GetFiles(pathC, "MW*"); //Результаты для Манна-Уитни
                methodName = "MW";
            }
            else
            {
                dirs = Directory.GetFiles(pathC, "KW*"); //Результаты для Краскела-Уоллиса
                methodName = "KW";
            }
            var grpRes = Grouping(_resList);

            //Вывод результата
            for (int i = 0; i < grpRes.Count; i++)
            {
                WordTable rt1;
                Sample s;

                //Текст перед таблицей
                string grpFact = grpRes[i][0].GroupFact;
                string text = "В теблице " + (i + 1).ToString() + " приведен результат статистического анализа данных, проведенного с использованием " 
                              + "критерия " + _methodName + "." + " В качестве группирующего фактора используется переменная " + "\"" + grpFact + "\"" + ".";
                _wordDocument.SetTextAlign(WordTextAlign.Justified);
                _wordDocument.SetParagraph(0, 567);
                _wordDocument.WriteLine(text);

                //Название таблицы
                string tableNumber = Environment.NewLine + "Таблица " + (i + 1).ToString() + " - ";  
                string tableCaption = "Сравнение средних уровней переменных в группах: ";
                for (int j = 0; j < grpRes[i].Count - 1; j++)
                    tableCaption += "\"" + grpRes[i][j].NameSign + "\"" + ", ";
                tableCaption += "\"" + grpRes[i].Last().NameSign + "\"";
                _wordDocument.SetTextAlign(WordTextAlign.Left);
                _wordDocument.SetParagraph(0, 0);
                _wordDocument.WriteLine(tableNumber + tableCaption);

                //Создаем табилицу и заполняем шапку
                _wordDocument.SetParagraph(0, 0);
                rt1 = _wordDocument.NewTable(settings.FontStandart, Color.Black, grpRes[i].Count + 1, grpRes[i][0].SubSampleList.Count + 2, 2);
                //Заполнение таблицы
                rt1.Rows[0][0].Write("Показатель");
                for (int k = 0; k < grpRes[i][0].SubSampleList.Count; k++)
                    rt1.Rows[0][1 + k].Write(grpRes[i][0].SubSampleList[k].UniqueVal + "\n(n = " + grpRes[i][0].SubSampleList[k].SampleList.Count.ToString() + ")");
                rt1.Rows[0][grpRes[i][0].SubSampleList.Count + 1].Write("p-значение");

                //Отрисуем рамки у ячеек
                for (int j = 0; j < grpRes[i][0].SubSampleList.Count + 3; j++)
                    rt1.Rows[0][j].SetBorders(Color.Black, 1, true, true, true, true);

                //Выводим данные в таблицу
                List<string> statSignif = new List<string>();
                for (int j = 0; j < grpRes[i].Count; j++)
                {
                    s = grpRes[i][j];
                    _wordDocument.SetTextAlign(WordTextAlign.Left);
                    rt1.Rows[j + 1][0].Write(s.NameSign);
                    for (int k = 0; k < grpRes[i][0].SubSampleList.Count; k++)
                    {
                        rt1.Rows[j + 1][1 + k].WriteLine(s.SubSampleList[k].AverageValue.ToString() + " " + (char)177 + " " + Math.Round(s.SubSampleList[k].StandartDeviation, 3).ToString());
                        rt1.Rows[j + 1][1 + k].WriteLine(Math.Round(s.SubSampleList[k].Median, 3).ToString());
                        rt1.Rows[j + 1][1 + k].Write("(" + Math.Round(s.SubSampleList[k].LowerQuintile, 3).ToString() + "; " + Math.Round(s.SubSampleList[k].TopQuintile, 3).ToString() + ")");
                    }

                    double p = Math.Round(grpRes[i][j].PValue, 3);
                    if (j == 0) p = 0.03;
                    if (j == 1) p = 0.0003;
                    if (p <= 0.05)
                    {
                        if (p > 0.001)
                        {
                            rt1.SetFont(settings.FontBold);
                            rt1.Rows[j + 1][grpRes[i][0].SubSampleList.Count + 1].Write(p.ToString() + "*");
                            pval = true;
                            statSignif.Add(grpRes[i][j].NameSign);
                        }
                        else
                        {
                            rt1.SetFont(settings.FontStandart);
                            rt1.Rows[j + 1][grpRes[i][0].SubSampleList.Count + 1].Write("p < 0.001");
                            statSignif.Add(grpRes[i][j].NameSign);
                        }
                    }
                    else
                    {
                        rt1.SetFont(settings.FontStandart);
                        rt1.Rows[j + 1][grpRes[i][0].SubSampleList.Count + 1].Write(p.ToString());
                    }
                    rt1.SetFont(settings.FontStandart);

                    //Отрисуем рамки у ячеек
                    for (int k = 0; k < grpRes[i][0].SubSampleList.Count + 2; k++)
                        rt1.Rows[j + 1][k].SetBorders(Color.Black, 1, true, true, true, true);
                }
                rt1.SaveToDocument(10000, 0);

                //Примечание к таблице
                _wordDocument.WriteControlWord(@"sl360\slmult1");
                _wordDocument.SetTextAlign(WordTextAlign.Justified);
                string note = "Примечание: формат представления информации в ячейке: среднее значение " + (char)177 + 
                              " среднеквадратическое отклонение, медиана, (нижний; верхний) квартили.";

                //Добавление информации о "*" в таблице, при условии, что в ней есть p-value < 0.05
                string star = "";
                if (pval)
                {
                    star = " * - статистически значимое различие между ";
                    if (grpRes[i][0].SubSampleList.Count == 2)
                    {
                        star += "\"" + grpRes[i][0].SubSampleList[0].UniqueVal + "\"" + " и " + "\"" + grpRes[i][0].SubSampleList[1].UniqueVal + "\"";
                    }
                    else
                    {
                        for (int k = 0; k < grpRes[i][0].SubSampleList.Count - 1; k++)
                            star += "\"" + grpRes[i][0].SubSampleList[k].UniqueVal + "\"";
                        star += " и " + "\"" + grpRes[i][0].SubSampleList.Last().UniqueVal + "\"";
                    }
                    star += " p < 0.05.";
                }
                _wordDocument.WriteLine(note + star);

                //Вывод из таблицы
                _wordDocument.SetParagraph(0, 567);
                string conclusion = "Из таблицы " + (i + 1).ToString() + " видно, что ";
                if (pval)
                {
                    conclusion += "статистически значимое различие есть у ";
                    if (statSignif.Count == 1)
                    {
                        conclusion += "признака " + "\"" + statSignif.First() + "\", ";
                    }
                    else
                    {
                        conclusion += "признаков ";
                        for (int j = 0; j < statSignif.Count; j++)
                            conclusion += "\"" + statSignif[j] + "\"" + ", ";
                    }
                    conclusion += " с учетом уровня значимости равного 0.05.";
                }
                else
                {
                    conclusion += " статистически значимых различий не обнаружено. С уровнем значимости равным 0.05.";
                }
                _wordDocument.WriteLine(conclusion);

                //Предисловие к графикам
                string preface = "";
                if (grpRes[i].Count == 1)
                {
                    preface += "На рисунке 1 изображена диаграмма размахов признака, приведенного в таблице " + (i + 1).ToString() + ".";
                }
                else
                {
                    preface += "На рисунках 1 - " + grpRes[i].Count.ToString() + " изображены диаграммы размахов признаков, приведенных в таблице " 
                        + (i + 1).ToString() + ".";
                }
                _wordDocument.WriteLine(preface);

                //Вставим график
                _wordDocument.SetTextAlign(WordTextAlign.Center);
                List<string> graphNames = new List<string>();
                //Получим список названий графиков
                for (int j = 0; j < dirs.Length; j++)
                {
                    string name = Path.GetFileNameWithoutExtension(dirs[j]);
                    graphNames.Add(name);
                }
                //Добавляем графики
                for (int j = 0; j < grpRes[i].Count; j++)
                {
                    int curGraph = graphNames.IndexOf(methodName + "_" + grpRes[i][j].GroupFact + "_" + grpRes[i][j].NameSign);
                    _wordDocument.PutImage(dirs[curGraph], 96); //96 - истинный dpi
                    _wordDocument.WriteLine();

                    //Подпись к графику
                    _wordDocument.WriteControlWord(@"sl360\slmult1");
                    note = "Рисунок " + (j + 1).ToString() + " - " +
                            "Диаграмма размаха переменной " + "\"" + grpRes[i][j].NameSign + "\"";
                    _wordDocument.WriteLine(note);
                }
            }
            
            return _wordDocument;
        }

        /// <summary>
        /// Группирование данных по-группирующему признаку
        /// </summary>
        /// <param name="_testRes"></param>
        /// <returns></returns>
        private List<List<Sample>> Grouping(List<Sample> _testRes)
        {
            List<List<Sample>> res = new List<List<Sample>>();
            List<Sample> list = new List<Sample>();
            string grpFact = "";

            for (int i = 0; i < _testRes.Count; i++)
            {
                //Берем группирующий признак
                grpFact = _testRes[i].GroupFact;
                for (int j = i; j < _testRes.Count; j++)
                {
                    if (grpFact == _testRes[j].GroupFact)
                    {
                        list.Add(_testRes[j]);
                    }
                    else
                    {
                        i = j - 1;
                        break;
                    }
                    if (j + 1 == _testRes.Count)
                    {
                        i = _testRes.Count;
                    }
                }
                res.Add(list);
                list = new List<Sample>();
            }
            return res;
        }

        #endregion

        #region Сравнение зависимых групп

        /// <summary>
        /// Выоплнить "Сравнение зависимых групп"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void metroButton15_Click(object sender, EventArgs e)
        {
            engine = REngine.GetInstance();
            GenericVector tmp;
            Sample sample = new Sample();
            sample = Sample.GetSample(data, signsList);

            //Очистим списки результатов
            resDepWilcoxon.Clear();
            resDepFridman.Clear();

            if (signsList.Count == 2)
            {
                //Вилкоксон
                NumericVector group1 = engine.CreateNumericVector(sample.SubSampleList[0].SampleList);
                engine.SetSymbol("Rgroup1", group1);
                NumericVector group2 = engine.CreateNumericVector(sample.SubSampleList[1].SampleList);
                engine.SetSymbol("Rgroup2", group2);
                tmp = engine.Evaluate("wilcox.test(Rgroup1, Rgroup2)").AsList();
                double p = tmp["p.value"].AsNumeric().First();
                sample.PValue = p;

                //Посчитаем медиану и стандартное отклонение
                sample = FillingResults(sample.SubSampleList, sample);
                resDepWilcoxon.Add(sample);
            }
            else
            {
                //Фридман
                //Создадим список выборок для отправки в тест Фридмана
                GenericVector gV = new GenericVector(engine, sample.SubSampleList.Count);
                NumericVector nV;
                for (int j = 0; j < sample.SubSampleList.Count; j++)
                {
                    nV = new NumericVector(engine, sample.SubSampleList[j].SampleList);
                    gV[j] = nV;
                }

                engine.SetSymbol("sample", gV);
                tmp = engine.Evaluate("kruskal.test(sample)").AsList();
                double p = tmp["p.value"].AsNumeric().First();
                sample.PValue = p;

                //Посчитаем медиану и стандартное отклонение
                sample = FillingResults(sample.SubSampleList, sample);
                resDepFridman.Add(sample);
            }
            //Отобразим кнопку для вывода в Word
            metroButton105.Visible = true;
        }

        /// <summary>
        /// Вывод результатов проверки зависимых выборок в Word
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void metroButton16_Click(object sender, EventArgs e)
        {
            WordDocument report = new WordDocument(WordDocumentFormat.InCentimeters(21, 29.7, 2.5, 1, 2, 2));

            report.SetFont(settings.FontStandart);
            report.SetTextAlign(WordTextAlign.Center);
            report.WriteLine("Сравнение средних по 2 зависимым выборкам ");

            report = OutResultInTableDependent(report, "Критерий Вилкоксона", resDepWilcoxon);
            report = OutResultInTableDependent(report, "Критерий Фридмана", resDepFridman);

            report.SaveToFile("..\\..\\ResultDep.doc");
            System.Diagnostics.Process.Start("..\\..\\ResultDep.doc");
        }

        /// <summary>
        /// Вывод информации в таблицы
        /// </summary>
        /// <param name="_wordDocument">Документ</param>
        /// <param name="_methodName">Название метода</param>
        /// <param name="_resList">Список с результатами проверки</param>
        /// <returns></returns>
        private WordDocument OutResultInTableDependent(WordDocument _wordDocument, string _methodName, List<Sample> _resList)
        {
            WordTable rt1;

            var grpRes = Grouping(_resList);
            for (int i = 0; i < grpRes.Count; i++)
            {
                Sample s;
                //Создаем табилицу и заполняем шапку
                rt1 = _wordDocument.NewTable(settings.FontStandart, Color.Black, grpRes[i].Count + 1, grpRes[i][0].SubSampleList.Count + 1, 2);
                _wordDocument.SetTextAlign(WordTextAlign.Left);
                _wordDocument.SetFont(settings.FontStandart);
                _wordDocument.SetTextAlign(WordTextAlign.Left);
                _wordDocument.WriteLine();
                _wordDocument.WriteLine(_methodName);
                for (int k = 0; k < grpRes[i][0].SubSampleList.Count; k++)
                    rt1.Rows[0][k].Write(grpRes[i][0].SubSampleList[k].UniqueVal + ", n = " + grpRes[i][0].SubSampleList[k].SampleList.Count.ToString());
                rt1.Rows[0][grpRes[i][0].SubSampleList.Count].Write("p-значение");

                //Отрисуем рамки у ячеек
                for (int j = 0; j < grpRes[i][0].SubSampleList.Count + 1; j++)
                    rt1.Rows[0][j].SetBorders(Color.Black, 1, true, true, true, true);

                //Выводим данные в таблицу
                for (int j = 0; j < grpRes[i].Count; j++)
                {
                    s = grpRes[i][j];
                    _wordDocument.SetTextAlign(WordTextAlign.Left);
                    for (int k = 0; k < grpRes[i][0].SubSampleList.Count; k++)
                    {
                        rt1.Rows[j + 1][k].WriteLine(s.SubSampleList[k].AverageValue.ToString() + " " + (char)177 + " " + Math.Round(s.SubSampleList[k].StandartDeviation, 3).ToString());
                        rt1.Rows[j + 1][k].WriteLine(Math.Round(s.SubSampleList[k].Median, 3).ToString());
                        rt1.Rows[j + 1][k].Write("(" + Math.Round(s.SubSampleList[k].LowerQuintile, 3).ToString() + "; " + Math.Round(s.SubSampleList[k].TopQuintile, 3).ToString() + ")");
                    }

                    double p = grpRes[i][j].PValue;
                    rt1.Rows[j + 1][grpRes[i][0].SubSampleList.Count].Write((Math.Round(p, 3).ToString()));

                    //Отрисуем рамки у ячеек
                    for (int k = 0; k < grpRes[i][0].SubSampleList.Count + 1; k++)
                        rt1.Rows[j + 1][k].SetBorders(Color.Black, 1, true, true, true, true);
                }
                rt1.SaveToDocument(9600, 0);
            }
            return _wordDocument;
        }

        #endregion

        #region Корреляционный анализ

        /// <summary>
        /// Выполнить "Корреляционный анализ"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void metroButton21_Click(object sender, EventArgs e)
        {
            engine = REngine.GetInstance();
            Sample sample = new Sample();
            СorrelationResult cor = new СorrelationResult();
            sample = Sample.GetSample(data, signsList);
            int colSample = sample.SubSampleList.Count;

            //Очистим списки результатов
            resCorPearson.Clear();
            resCorSpearman.Clear();

            for (int i = 0; i < colSample; i++)
            {
                string ThatCompare = data.TakeVariableNameAtIndex(signsList[i]);
                //Создаем первый вектор, который будем сравнивать (что сравниваем)
                int sampleSize = sample.SubSampleList[0].SampleList.Count;
                NumericVector x = new NumericVector(engine, sample.SubSampleList[i].SampleList);
                engine.SetSymbol("x", x);

                //StreamWriter sw = new StreamWriter("outX.txt");
                //string buf = "";
                //for (int a = 0; a < x.Length; a++)
                //    buf += x[a].ToString() + Environment.NewLine;
                //sw.WriteLine(buf);
                //sw.Close();

                for (int j = 0; j < i; j++)
                {
                    string ToCompare = data.TakeVariableNameAtIndex(signsList[j]);
                    //Создаем второй вектор (с чем сравниваем)
                    NumericVector y = new NumericVector(engine, sample.SubSampleList[j].SampleList);
                    engine.SetSymbol("y", y);

                    //sw = new StreamWriter("outY.txt");
                    //buf = "";
                    //for (int a = 0; a < x.Length; a++)
                    //    buf += y[a].ToString() + Environment.NewLine;
                    //sw.WriteLine(buf);
                    //sw.Close();

                    GenericVector tmpRes;

                    //Пирсон
                    tmpRes = engine.Evaluate("cor.test(x, y, method=\"pearson\")").AsList();
                    cor.p = tmpRes["p.value"].AsNumeric().First();
                    cor.r = tmpRes["statistic"].AsNumeric().First();
                    cor.thatCompare = signsList[i];
                    string tC = data.TakeVariableNameAtIndex(cor.thatCompare);
                    cor.toCompare = signsList[j];
                    string toC = data.TakeVariableNameAtIndex(cor.toCompare);
                    resCorPearson.Add(cor);
                    cor = new СorrelationResult();

                    //Спирмен
                    tmpRes = engine.Evaluate("cor.test(x, y, method=\"spearman\")").AsList();
                    cor.p = tmpRes["p.value"].AsNumeric().First();
                    cor.r = tmpRes["statistic"].AsNumeric().First();
                    cor.thatCompare = signsList[i];
                    cor.toCompare = signsList[j];
                    resCorSpearman.Add(cor);
                    cor = new СorrelationResult();
                }
            }
            //Отобразим кнопку для вывода в Word
            metroButton205.Visible = true;
        }

        /// <summary>
        /// Вывод результатов корреляционного анализа в Word
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void metroButton22_Click(object sender, EventArgs e)
        {
            WordDocument report = new WordDocument(WordDocumentFormat.InCentimeters(21, 29.7, 2.5, 1, 2, 2));

            report.SetFont(settings.FontStandart);
            report.SetTextAlign(WordTextAlign.Center);
            report.WriteLine("Корреляционный анализ");

            report = OutResultInTableCorrelation(report, "Метод Пирсона", resCorPearson);
            report = OutResultInTableCorrelation(report, "Метод Спирмена", resCorSpearman);

            report.SaveToFile("..\\..\\ResultCor.doc");
            System.Diagnostics.Process.Start("..\\..\\ResultCor.doc");
        }

        /// <summary>
        /// Вывод информации в таблицы
        /// </summary>
        /// <param name="_wordDocument">Документ</param>
        /// <param name="_methodName">Название метода</param>
        /// <param name="_resList">Список с результатами проверки</param>
        /// <returns></returns>
        private WordDocument OutResultInTableCorrelation(WordDocument _wordDocument, string _methodName, List<СorrelationResult> _result)
        {
            //максимальное размер таблицы
            int MaxSize = 5;
            //Размер общей таблицы
            int sizeAll = signsList.Count;
            //Сколько блоков в столбце
            int NumBlock = sizeAll / MaxSize;
            //остаток от блоков. Если ли еще блок, меньший по размеру чем maxColumn
            if (sizeAll % MaxSize > 0) NumBlock++;

            //Список таблиц
            List<WordTable> tables = new List<WordTable>();

            //Заполним список таблицами с начальными данными: заголовками и размерами
            int widthTable, heightTable;
            for (int i = 0; i < NumBlock; i++)
            {
                //Ширина таблицы
                if (i + 1 == NumBlock)
                    if (sizeAll % MaxSize != 0)
                        widthTable = sizeAll % MaxSize;
                    else
                        widthTable = MaxSize;
                else
                    widthTable = MaxSize;

                for (int j = 0; j <= i; j++)
                {
                    //Высота таблицы
                    if (j + 1 == NumBlock)
                        if (sizeAll % MaxSize != 0)
                            heightTable = sizeAll % MaxSize;
                        else
                            heightTable = MaxSize;
                    else
                        heightTable = MaxSize;

                    WordTable wt = _wordDocument.NewTable(settings.FontStandart, Color.Black, widthTable + 1, heightTable + 1, 2);
                    //Заполняем первую строку таблицы
                    string strRow = "";
                    for (int k = 0; k < widthTable; k++)
                    {
                        strRow = data.TakeVariableNameAtIndex(signsList[i * MaxSize + k]);
                        wt.Rows[k + 1][0].Write(strRow);
                    }
                    //Заполняем первый столбец таблицы
                    string strCol = "";
                    for (int k = 0; k < heightTable; k++)
                    {
                        strCol = data.TakeVariableNameAtIndex(signsList[j * MaxSize + k]);
                        wt.Rows[0][k + 1].Write(strCol);
                    }
                    //Отрисуем рамки у ячеек
                    for (int k = 0; k < widthTable + 1; k++)
                        for (int l = 0; l < heightTable + 1; l++)
                            wt.Rows[k][l].SetBorders(Color.Black, 1, true, true, true, true);

                    _wordDocument.SetTextAlign(WordTextAlign.Left);
                    _wordDocument.SetFont(settings.FontStandart);
                    _wordDocument.SetTextAlign(WordTextAlign.Left);
                    _wordDocument.WriteLine(Environment.NewLine + "Таблица " + (tables.Count + 1).ToString() + " - Корреляционный анализ. " + _methodName);

                    tables.Add(wt);
                    wt.SaveToDocument(9600, 0);

                    //Добавим примечание
                    _wordDocument.SetTextAlign(WordTextAlign.Justified);
                    _wordDocument.SetFont(settings.FontStandart);
                    _wordDocument.WriteLine("Примечание: r - коэффициент корреляции; p - уровень статистической значимости; жирным шрифтом " +
                                            "выделена статистически значимая связь.");
                }
            }

            //Заполнение первой таблицы
            WordTable w = tables[0];
            int tmp = 0;
            for (int i = 0; i < w.Rows.Length - 1; i++)
            {
                for (int j = 0; j <= i; j++)
                {
                    if (i == j)
                        w.Rows[i + 1][j + 1].Write("r = 1");
                    else
                    {
                        //Настройка вывода ячейки
                        double p = _result[tmp].p;
                        double r = Math.Round(_result[tmp].r, 3);

                        if (p <= 0.05)
                        {
                            w.Rows[i + 1][j + 1].SetFont(settings.FontStandart);
                            w.Rows[i + 1][j + 1].WriteLine("r = " + Math.Round(r, 3).ToString());
                            if (p > 0.001)
                                w.Rows[i + 1][j + 1].Write("p = " + Math.Round(p, 3).ToString());
                            else
                                w.Rows[i + 1][j + 1].Write("p < 0.001");
                        }
                        else
                        {
                            w.Rows[i + 1][j + 1].SetFont(settings.FontStandart);
                            w.Rows[i + 1][j + 1].WriteLine("r = " + Math.Round(r, 3).ToString());
                            w.Rows[i + 1][j + 1].Write("p = " + Math.Round(p, 3).ToString());
                        }
                        tmp++;
                    }
                }
            }
            tables.RemoveAt(0);
            tables.Insert(0, w);

            //Заполним ячейки таблиц данными
            int str = 0;
            int count = 1;
            int diag = 2, diagPP = 3, dropBlock = 1;
            bool dropCount = false, strPP = false;
            for (int i = tmp; i < _result.Count;)
            {
                //Берем таблицу
                WordTable wt = tables[count];

                //Диагональная ли таблица?
                if (count == diag)
                {
                    //Да
                    for (int j = 0; j <= str; j++)
                    {
                        if (j == str)
                        {
                            wt.Rows[str + 1][j + 1].Write("r = 1");
                        }
                        else
                        {
                            //Настройка вывода ячейки
                            double p = _result[i].p;
                            double r = Math.Round(_result[i].r, 3);

                            if (p <= 0.05)
                            {
                                wt.Rows[str + 1][j + 1].SetFont(settings.FontStandart);
                                wt.Rows[str + 1][j + 1].WriteLine("r = " + Math.Round(r, 3).ToString());
                                if (p > 0.001)
                                    wt.Rows[str + 1][j + 1].Write("p = " + Math.Round(p, 3).ToString());
                                else
                                    wt.Rows[str + 1][j + 1].Write("p < 0.001");
                            }
                            else
                            {
                                wt.Rows[str + 1][j + 1].SetFont(settings.FontStandart);
                                wt.Rows[str + 1][j + 1].WriteLine("r = " + Math.Round(r, 3).ToString());
                                wt.Rows[str + 1][j + 1].Write("p = " + Math.Round(p, 3).ToString());
                            }
                            i++;
                        }
                    }

                    //Если вывели последний элемент в диагонали
                    if (str == wt.Rows.Length - 2)
                    {
                        dropCount = false;
                        //Считаем индекс диагонального блока
                        diag += diagPP;
                        diagPP++;
                        dropBlock++;
                    }
                    else
                    {
                        dropCount = true;
                    }
                    strPP = true;
                }
                else
                {
                    //Нет
                    for (int j = 0; j < MaxSize; j++)
                    {
                        //Настройка вывода ячейки
                        double p = _result[i].p;
                        double r = Math.Round(_result[i].r, 3);

                        if (p <= 0.05)
                        {
                            wt.Rows[str + 1][j + 1].SetFont(settings.FontStandart);
                            wt.Rows[str + 1][j + 1].WriteLine("r = " + Math.Round(r, 3).ToString());
                            if (p > 0.001)
                                wt.Rows[str + 1][j + 1].Write("p = " + Math.Round(p, 3).ToString());
                            else
                                wt.Rows[str + 1][j + 1].Write("p < 0.001");
                        }
                        else
                        {
                            wt.Rows[str + 1][j + 1].SetFont(settings.FontStandart);
                            wt.Rows[str + 1][j + 1].WriteLine("r = " + Math.Round(r, 3).ToString());
                            wt.Rows[str + 1][j + 1].Write("p = " + Math.Round(p, 3).ToString());
                        }
                        i++;
                    }

                    //Остаемся на этой же строке
                    strPP = false;
                    dropCount = false;
                }

                //Изменяем таблицу в списке
                tables.RemoveAt(count);
                tables.Insert(count, wt);

                if (dropCount == true)
                    count -= dropBlock;
                else
                    count++;

                if (strPP == true)
                    if (str == MaxSize - 1)
                        str = 0;
                    else
                        str++;
            }

            //Добавим отчет к таблицам
            _wordDocument = ReportByCorrelationAnalysis(_wordDocument, _result);
            return _wordDocument;
        }

        /// <summary>
        /// Отчет по-корреляционному анализу
        /// </summary>
        /// <param name="_wordDocument">Документ</param>
        /// <param name="_result">Список с результатами проверки</param>
        /// <returns></returns>
        private WordDocument ReportByCorrelationAnalysis(WordDocument _wordDocument, List<СorrelationResult> _result)
        {
            _wordDocument.WriteLine();
            _wordDocument.SetTextAlign(WordTextAlign.Justified);
            List<СorrelationResult> significiant = SignificiantAssociation(_result);
            List<СorrelationResult> insignificant = InsignificiantAssociation(_result);

            //Вывод информации о значимых связях
            //Наиболее значимая связь
            if (significiant.Count != 0)
            {
                _wordDocument.Write("С помощью корреляционного анализа было выявлено, что в исследуемых признаках наиболее значимая" +
                  " связь – связь между признаками " + (char)171 + data.TakeVariableNameAtIndex(significiant[0].thatCompare) + (char)187 +
                  " и " + (char)171 + data.TakeVariableNameAtIndex(significiant[0].toCompare) + (char)187 + " (r = " +
                  Math.Round(significiant[0].r, 3).ToString() + "). ");
                _wordDocument.Write("Корреляция имеет положительный знак, то есть при увеличение одного признака второй тоже увеличивается.");
            }
            else
            {
                _wordDocument.WriteLine("Значимых связей не обнаружено. ");
            }
            //Остальные значимые связи
            if (significiant.Count > 1)
            {
                _wordDocument.Write(" Так же, положительная корреляция была замечена у следующих признаков: ");
                for (int i = 1; i < significiant.Count; i++)
                {
                    _wordDocument.Write((char)171 + data.TakeVariableNameAtIndex(significiant[i].thatCompare) + (char)187 +
                        " и " + (char)171 + data.TakeVariableNameAtIndex(significiant[i].toCompare) + (char)187 +
                        " (r = " + Math.Round(significiant[i].r, 3).ToString() + ")");
                    if (i + 1 == significiant.Count)
                        _wordDocument.WriteLine(".");
                    else
                        _wordDocument.Write(", ");
                }
            }

            //Вывод информации о незначимых связях
            //Наименее значимая связь
            if (insignificant.Count != 0)
            {
                _wordDocument.Write("С помощью корреляционного анализа было выявлено, что в исследуемых признаках наименее значимая" +
                " связь – связь между признаками " + (char)171 + data.TakeVariableNameAtIndex(insignificant[0].thatCompare) + (char)187 +
                " и " + (char)171 + data.TakeVariableNameAtIndex(insignificant[0].toCompare) + (char)187 + " (r = " +
                Math.Round(insignificant[0].r, 3).ToString() + ").");
                _wordDocument.Write("Корреляция имеет отрицательный знак, то есть при увеличение одного признака второй уменьшается.");
            }
            else
            {
                _wordDocument.WriteLine("Незначимых связей не обнаружено. ");
            }
            //Остальные незначимые связи
            if (insignificant.Count > 1)
            {
                _wordDocument.Write(" Так же, отрицательная корреляция была замечена у следующих признаков: ");
                for (int i = 1; i < insignificant.Count; i++)
                {
                    _wordDocument.Write((char)171 + data.TakeVariableNameAtIndex(insignificant[i].thatCompare) + (char)187 +
                        " и " + (char)171 + data.TakeVariableNameAtIndex(insignificant[i].toCompare) + (char)187 +
                        " (r = " + Math.Round(insignificant[i].r, 3).ToString() + ")");
                    if (i + 1 == insignificant.Count)
                        _wordDocument.WriteLine(".");
                    else
                        _wordDocument.Write(", ");
                }
            }

            return _wordDocument;
        }

        /// <summary>
        /// Значимые свзяи
        /// </summary>
        /// <param name="_result"></param>
        /// <returns></returns>
        private List<СorrelationResult> SignificiantAssociation(List<СorrelationResult> _result)
        {
            //Составим список из связей, где гипотеза не отвергается и положительный коэффициент
            List<СorrelationResult> H = new List<СorrelationResult>();
            foreach (СorrelationResult h in _result)
            {
                if (h.p <= 0.05 && h.r >= 0)
                    H.Add(h);
            }
            if (H.Count != 0)
            {
                //Найдем максимальный элемент в списке H
                int index = 0;
                double max = H[0].r;
                for (int i = 1; i < H.Count; i++)
                {
                    if (H[i].r > max)
                        index = i;
                }
                //Поставим максимальный элемент на первое место
                СorrelationResult buf = H[0];
                H[0] = H[index];
                H[index] = buf;
            }
            return H;
        }

        /// <summary>
        /// Незначимые связи
        /// </summary>
        /// <param name="_result"></param>
        /// <returns></returns>
        private List<СorrelationResult> InsignificiantAssociation(List<СorrelationResult> _result)
        {
            //Составим список из связей, где гипотеза не отвергается и отрицатедльный коэффициент
            List<СorrelationResult> H = new List<СorrelationResult>();
            foreach (СorrelationResult h in _result)
            {
                if (h.p <= 0.05 && h.r < 0)
                    H.Add(h);
            }
            if (H.Count != 0)
            {
                //Найдем минимальный элемент в списке H
                int index = 0;
                double min = H[0].r;
                for (int i = 1; i < H.Count; i++)
                {
                    if (H[i].r < min)
                        index = i;
                }
                //Поставим минимальный элемент на первое место
                СorrelationResult buf = H[0];
                H[0] = H[index];
                H[index] = buf;
            }
            return H;
        }

        #endregion

        #region Таблицы сопряженности

        /// <summary>
        /// Выполнить анализ
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void metroButton25_Click(object sender, EventArgs e)
        {
            engine = REngine.GetInstance();
            ContingencyTableResult res = new ContingencyTableResult();
            //Получим список таблиц
            List<ContingencyTable> tables = ContingencyTable.GetTableList(data, groupFactList);
            foreach (ContingencyTable table in tables)
            {
                double stat;
                if (table.RowCount == 2 && table.ColumnCount == 2)
                    stat = ContingencyTable.Wulf(table);
                else
                    stat = ContingencyTable.Hi2Pearson(table);
                
                string str = "pv <- 1 - pchisq(" + stat.ToString().Replace(',','.') + ", " + table.Df.ToString() + ")";
                var pval = engine.Evaluate(str).AsNumeric();

                res.stat = stat;
                res.pvalue = pval[0];
                res.table = table;
                resContingencyTables.Add(res);
            }
            //NumericVector group1 = engine.CreateNumericVector(rescontingencyTables[0].Data.Count);
            //for (int i = 0; i < rescontingencyTables[0].Data.Count; i++)
            //    group1[i] = rescontingencyTables[0].Data[i];
            //engine.SetSymbol("Rgroup1", group1);
            //var tmp = engine.Evaluate("mice <- matrix(Rgroup1, nrow = " + rescontingencyTables[0].RowCount + ", ncol = " + rescontingencyTables[0].ColumnCount + ", byrow = TRUE)");
            //engine.Evaluate("mice");
            //var asd = engine.Evaluate("chisq.test(mice)");
            metroButton305.Visible = true;
        }

        /// <summary>
        /// Выведем таблицы сопряженности в файл
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void metroButton303_Click(object sender, EventArgs e)
        {
            WordDocument report = new WordDocument(WordDocumentFormat.InCentimeters(21, 29.7, 2.5, 1, 2, 2));

            report.SetFont(settings.FontStandart);
            report.SetTextAlign(WordTextAlign.Center);
            report.WriteLine("Таблицы сопряженности ");

            report = OutResultInContingencyTable(report, resContingencyTables);

            report.SaveToFile("..\\..\\Table.doc");
            System.Diagnostics.Process.Start("..\\..\\Table.doc");
        }

        /// <summary>
        /// Вывод информации в таблицы
        /// </summary>
        /// <param name="_wordDocument">Документ</param>
        /// <param name="_methodName">Название метода</param>
        /// <param name="_resList">Список с результатами проверки</param>
        /// <returns></returns>
        private WordDocument OutResultInContingencyTable(WordDocument _wordDocument, List<ContingencyTableResult> _resList)
        {
            //Идем по спискам тиблиц
            foreach(ContingencyTableResult table in _resList)
            {
                WordTable rt1 = _wordDocument.NewTable(settings.FontStandart, Color.Black, table.table.Variable1List.Count + 2, table.table.Variable2List.Count + 2, 2);
                _wordDocument.SetTextAlign(WordTextAlign.Left);
                _wordDocument.SetFont(settings.FontStandart);
                _wordDocument.SetTextAlign(WordTextAlign.Left);
                if (table.table.RowCount == 2 && table.table.ColumnCount == 2)
                    _wordDocument.WriteLine("Критерий Вулфа");
                else
                    _wordDocument.WriteLine("Критерий Хи-квадрат Пирсона");
                
                //Выведем список уникальных значений первой переменной. Вывод по строке
                for (int j = 0; j < table.table.Variable1List.Count; j++)
                    rt1.Rows[j + 2][1].Write(table.table.Variable1List[j]);
                //Выведем список уникальных значений второй переменной. Вывод по столбцу
                for (int j = 0; j < table.table.Variable2List.Count; j++)
                    rt1.Rows[1][j + 2].Write(table.table.Variable2List[j]);

                //Объединение ячеек и заполнение их данными
                rt1.Rows[0][0].ColSpan = 2;
                rt1.Rows[1][0].ColSpan = 2;
                rt1.Rows[0][0].RowSpan = 2;
                rt1.Rows[0][0].SetBorders(Color.Black, 1, true, true, true, true);

                rt1.Rows[2][0].RowSpan = table.table.Variable1List.Count;
                rt1.Rows[2][0].Write(table.table.Variable1);
                rt1.Rows[2][0].SetBorders(Color.Black, 1, true, true, true, true);

                rt1.Rows[0][2].ColSpan = table.table.Variable2List.Count;
                rt1.Rows[0][2].Write(table.table.Variable2);
                rt1.Rows[0][2].SetBorders(Color.Black, 1, true, true, true, true);

                //Выведем данные в таблицу
                for (int j = 0; j < table.table.Variable1List.Count; j++)
                    //Идем по столбцам
                    for (int k = 0; k < table.table.Variable2List.Count; k++)
                        rt1.Rows[j + 2][k + 2].Write(table.table.Data[k + j * table.table.Variable2List.Count].ToString());

                //Нарисуем рамки у ячеек
                for (int j = 0; j < table.table.Variable1List.Count + 2; j++)
                    for (int k = 0; k < table.table.Variable2List.Count + 2; k++)
                        rt1.Rows[j][k].SetBorders(Color.Black, 1, true, true, true, true);

                rt1.SaveToDocument(9600, 0);
                _wordDocument.WriteLine();
                _wordDocument.WriteLine("Stat = " + table.stat.ToString());
                _wordDocument.WriteLine("Pval = " + table.pvalue.ToString());
                _wordDocument.WriteLine();
            }
            return _wordDocument;
        }

        #endregion

        /// <summary>
        /// Изменение размера вкладки "Сравнение независимых групп"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void metroTabPage2_SizeChanged(object sender, EventArgs e)
        {
            /***************************Внешний вид кладки***************************/
            //GroupBox'ы
            //Группирующие факторы
            Size size = groupBox1.Size;
            size.Width = (metroTabPage2.Size.Width - 10) / 2;
            groupBox1.Size = size;
            //Бинарные переменные
            size = groupBox2.Size;
            size.Width = (groupBox1.Size.Width / 2) - 10;
            groupBox2.Size = size;
            //Номинальные переменные
            groupBox3.Location = new Point(size.Width + settings.PaddingBetweenGroupBoxes, groupBox3.Location.Y);
            groupBox3.Size = new Size(size.Width, groupBox3.Size.Height);

            //Признаки
            size = groupBox1.Size;
            groupBox4.Location = new Point(size.Width + settings.PaddingBetweenGroupBoxes, groupBox4.Location.Y);
            groupBox4.Size = size;
            //Количественные переменные
            size = groupBox5.Size;
            size.Width = (groupBox4.Size.Width / 2) - 10;
            groupBox5.Size = size;
            //Порядковые переменные
            groupBox6.Location = new Point(size.Width + settings.PaddingBetweenGroupBoxes, groupBox6.Location.Y);
            groupBox6.Size = new Size(size.Width, groupBox6.Size.Height);

            //Кнопки
            //Подвинем кнопки под groupBox'ами
            //Изменение положения кнопок бинарных переменных
            int padding = settings.PaddingFromBtnSelectAll;
            Point grpLoc = groupBox2.Location;
            metroButton1.Location = new Point(grpLoc.X + padding, groupBox2.Height + grpLoc.Y + settings.StandartPadding * 2);
            metroButton2.Location = new Point(grpLoc.X + groupBox2.Width - metroButton2.Width - padding, groupBox2.Height + grpLoc.Y + settings.StandartPadding * 2);

            //Изменение положения кнопок номинальных переменных
            grpLoc = groupBox3.Location;
            metroButton3.Location = new Point(grpLoc.X + padding, groupBox3.Height + grpLoc.Y + settings.StandartPadding * 2);
            metroButton4.Location = new Point(grpLoc.X + groupBox3.Width - metroButton3.Width - padding, groupBox3.Height + grpLoc.Y + settings.StandartPadding * 2);

            //Изменение положения кнопок количественных переменных
            grpLoc = groupBox5.Location;
            metroButton5.Location = new Point(grpLoc.X + padding, groupBox5.Height + grpLoc.Y + settings.StandartPadding * 2);
            metroButton6.Location = new Point(grpLoc.X + groupBox5.Width - metroButton6.Width - padding, groupBox5.Height + grpLoc.Y + settings.StandartPadding * 2);

            //Изменение положения кнопок порядковых переменных  
            grpLoc = groupBox6.Location;
            metroButton7.Location = new Point(grpLoc.X + padding, groupBox6.Height + grpLoc.Y + settings.StandartPadding * 2);
            metroButton8.Location = new Point(grpLoc.X + groupBox6.Width - metroButton8.Width - padding, groupBox6.Height + grpLoc.Y + settings.StandartPadding * 2);

            //Выровняем кнопки "Выполнить сравнение" и "Вывести в Word"
            metroButton9.Location = new Point(groupBox1.Location.X + groupBox1.Size.Width - metroButton9.Width, groupBox1.Size.Height + groupBox1.Location.Y + settings.StandartPadding * 2);
            metroButton10.Location = new Point(groupBox4.Location.X, groupBox4.Size.Height + groupBox4.Location.Y + settings.StandartPadding * 2);
        }

        /// <summary>
        /// Отрисовка графиков по полученным выборкам
        /// </summary>
        /// <param name="_samples">Список сгруппированных выборок</param>
        private void CreateGraphic(Sample _sample, string _Methodname)
        {
            //Создадим дирректорию для хранения графиков
            if (!Directory.Exists(pathC))
                Directory.CreateDirectory(pathC); 

            string data = "";
            string names = "";
            for (int i = 0; i < _sample.SubSampleList.Count; i++)
            {
                //Создать вектор
                string name =  "group" + i.ToString();
                NumericVector group = engine.CreateNumericVector(_sample.SubSampleList[i].SampleList);
                //Перевести его в R
                engine.SetSymbol(name, group);
                if (i != _sample.SubSampleList.Count - 1)
                    data += name + ",";
                else
                    data += name;

                if (i != _sample.SubSampleList.Count - 1)
                    names += "\"" + _sample.SubSampleList[i].UniqueVal.ToString() + "\"" + ",";
                else
                    names += "\"" + _sample.SubSampleList[i].UniqueVal.ToString() + "\"";
            }
            engine.Evaluate("jpeg(\"" + pathR + _Methodname + "_" +_sample.GroupFact + "_" + _sample.NameSign + ".jpg\")");
            engine.Evaluate("boxplot(" + data + ", main=\"" + "Диаграмма размаха" + "\", names=c(" + names + "), ylab=\"" + _sample.NameSign.ToString()+"\", xlab=\"" + _sample.GroupFact.ToString() + "\")");
            engine.Evaluate("dev.off()");
        }

        private void metroTabPage3_SizeChanged(object sender, EventArgs e)
        {
            //grpBoxes 7 -> [8, 9]
            //Bth 100 - 104 ;   105, 106
            //970 -> 480 -> 230

            //Признаки
            //Контейнер для признаков
            Size size = groupBox7.Size;
            size.Width = (metroTabPage3.Size.Width - 10) / 2;
            groupBox7.Size = size;
            //Количественные переменные
            size = groupBox8.Size;
            size.Width = (groupBox7.Size.Width / 2) - 10;
            groupBox8.Size = size;
            //Порядковые переменные
            groupBox9.Location = new Point(size.Width + settings.PaddingBetweenGroupBoxes, groupBox9.Location.Y);
            groupBox9.Size = new Size(size.Width, groupBox9.Size.Height);

            //Кнопки
            int padding = settings.PaddingFromBtnSelectAll;
            Point grpLoc = groupBox8.Location;
            metroButton100.Location = new Point(grpLoc.X + padding, groupBox8.Height + grpLoc.Y + settings.StandartPadding * 2);
            metroButton101.Location = new Point(grpLoc.X + groupBox8.Width - metroButton101.Width - padding, groupBox8.Height + grpLoc.Y + settings.StandartPadding * 2);
            grpLoc = groupBox9.Location;
            metroButton102.Location = new Point(grpLoc.X + padding, groupBox9.Height + grpLoc.Y + settings.StandartPadding * 2);
            metroButton103.Location = new Point(grpLoc.X + groupBox9.Width - metroButton103.Width - padding, groupBox9.Height + grpLoc.Y + settings.StandartPadding * 2);
        }

        private void metroTabPage4_SizeChanged(object sender, EventArgs e)
        {
            //grpBoxes 10 -> [11, 12]
            //Bth 200 - 204 ;   205, 206
            //970 -> 480 -> 230

            //Признаки
            //Контейнер для признаков
            Size size = groupBox10.Size;
            size.Width = (metroTabPage4.Size.Width - 10) / 2;
            groupBox10.Size = size;
            //Количественные переменные
            size = groupBox11.Size;
            size.Width = (groupBox10.Size.Width / 2) - 10;
            groupBox11.Size = size;
            //Порядковые переменные
            groupBox12.Location = new Point(size.Width + settings.PaddingBetweenGroupBoxes, groupBox12.Location.Y);
            groupBox12.Size = new Size(size.Width, groupBox12.Size.Height);

            //Кнопки
            int padding = settings.PaddingFromBtnSelectAll;
            Point grpLoc = groupBox11.Location;
            metroButton200.Location = new Point(grpLoc.X + padding, groupBox11.Height + grpLoc.Y + settings.StandartPadding * 2);
            metroButton201.Location = new Point(grpLoc.X + groupBox11.Width - metroButton201.Width - padding, groupBox11.Height + grpLoc.Y + settings.StandartPadding * 2);
            grpLoc = groupBox12.Location;
            metroButton202.Location = new Point(grpLoc.X + padding, groupBox12.Height + grpLoc.Y + settings.StandartPadding * 2);
            metroButton203.Location = new Point(grpLoc.X + groupBox12.Width - metroButton203.Width - padding, groupBox12.Height + grpLoc.Y + settings.StandartPadding * 2);
        }

        private void metroTabPage5_SizeChanged(object sender, EventArgs e)
        {
            //grpBoxes 13 -> [14, 15]
            //10 11 12
            //Bth 300 - 304 ;   305, 306
            //970 -> 480 -> 230

            //Признаки
            //Контейнер для признаков
            Size size = groupBox13.Size;
            size.Width = (metroTabPage5.Size.Width - 10) / 2;
            groupBox13.Size = size;
            //Номинальные переменные
            size = groupBox14.Size;
            size.Width = (groupBox13.Size.Width / 2) - 10;
            groupBox14.Size = size;
            //Бинарные переменные
            groupBox15.Location = new Point(size.Width + settings.PaddingBetweenGroupBoxes, groupBox15.Location.Y);
            groupBox15.Size = new Size(size.Width, groupBox15.Size.Height);

            //Кнопки
            int padding = settings.PaddingFromBtnSelectAll;
            Point grpLoc = groupBox14.Location;
            metroButton300.Location = new Point(grpLoc.X + padding, groupBox14.Height + grpLoc.Y + settings.StandartPadding * 2);
            metroButton301.Location = new Point(grpLoc.X + groupBox14.Width - metroButton301.Width - padding, groupBox14.Height + grpLoc.Y + settings.StandartPadding * 2);
            grpLoc = groupBox15.Location;
            metroButton302.Location = new Point(grpLoc.X + padding, groupBox15.Height + grpLoc.Y + settings.StandartPadding * 2);
            metroButton303.Location = new Point(grpLoc.X + groupBox15.Width - metroButton303.Width - padding, groupBox15.Height + grpLoc.Y + settings.StandartPadding * 2);
        }
    }
}