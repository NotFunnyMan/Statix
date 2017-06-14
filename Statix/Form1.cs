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
        private Sample resDepWilcoxon;

        /// <summary>
        /// Сравнение зависимых групп. Тест Фридмана
        /// </summary>
        private Sample resDepFridman;

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

        /// <summary>
        /// Дирректории для C# и R
        /// </summary>
        private string pathC = "..\\..\\graphics";
        private string pathR = "../../graphics/";

        /// <summary>
        /// Сквозная нумерация для отчетов
        /// </summary>
        private int endToEndTable;
        private int endToEndPicture;

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
        private void btnLoadData_Click(object sender, EventArgs e)
        {
            //Считывание данных из файла
            //Далее будет выбор файла
            data = new Data("ExampleCSV.csv", Encoding.Default);

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
            //        Вывод списка ошибок в новом окне
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
            GrdData.ColumnCount = records;
            GrdData.RowCount = data.PatientsCount + 3;

            //Строка "Уникальные значения"
            GrdData.Rows[0].DefaultCellStyle.BackColor = settings.ColorTableCaption;
            for (int i = 0; i < records; i++)
                GrdData.Rows[0].Cells[i].Value = data.TakeUniqueValuesAtIndex(i);

            //Строка "Тип шкалы"
            GrdData.Rows[1].DefaultCellStyle.BackColor = settings.ColorTableCaption;
            for (int i = 0; i < records; i++)
                GrdData.Rows[1].Cells[i].Value = data.TakeScaleTypeAtIndex(i);

            //Строка "Название переменной"
            GrdData.Rows[2].DefaultCellStyle.BackColor = settings.ColorTableCaption;
            for (int i = 0; i < records; i++)
                GrdData.Rows[2].Cells[i].Value = data.TakeVariableNameAtIndex(i);

            for (int j = 0; j < data.PatientsCount; j++)
            {
                List<string> person = data.TakePatientAtIndex(j);
                //Цвет строки в таблице
                if (person.IndexOf("NA") != -1)
                    GrdData.Rows[j + 3].DefaultCellStyle.BackColor = settings.ColorTableError;
                else
                    GrdData.Rows[j + 3].DefaultCellStyle.BackColor = settings.ColorTableOK;

                for (int i = 0; i < records; i++)
                    GrdData.Rows[j + 3].Cells[i].Value = person[i];
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
                grpBIndepBin.Controls.Clear();
                grpBIndepNom.Controls.Clear();
                grpBIndepCol.Controls.Clear();
                grpBIndepPor.Controls.Clear();
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
                    grpBIndepBin.Controls.Add(rB[i]);
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
                    grpBIndepNom.Controls.Add(rB[i]);
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
                    grpBIndepCol.Controls.Add(rB[i]);
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
                    grpBIndepPor.Controls.Add(rB[i]);
                }
                
                //Скроем кнопки с других вкладок
                //Кнопки выполнения анализа
                btnDepExecute.Visible = false;
                btnCorrelExecute.Visible = false;
                btnTableExecute.Visible = false;

                //Кнопки вывод в Word
                btnIndepReport.Visible = false;
                btnDepReport.Visible = false;
                btnCorrelReport.Visible = false;
                btnTableReport.Visible = false;

                //Скроем вкладку с признаками. Это для корректного отображения при возврате на данную вкладку
                grpBIndepSign.Visible = false;

                //Настроим высоту groupBox'ов, которые хранят в себе CheckBox'ы
                int height = grpBIndepBin.Size.Height;
                int width = grpBIndepBin.Size.Width;
                if (grpBIndepBin.Controls.Count * settings.PaddingBetweenCheckBoxes > height)
                    height = grpBIndepBin.Controls.Count * settings.PaddingBetweenCheckBoxes;
                grpBIndepBin.Size = new Size(width, height);

                height = grpBIndepNom.Size.Height;
                width = grpBIndepNom.Size.Width;
                if (grpBIndepNom.Controls.Count * settings.PaddingBetweenCheckBoxes > height)
                    height = grpBIndepNom.Controls.Count * settings.PaddingBetweenCheckBoxes;
                grpBIndepNom.Size = new Size(width, height);

                height = grpBIndepCol.Size.Height;
                width = grpBIndepCol.Size.Width;
                if (grpBIndepCol.Controls.Count * settings.PaddingBetweenCheckBoxes > height)
                    height = grpBIndepCol.Controls.Count * settings.PaddingBetweenCheckBoxes;
                grpBIndepCol.Size = new Size(width, height);

                height = grpBIndepPor.Size.Height;
                width = grpBIndepPor.Size.Width;
                if (grpBIndepPor.Controls.Count * settings.PaddingBetweenCheckBoxes > height)
                    height = grpBIndepPor.Controls.Count * settings.PaddingBetweenCheckBoxes;
                grpBIndepPor.Size = new Size(width, height);

                //Выровним groupBox'ы для Груп. факторов и признаков
                //64 - для отображения кнопок под grpBox. Иначе их не видно! ПОФИКСИТЬ!
                if (grpBIndepBin.Size.Height > grpBIndepNom.Size.Height)
                    grpBIndepGroupFact.Size = new Size(grpBIndepGroupFact.Size.Width, grpBIndepBin.Size.Height + 64);
                if (grpBIndepCol.Size.Height > grpBIndepPor.Size.Height)
                    grpBIndepSign.Size = new Size(grpBIndepSign.Size.Width, grpBIndepCol.Size.Height + 64);

                if (grpBIndepGroupFact.Height > grpBIndepSign.Height)
                    grpBIndepSign.Height = grpBIndepGroupFact.Height;
                else
                    grpBIndepGroupFact.Height = grpBIndepSign.Height;

                //Настроим внешний вид страницы
                metroTabPage2_SizeChanged(sender, e);
            }
            #endregion

            #region "Сравнение зависимых групп"

            if (e.TabPageIndex == 2)
            {
                grpBDepCol.Controls.Clear();
                grpBDepPor.Controls.Clear();
                MetroCheckBox[] rB;

                //Скроем кнопки с других вкладок
                btnIndepExecute.Visible = false;
                btnCorrelExecute.Visible = false;
                btnTableExecute.Visible = false;

                //Кнопки вывод в Word
                btnIndepReport.Visible = false;
                btnDepReport.Visible = false;
                btnCorrelReport.Visible = false;
                btnTableReport.Visible = false;

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
                    grpBDepCol.Controls.Add(rB[i]);
                }

                //Изменение положения кнопок
                Point point = new Point(btnDepColSelectAll.Location.X, btnDepColSelectAll.Location.Y);
                point.Y = grpBDepCol.Size.Height + 25;
                btnDepColSelectAll.Location = point;
                point = new Point(btnDepColCleanAll.Location.X, btnDepColCleanAll.Location.Y);
                point.Y = grpBDepCol.Size.Height + 25;
                btnDepColCleanAll.Location = point;

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
                    grpBDepPor.Controls.Add(rB[i]);
                }

                //Изменение положения кнопок
                point = new Point(btnDepPorSelectAll.Location.X, btnDepPorSelectAll.Location.Y);
                point.Y = grpBDepPor.Size.Height + 25;
                btnDepPorSelectAll.Location = point;
                point = new Point(btnDepPorCleanAll.Location.X, btnDepPorCleanAll.Location.Y);
                point.Y = grpBDepPor.Size.Height + 25;
                btnDepPorCleanAll.Location = point;

                //Настроим высоту groupBox'ов, которые хранят в себе CheckBox'ы
                int height = grpBDepCol.Size.Height;
                int width = grpBDepCol.Size.Width;
                if (grpBDepCol.Controls.Count * settings.PaddingBetweenCheckBoxes > height)
                    height = grpBDepCol.Controls.Count * settings.PaddingBetweenCheckBoxes;
                grpBDepCol.Size = new Size(width, height);

                height = grpBDepPor.Size.Height;
                width = grpBDepPor.Size.Width;
                if (grpBDepPor.Controls.Count * settings.PaddingBetweenCheckBoxes > height)
                    height = grpBDepPor.Controls.Count * settings.PaddingBetweenCheckBoxes;
                grpBDepPor.Size = new Size(width, height);


                if (grpBDepCol.Size.Height > grpBDepPor.Size.Height)
                    grpBDepSign.Size = new Size(grpBDepSign.Size.Width, grpBDepCol.Size.Height + 64);
                else
                    grpBDepSign.Size = new Size(grpBDepSign.Size.Width, grpBDepPor.Size.Height + 64);

                //Настроим внешний вид страницы
                metroTabPage3_SizeChanged(sender, e);
            }
            #endregion

            #region Корреляционный анализ

            if (e.TabPageIndex == 3)
            {
                grpBCorrelCol.Controls.Clear();
                grpBCorrelPor.Controls.Clear();
                MetroCheckBox[] rB;

                //Скроем кнопки с других вкладок
                btnIndepExecute.Visible = false;
                btnDepExecute.Visible = false;
                btnTableExecute.Visible = false;

                //Кнопки вывод в Word
                btnIndepReport.Visible = false;
                btnDepReport.Visible = false;
                btnCorrelReport.Visible = false;
                btnTableReport.Visible = false;

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
                    grpBCorrelCol.Controls.Add(rB[i]);
                }

                //Изменение положения кнопок
                Point point = new Point(btnCorrelColSelectAll.Location.X, btnCorrelColSelectAll.Location.Y);
                point.Y = grpBCorrelCol.Size.Height + 25;
                btnCorrelColSelectAll.Location = point;
                point = new Point(btnCorrelColCleanAll.Location.X, btnCorrelColCleanAll.Location.Y);
                point.Y = grpBCorrelCol.Size.Height + 25;
                btnCorrelColCleanAll.Location = point;

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
                    grpBCorrelPor.Controls.Add(rB[i]);
                }

                //Изменение положения кнопок
                point = new Point(btnCorrelPorSelectAll.Location.X, btnCorrelPorSelectAll.Location.Y);
                point.Y = grpBCorrelPor.Size.Height + 25;
                btnCorrelPorSelectAll.Location = point;
                point = new Point(btnCorrelPorCleanAll.Location.X, btnCorrelPorCleanAll.Location.Y);
                point.Y = grpBCorrelPor.Size.Height + 25;
                btnCorrelPorCleanAll.Location = point;

                //Настроим высоту groupBox'ов, которые хранят в себе CheckBox'ы
                int height = grpBCorrelCol.Size.Height;
                int width = grpBCorrelCol.Size.Width;
                if (grpBCorrelCol.Controls.Count * settings.PaddingBetweenCheckBoxes > height)
                    height = grpBCorrelCol.Controls.Count * settings.PaddingBetweenCheckBoxes;
                grpBCorrelCol.Size = new Size(width, height);

                height = grpBCorrelPor.Size.Height;
                width = grpBCorrelPor.Size.Width;
                if (grpBCorrelPor.Controls.Count * settings.PaddingBetweenCheckBoxes > height)
                    height = grpBCorrelPor.Controls.Count * settings.PaddingBetweenCheckBoxes;
                grpBCorrelPor.Size = new Size(width, height);

                if (grpBCorrelCol.Size.Height > grpBCorrelPor.Size.Height)
                    grpBCorrelSign.Size = new Size(grpBCorrelSign.Size.Width, grpBCorrelCol.Size.Height + 64);
                else
                    grpBCorrelSign.Size = new Size(grpBCorrelSign.Size.Width, grpBCorrelPor.Size.Height + 64);

                //Настроим внешний вид страницы
                metroTabPage4_SizeChanged(sender, e);
            }

            #endregion

            #region Таблицы сопряженности

            if (e.TabPageIndex == 4)
            {
                grpBTableNom.Controls.Clear();
                grpBTableBin.Controls.Clear();
                MetroCheckBox[] rB;

                //Скроем кнопки с других вкладок
                btnIndepExecute.Visible = false;
                btnDepExecute.Visible = false;
                btnCorrelExecute.Visible = false;

                //Кнопки вывод в Word
                btnIndepReport.Visible = false;
                btnDepReport.Visible = false;
                btnCorrelReport.Visible = false;
                btnTableReport.Visible = false;

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
                    grpBTableNom.Controls.Add(rB[i]);
                }

                //Изменение положения кнопок
                Point point = new Point(btnTableNomSelectAll.Location.X, btnTableNomSelectAll.Location.Y);
                point.Y = grpBTableNom.Size.Height + 25;
                btnTableNomSelectAll.Location = point;
                point = new Point(btnTableNomCleanAll.Location.X, btnTableNomCleanAll.Location.Y);
                point.Y = grpBTableNom.Size.Height + 25;
                btnTableNomCleanAll.Location = point;

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
                    grpBTableBin.Controls.Add(rB[i]);
                }

                //Изменение положения кнопок
                point = new Point(btnTableBinSelectAll.Location.X, btnTableBinSelectAll.Location.Y);
                point.Y = grpBTableBin.Size.Height + 25;
                btnTableBinSelectAll.Location = point;
                point = new Point(btnTableBinCleanAll.Location.X, btnTableBinCleanAll.Location.Y);
                point.Y = grpBTableBin.Size.Height + 25;
                btnTableBinCleanAll.Location = point;

                //Настроим высоту groupBox'ов, которые хранят в себе CheckBox'ы
                int height = grpBTableNom.Size.Height;
                int width = grpBTableNom.Size.Width;
                if (grpBTableNom.Controls.Count * settings.PaddingBetweenCheckBoxes > height)
                    height = grpBTableNom.Controls.Count * settings.PaddingBetweenCheckBoxes;
                grpBTableNom.Size = new Size(width, height);

                height = grpBTableBin.Size.Height;
                width = grpBTableBin.Size.Width;
                if (grpBTableBin.Controls.Count * settings.PaddingBetweenCheckBoxes > height)
                    height = grpBTableBin.Controls.Count * settings.PaddingBetweenCheckBoxes;
                grpBTableBin.Size = new Size(width, height);

                if (grpBTableNom.Size.Height > grpBTableBin.Size.Height)
                    grpBTableSign.Size = new Size(grpBTableSign.Size.Width, grpBTableNom.Size.Height + 64);
                else
                    grpBTableSign.Size = new Size(grpBTableSign.Size.Width, grpBTableBin.Size.Height + 64);

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
                    grpBIndepSign.Visible = true;
                }
                if (Convert.ToInt32(mCB.Parent.Tag) == 4 && groupFactList.Count >= 2)
                    btnTableExecute.Visible = true;
            }
            else
            {
                groupFactList.Remove(index);
                if (groupFactList.Count == 0)
                    grpBIndepSign.Visible = false;
                if (Convert.ToInt32(mCB.Parent.Tag) == 4 && groupFactList.Count < 2)
                    btnTableExecute.Visible = false;
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
                    case 1: btnIndepExecute.Visible = true; break;
                    case 2: btnDepExecute.Visible = true; break;
                    case 3: btnCorrelExecute.Visible = true; break;
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
                        case 1: btnIndepExecute.Visible = false; btnIndepReport.Visible = false; break;
                        case 2: btnDepExecute.Visible = false; btnDepReport.Visible = false; break;
                        case 3: btnCorrelExecute.Visible = false; btnCorrelReport.Visible = false; break;
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
                case 1: Check(grpBIndepBin); break;
                //Номинальные переменные
                case 3: Check(grpBIndepNom); break;
                //Количественные переменные
                case 5: Check(grpBIndepCol); break;
                //Порядковые переменные
                case 7: Check(grpBIndepPor); break;

                //"Сравнение зависимых групп"
                //Количественные переменные
                case 9: Check(grpBDepCol); break;
                //Порядковые переменные
                case 11: Check(grpBDepPor); break;

                //"Корреляционный анализ"
                //Количественные переменные
                case 13: Check(grpBCorrelCol); break;
                //Порядковые переменные
                case 15: Check(grpBCorrelPor); break;

                //"Таблицы сопряженности"
                //Номинальные переменные
                case 17: Check(grpBTableNom); break;
                //Бинарные переменные
                case 19: Check(grpBTableBin); break;
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
                case 2: Uncheck(grpBIndepBin); break;
                //Номинальные переменные
                case 4: Uncheck(grpBIndepNom); break;
                //Количественные переменные
                case 6: Uncheck(grpBIndepCol); break;
                //Порядковые переменные
                case 8: Uncheck(grpBIndepPor); break;
                    
                //"Сравнение зависимых групп"
                //Количественные переменные
                case 10: Uncheck(grpBDepCol); break;
                //Порядковые переменные
                case 12: Uncheck(grpBDepPor); break;

                //"Корреляционный анализ"
                //Количественные переменные
                case 14: Uncheck(grpBCorrelCol); break;
                //Порядковые переменные
                case 16: Uncheck(grpBCorrelPor); break;

                //"Таблицы сопряженности"
                //Номинальные переменные
                case 18: Uncheck(grpBTableNom); break;
                //Бинарные переменные
                case 20: Uncheck(grpBTableBin); break;
            }

            //Если сняли все галочки на группирующих переменных, то
            //Очистим список выбранных признаков, скроем поле с их выбором и снимем галочки
            if (groupFactList.Count == 0)
            {
                Uncheck(grpBIndepCol);
                Uncheck(grpBIndepPor);
                grpBIndepSign.Visible = false;
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
        private void btnIndepExecute_Click(object sender, EventArgs e)
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

                        //Если получилась статистически значимая связь, то нарисуем график
                        if (p <= settings.Statistical_significance)
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

                            //StreamWriter sw = new StreamWriter("out" + j.ToString() + ".txt");
                            //string buf = "";
                            //for (int a = 0; a < nV.Length; a++)
                            //    buf += nV[a].ToString() + Environment.NewLine;
                            //sw.WriteLine(buf);
                            //sw.Close();
                        }

                        engine.SetSymbol("sample", gV);
                        tmp = engine.Evaluate("kruskal.test(sample)").AsList();
                        double p = tmp["p.value"].AsNumeric().First();
                        testRes.PValue = p;

                        //Посчитаем медиану и стандартное отклонение
                        testRes = FillingResults(samples[i].SubSampleList, testRes);
                        resIndKruskalWallis.Add(testRes);

                        //Если получилась статистически значимая связь, то нарисуем график
                        if (p <= settings.Statistical_significance)
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
            btnIndepReport.Visible = true;
        }

        /// <summary>
        /// Вывод результатов проверки независимых выборок в Word
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnIndepReport_Click(object sender, EventArgs e)
        {
            WordDocument report = new WordDocument(WordDocumentFormat.InCentimeters(21, 29.7, 2.5, 1, 2, 2));
            //1.5 отступ в документе
            report.WriteControlWord(@"sl360\slmult1");
            report.SetFont(settings.FontStandart);
            report.SetTextAlign(WordTextAlign.Center);
            report.WriteLine("Сравнение средних по 2 независимым выборкам ");
            //Сквозная нумерация
            endToEndTable = 1;
            endToEndPicture = 1;
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

            //Получим список названий графиков
            List<string> graphNames = new List<string>();
            for (int j = 0; j < dirs.Length; j++)
            {
                string name = Path.GetFileNameWithoutExtension(dirs[j]);
                graphNames.Add(name);
            }

            var grpRes = Grouping(_resList);

            //Вывод результата
            for (int i = 0; i < grpRes.Count; i++)
            {
                WordTable rt1;
                Sample s;

                //Текст перед таблицей
                string grpFact = grpRes[i][0].GroupFact;
                string text = "В таблице " + endToEndTable.ToString() + " приведен результат статистического анализа данных, проведенного с использованием " 
                              + "критерия " + _methodName + "." + " В качестве группирующего фактора используется переменная " + "\"" + grpFact + "\"" + ".";
                _wordDocument.SetTextAlign(WordTextAlign.Justified);
                _wordDocument.SetParagraph(0, 567);
                _wordDocument.WriteLine(text);

                //Название таблицы
                string tableNumber = Environment.NewLine + "Таблица " + endToEndTable.ToString() + " - ";  
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
                string conclusion = "Из таблицы " + endToEndTable.ToString() + " видно, что ";
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
                            star += "\"" + grpRes[i][0].SubSampleList[k].UniqueVal + "\"" + ", ";
                        star += " и " + "\"" + grpRes[i][0].SubSampleList.Last().UniqueVal + "\"";
                    }
                    star += " p < 0.05.";
                    _wordDocument.WriteLine(note + star);
                    
                    //Вывод из таблицы
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
                    conclusion += " с учетом уровня значимости равного " + settings.Statistical_significance.ToString() + ".";
                    _wordDocument.SetParagraph(0, 567);
                    _wordDocument.WriteLine(conclusion);
                }
                else
                {
                    _wordDocument.WriteLine(note + star);
                    conclusion += "статистически значимых различий не обнаружено. Заданный уровень значимости равен " + settings.Statistical_significance.ToString() + ".";
                    _wordDocument.SetParagraph(0, 567);
                    _wordDocument.WriteLine(conclusion);
                }
                pval = false;

                //Есть ли графики к результату статистического анализа
                if (statSignif.Count != 0)
                {
                    //Добавляем графики
                    for (int j = 0; j < statSignif.Count; j++)
                    {
                        int curGraph = graphNames.IndexOf(methodName + "_" + grpRes[i][j].GroupFact + "_" + statSignif[j]);
                        if (curGraph != -1)
                        {
                            //Предисловие к графикам
                            string preface = "";
                            if (dirs.Length == 1)
                            {
                                preface += "На рисунке " + endToEndPicture.ToString() + " изображена диаграмма размахов признака, приведенного в таблице "
                                           + endToEndTable.ToString() + ".";
                            }
                            else
                            {
                                preface += "На рисунках " + endToEndPicture.ToString() + " - " + (endToEndPicture + dirs.Length).ToString() +
                                           " изображены диаграммы размахов признаков, приведенных в таблице " + endToEndTable.ToString() + ".";
                            }
                            _wordDocument.WriteLine(preface);

                            //Вставим график
                            _wordDocument.SetTextAlign(WordTextAlign.Center);
                            _wordDocument.PutImage(dirs[curGraph], 96); //96 - истинный dpi
                            _wordDocument.WriteLine();
                            //Подпись к графику
                            _wordDocument.WriteControlWord(@"sl360\slmult1");
                            note = "Рисунок " + endToEndPicture.ToString() + " - " +
                                    "Диаграмма размаха переменной " + "\"" + statSignif[j] + "\"";
                            _wordDocument.WriteLine(note);
                        }
                    }
                }
                endToEndTable++;
                endToEndPicture += statSignif.Count;
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
        private void btnDepExecute_Click(object sender, EventArgs e)
        {
            engine = REngine.GetInstance();
            GenericVector tmp;
            Sample sample = new Sample();
            sample = Sample.GetSample(data, signsList);
            
            if (signsList.Count == 2)
            {
                //Вилкоксон
                NumericVector group1 = engine.CreateNumericVector(sample.SubSampleList[0].SampleList);
                engine.SetSymbol("Rgroup1", group1);
                NumericVector group2 = engine.CreateNumericVector(sample.SubSampleList[1].SampleList);
                engine.SetSymbol("Rgroup2", group2);
                tmp = engine.Evaluate("wilcox.test(Rgroup1, Rgroup2, paired=TRUE)").AsList();
                double p = tmp["p.value"].AsNumeric().First();
                sample.PValue = p;

                //Посчитаем медиану и стандартное отклонение
                sample = FillingResults(sample.SubSampleList, sample);
                resDepWilcoxon = sample;

                //Если получилась статистически значимая связь, то нарисуем график
                if (p <= settings.Statistical_significance)
                    //Нарисовать график для выборки
                    CreateGraphic(sample.SubSampleList, "WI");
            }
            else
            {
                //Фридман
                //Создадим список выборок для отправки в тест Фридмана
                double[] list = new double[sample.SubSampleList.Count * sample.SubSampleList[0].SampleList.Count];
                for (int j = 0; j < sample.SubSampleList.Count; j++)
                    Array.Copy(sample.SubSampleList[j].SampleList.ToArray(), 0, list, j * sample.SubSampleList[j].SampleList.Count, sample.SubSampleList[j].SampleList.Count);
                NumericVector nV = new NumericVector(engine, list);

                engine.SetSymbol("lst", nV);
                engine.Evaluate("matr <- matrix(lst, ncol=" + sample.SubSampleList.Count.ToString() + ", byrow=FALSE)");
                tmp = engine.Evaluate("friedman.test(matr)").AsList();
                double p = tmp["p.value"].AsNumeric().First();
                sample.PValue = p;

                //Посчитаем медиану и стандартное отклонение
                sample = FillingResults(sample.SubSampleList, sample);
                resDepFridman = sample;

                //Если получилась статистически значимая связь, то нарисуем график
                if (p <= settings.Statistical_significance)
                    //Нарисовать график для выборки
                    CreateGraphic(sample.SubSampleList, "FR");
            }
            //Отобразим кнопку для вывода в Word
            btnDepReport.Visible = true;
        }

        /// <summary>
        /// Вывод результатов проверки зависимых выборок в Word
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDepReport_Click(object sender, EventArgs e)
        {
            WordDocument report = new WordDocument(WordDocumentFormat.InCentimeters(21, 29.7, 2.5, 1, 2, 2));
            //1.5 отступ в документе
            report.WriteControlWord(@"sl360\slmult1");
            report.SetFont(settings.FontStandart);
            report.SetTextAlign(WordTextAlign.Center);
            report.WriteLine("Сравнение средних по 2 зависимым выборкам ");
            //Сквозная нумерация
            endToEndTable = 1;
            endToEndPicture = 1;
            report = OutResultInTableDependent(report, "Вилкоксона", resDepWilcoxon);
            report = OutResultInTableDependent(report, "Фридмана", resDepFridman);

            report.SaveToFile("..\\..\\ResultDep.doc");
            System.Diagnostics.Process.Start("..\\..\\ResultDep.doc");

            //Удалим все созданные графики
            var files = Directory.GetFiles(pathC);
            foreach (string file in files)
                File.Delete(file);
        }

        /// <summary>
        /// Вывод информации в таблицы
        /// </summary>
        /// <param name="_wordDocument">Документ</param>
        /// <param name="_methodName">Название метода</param>
        /// <param name="_resList">Список с результатами проверки</param>
        /// <returns></returns>
        private WordDocument OutResultInTableDependent(WordDocument _wordDocument, string _methodName, Sample _res)
        {
            if (_res != null)
            {
                //Получим пути созданных графиков для текущего метода
                string[] dirs;
                //string methodName = "";
                bool pval = false;
                if (_methodName.Contains("Вилкоксона"))
                {
                    dirs = Directory.GetFiles(pathC, "WI*"); //Результаты для Манна-Уитни
                    //methodName = "WI";
                }
                else
                {
                    dirs = Directory.GetFiles(pathC, "FR*"); //Результаты для Краскела-Уоллиса
                    //methodName = "FR";
                }

                //Получим список названий графиков
                List<string> graphNames = new List<string>();
                for (int j = 0; j < dirs.Length; j++)
                {
                    string name = Path.GetFileNameWithoutExtension(dirs[j]);
                    graphNames.Add(name);
                }

                //Текст перед таблицей
                string text = "В таблице " + endToEndTable.ToString() + " приведен результат статистического анализа данных, проведенного с использованием "
                                + "критерия " + _methodName + ".";
                _wordDocument.SetTextAlign(WordTextAlign.Justified);
                _wordDocument.SetParagraph(0, 567);
                _wordDocument.WriteLine(text);

                //Название таблицы
                string tableNumber = Environment.NewLine + "Таблица " + endToEndTable.ToString() + " - ";
                string tableCaption = "Сравнение средних уровней переменных: ";
                for (int j = 0; j < _res.SubSampleList.Count - 1; j++)
                    tableCaption += "\"" + _res.SubSampleList[j].UniqueVal + "\"" + ", ";
                tableCaption += "\"" + _res.SubSampleList.Last().UniqueVal + "\"";
                _wordDocument.SetTextAlign(WordTextAlign.Left);
                _wordDocument.SetParagraph(0, 0);
                _wordDocument.WriteLine(tableNumber + tableCaption);

                //Создаем табилицу и заполняем шапку
                _wordDocument.SetParagraph(0, 0);
                WordTable rt1 = _wordDocument.NewTable(settings.FontStandart, Color.Black, 2, _res.SubSampleList.Count + 1, 2);
                //Заполнение таблицы
                for (int k = 0; k < _res.SubSampleList.Count; k++)
                    rt1.Rows[0][k].Write(_res.SubSampleList[k].UniqueVal + ", n = " + _res.SubSampleList[k].SampleList.Count.ToString());
                rt1.Rows[0][_res.SubSampleList.Count].Write("p-значение");

                //Отрисуем рамки у ячеек
                for (int j = 0; j < _res.SubSampleList.Count + 1; j++)
                    rt1.Rows[0][j].SetBorders(Color.Black, 1, true, true, true, true);

                //Выводим данные в таблицу
                _wordDocument.SetTextAlign(WordTextAlign.Left);
                for (int k = 0; k < _res.SubSampleList.Count; k++)
                {
                    rt1.Rows[1][k].WriteLine(_res.SubSampleList[k].AverageValue.ToString() + " " + (char)177 + " " + Math.Round(_res.SubSampleList[k].StandartDeviation, 3).ToString());
                    rt1.Rows[1][k].WriteLine(Math.Round(_res.SubSampleList[k].Median, 3).ToString());
                    rt1.Rows[1][k].Write("(" + Math.Round(_res.SubSampleList[k].LowerQuintile, 3).ToString() + "; " + Math.Round(_res.SubSampleList[k].TopQuintile, 3).ToString() + ")");
                }

                double p = Math.Round(_res.PValue, 3);
                if (p <= 0.05)
                {
                    if (p > 0.001)
                    {
                        rt1.SetFont(settings.FontBold);
                        rt1.Rows[1][_res.SubSampleList.Count].Write(p.ToString() + "*");
                    }
                    else
                    {
                        rt1.SetFont(settings.FontStandart);
                        rt1.Rows[1][_res.SubSampleList.Count].Write("p < 0.001");
                    }
                    pval = true;
                }
                else
                {
                    rt1.SetFont(settings.FontStandart);
                    rt1.Rows[1][_res.SubSampleList.Count + 1].Write(p.ToString());
                }
                rt1.SetFont(settings.FontStandart);
                
                //Отрисуем рамки у ячеек
                for (int k = 0; k < _res.SubSampleList.Count + 1; k++)
                    rt1.Rows[1][k].SetBorders(Color.Black, 1, true, true, true, true);
                
                rt1.SaveToDocument(10000, 0);

                //Примечание к таблице
                _wordDocument.WriteControlWord(@"sl360\slmult1");
                _wordDocument.SetTextAlign(WordTextAlign.Justified);
                string note = "Примечание: формат представления информации в ячейке: среднее значение " + (char)177 +
                              " среднеквадратическое отклонение, медиана, (нижний; верхний) квартили.";

                //Добавление информации о "*" в таблице, при условии, что в ней есть p-value < 0.05
                string star = "";
                string conclusion = "Из таблицы " + endToEndTable.ToString() + " видно, что ";
                if (pval)
                {
                    star = " * - статистически значимое различие при p-value < " + settings.Statistical_significance.ToString();
                    _wordDocument.WriteLine(note + star);

                    //Вывод из таблицы
                    conclusion += "статистически значимое различие есть у ";
                    if (_res.SubSampleList.Count == 2)
                    {
                        conclusion += "признаков " + "\"" + _res.SubSampleList[0].UniqueVal + "\"и " + _res.SubSampleList[1].UniqueVal;
                    }
                    else
                    {
                        conclusion += "признаков ";
                        for (int j = 0; j < _res.SubSampleList.Count; j++)
                            conclusion += "\"" + _res.SubSampleList[j].UniqueVal + "\"" + ", ";
                    }
                    conclusion += " с учетом уровня значимости равного " + settings.Statistical_significance.ToString() + ".";
                    _wordDocument.SetParagraph(0, 567);
                    _wordDocument.WriteLine(conclusion);

                    //Добавим график к результату статистического анализа
                    //Предисловие к графикам
                    string preface = "На рисунке 1 изображена диаграмма размахов признака, приведенного в таблице " + endToEndTable.ToString() + ".";
                    _wordDocument.WriteLine(preface);

                    //Вставим график
                    _wordDocument.SetTextAlign(WordTextAlign.Center);
                    _wordDocument.PutImage(dirs[0], 96); //96 - истинный dpi
                    _wordDocument.WriteLine();
                    //Подпись к графику
                    _wordDocument.WriteControlWord(@"sl360\slmult1");
                    note = "Рисунок 1 - Диаграмма размаха переменных ";
                    for (int j = 0; j < _res.SubSampleList.Count - 1; j++)
                        note += "\"" + _res.SubSampleList[j].UniqueVal + "\", ";
                    note += "\"" + _res.SubSampleList.Last().UniqueVal + "\"";
                    _wordDocument.WriteLine(note);
                }
                else
                {
                    _wordDocument.WriteLine(note + star);
                    conclusion += "статистически значимых различий не обнаружено. Заданный уровень значимости равен " + settings.Statistical_significance.ToString() + ".";
                    _wordDocument.SetParagraph(0, 567);
                    _wordDocument.WriteLine(conclusion);
                }
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
        private void btnCorrelExecute_Click(object sender, EventArgs e)
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
            btnCorrelReport.Visible = true;
        }

        /// <summary>
        /// Вывод результатов корреляционного анализа в Word
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCorrelReport_Click(object sender, EventArgs e)
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
                    wt.SaveToDocument(10000, 0);

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
        private void btnTableExecute_Click(object sender, EventArgs e)
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
            btnTableReport.Visible = true;
        }

        /// <summary>
        /// Выведем таблицы сопряженности в файл
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnTableReport_Click(object sender, EventArgs e)
        {
            WordDocument report = new WordDocument(WordDocumentFormat.InCentimeters(21, 29.7, 2.5, 1, 2, 2));

            report.SetFont(settings.FontStandart);
            report.SetTextAlign(WordTextAlign.Center);
            report.WriteLine("Таблицы сопряженности ");

            report = OutResultInContingencyTable(report, resContingencyTables);

            report.SaveToFile("..\\..\\ResultTable.doc");
            System.Diagnostics.Process.Start("..\\..\\ResultTable.doc");
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

                rt1.SaveToDocument(10000, 0);
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
            Size size = grpBIndepGroupFact.Size;
            size.Width = (metroTabPage2.Size.Width - 10) / 2;
            grpBIndepGroupFact.Size = size;
            //Бинарные переменные
            size = grpBIndepBin.Size;
            size.Width = (grpBIndepGroupFact.Size.Width / 2) - 10;
            grpBIndepBin.Size = size;
            //Номинальные переменные
            grpBIndepNom.Location = new Point(size.Width + settings.PaddingBetweenGroupBoxes, grpBIndepNom.Location.Y);
            grpBIndepNom.Size = new Size(size.Width, grpBIndepNom.Size.Height);

            //Признаки
            size = grpBIndepGroupFact.Size;
            grpBIndepSign.Location = new Point(size.Width + settings.PaddingBetweenGroupBoxes, grpBIndepSign.Location.Y);
            grpBIndepSign.Size = size;
            //Количественные переменные
            size = grpBIndepCol.Size;
            size.Width = (grpBIndepSign.Size.Width / 2) - 10;
            grpBIndepCol.Size = size;
            //Порядковые переменные
            grpBIndepPor.Location = new Point(size.Width + settings.PaddingBetweenGroupBoxes, grpBIndepPor.Location.Y);
            grpBIndepPor.Size = new Size(size.Width, grpBIndepPor.Size.Height);

            //Кнопки
            //Подвинем кнопки под groupBox'ами
            //Изменение положения кнопок бинарных переменных
            int padding = settings.PaddingFromBtnSelectAll;
            Point grpLoc = grpBIndepBin.Location;
            btnIndepBinSelectAll.Location = new Point(grpLoc.X + padding, grpBIndepBin.Height + grpLoc.Y + settings.StandartPadding * 2);
            btnIndepBinCleanAll.Location = new Point(grpLoc.X + grpBIndepBin.Width - btnIndepBinCleanAll.Width - padding, grpBIndepBin.Height + grpLoc.Y + settings.StandartPadding * 2);

            //Изменение положения кнопок номинальных переменных
            grpLoc = grpBIndepNom.Location;
            btnIndepNomSelectAll.Location = new Point(grpLoc.X + padding, grpBIndepNom.Height + grpLoc.Y + settings.StandartPadding * 2);
            btnIndepNomCleanAll.Location = new Point(grpLoc.X + grpBIndepNom.Width - btnIndepNomSelectAll.Width - padding, grpBIndepNom.Height + grpLoc.Y + settings.StandartPadding * 2);

            //Изменение положения кнопок количественных переменных
            grpLoc = grpBIndepCol.Location;
            btnIndepColSelectAll.Location = new Point(grpLoc.X + padding, grpBIndepCol.Height + grpLoc.Y + settings.StandartPadding * 2);
            btnIndepColCleanAll.Location = new Point(grpLoc.X + grpBIndepCol.Width - btnIndepColCleanAll.Width - padding, grpBIndepCol.Height + grpLoc.Y + settings.StandartPadding * 2);

            //Изменение положения кнопок порядковых переменных  
            grpLoc = grpBIndepPor.Location;
            btnIndepPorSelectAll.Location = new Point(grpLoc.X + padding, grpBIndepPor.Height + grpLoc.Y + settings.StandartPadding * 2);
            btnIndepPorCleanAll.Location = new Point(grpLoc.X + grpBIndepPor.Width - btnIndepPorCleanAll.Width - padding, grpBIndepPor.Height + grpLoc.Y + settings.StandartPadding * 2);

            //Выровняем кнопки "Выполнить сравнение" и "Вывести в Word"
            btnIndepExecute.Location = new Point(grpBIndepGroupFact.Location.X + grpBIndepGroupFact.Size.Width - btnIndepExecute.Width, grpBIndepGroupFact.Size.Height + grpBIndepGroupFact.Location.Y + settings.StandartPadding * 2);
            btnIndepReport.Location = new Point(grpBIndepSign.Location.X, grpBIndepSign.Size.Height + grpBIndepSign.Location.Y + settings.StandartPadding * 2);
        }

        /// <summary>
        /// Отрисовка графиков по полученным выборкам
        /// </summary>
        /// <param name="_samples">Список сгруппированных выборок</param>
        private void CreateGraphic(Sample _sample, string _MethodName)
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
            engine.Evaluate("jpeg(\"" + pathR + _MethodName + "_" +_sample.GroupFact + "_" + _sample.NameSign + ".jpg\")");
            engine.Evaluate("boxplot(" + data + ", main=\"" + "Диаграмма размаха" + "\", names=c(" + names + "), ylab=\"" + _sample.NameSign.ToString()+"\", xlab=\"" + _sample.GroupFact.ToString() + "\")");
            engine.Evaluate("dev.off()");
        }

        /// <summary>
        /// Отрисовка графиков по полученным выборкам
        /// </summary>
        /// <param name="_samples">Список сгруппированных выборок</param>
        private void CreateGraphic(List<Sample.SubSample> _subSample, string _MethodName)
        {
            //Создадим дирректорию для хранения графиков
            if (!Directory.Exists(pathC))
                Directory.CreateDirectory(pathC);

            string data = "";
            string names = "";
            for (int i = 0; i < _subSample.Count; i++)
            {
                //Создать вектор
                string name = "group" + i.ToString();
                NumericVector group = engine.CreateNumericVector(_subSample[i].SampleList);
                //Перевести его в R
                engine.SetSymbol(name, group);
                if (i != _subSample.Count - 1)
                    data += name + ",";
                else
                    data += name;

                if (i != _subSample.Count - 1)
                    names += "\"" + _subSample[i].UniqueVal + "\"" + ",";
                else
                    names += "\"" + _subSample[i].UniqueVal + "\"";
            }
            engine.Evaluate("jpeg(\"" + pathR + _MethodName + "_" + names.Replace("\"","") + ".jpg\")");
            engine.Evaluate("boxplot(" + data + ", main=\"" + "Диаграмма размаха" + "\", names=c(" + names + "))");
            engine.Evaluate("dev.off()");
        }

        private void metroTabPage3_SizeChanged(object sender, EventArgs e)
        {
            //grpBoxes 7 -> [8, 9]
            //Bth 100 - 104 ;   105, 106
            //970 -> 480 -> 230

            //Признаки
            //Контейнер для признаков
            Size size = grpBDepSign.Size;
            size.Width = (metroTabPage3.Size.Width - 10) / 2;
            grpBDepSign.Size = size;
            //Количественные переменные
            size = grpBDepCol.Size;
            size.Width = (grpBDepSign.Size.Width / 2) - 10;
            grpBDepCol.Size = size;
            //Порядковые переменные
            grpBDepPor.Location = new Point(size.Width + settings.PaddingBetweenGroupBoxes, grpBDepPor.Location.Y);
            grpBDepPor.Size = new Size(size.Width, grpBDepPor.Size.Height);

            //Кнопки
            int padding = settings.PaddingFromBtnSelectAll;
            Point grpLoc = grpBDepCol.Location;
            btnDepColSelectAll.Location = new Point(grpLoc.X + padding, grpBDepCol.Height + grpLoc.Y + settings.StandartPadding * 2);
            btnDepColCleanAll.Location = new Point(grpLoc.X + grpBDepCol.Width - btnDepColCleanAll.Width - padding, grpBDepCol.Height + grpLoc.Y + settings.StandartPadding * 2);
            grpLoc = grpBDepPor.Location;
            btnDepPorSelectAll.Location = new Point(grpLoc.X + padding, grpBDepPor.Height + grpLoc.Y + settings.StandartPadding * 2);
            btnDepPorCleanAll.Location = new Point(grpLoc.X + grpBDepPor.Width - btnDepPorCleanAll.Width - padding, grpBDepPor.Height + grpLoc.Y + settings.StandartPadding * 2);

            //Выровняем кнопки "Выполнить сравнение" и "Вывести в Word"
            btnDepExecute.Location = new Point(grpBDepSign.Location.X + grpBDepSign.Size.Width - btnIndepExecute.Width, grpBDepSign.Size.Height + grpBDepSign.Location.Y + settings.StandartPadding * 2);
            btnDepReport.Location = new Point(grpBDepSign.Location.X, grpBDepSign.Size.Height + grpBDepSign.Location.Y + settings.StandartPadding * 2);
        }

        private void metroTabPage4_SizeChanged(object sender, EventArgs e)
        {
            //grpBoxes 10 -> [11, 12]
            //Bth 200 - 204 ;   205, 206
            //970 -> 480 -> 230

            //Признаки
            //Контейнер для признаков
            Size size = grpBCorrelSign.Size;
            size.Width = (metroTabPage4.Size.Width - 10) / 2;
            grpBCorrelSign.Size = size;
            //Количественные переменные
            size = grpBCorrelCol.Size;
            size.Width = (grpBCorrelSign.Size.Width / 2) - 10;
            grpBCorrelCol.Size = size;
            //Порядковые переменные
            grpBCorrelPor.Location = new Point(size.Width + settings.PaddingBetweenGroupBoxes, grpBCorrelPor.Location.Y);
            grpBCorrelPor.Size = new Size(size.Width, grpBCorrelPor.Size.Height);

            //Кнопки
            int padding = settings.PaddingFromBtnSelectAll;
            Point grpLoc = grpBCorrelCol.Location;
            btnCorrelColSelectAll.Location = new Point(grpLoc.X + padding, grpBCorrelCol.Height + grpLoc.Y + settings.StandartPadding * 2);
            btnCorrelColCleanAll.Location = new Point(grpLoc.X + grpBCorrelCol.Width - btnCorrelColCleanAll.Width - padding, grpBCorrelCol.Height + grpLoc.Y + settings.StandartPadding * 2);
            grpLoc = grpBCorrelPor.Location;
            btnCorrelPorSelectAll.Location = new Point(grpLoc.X + padding, grpBCorrelPor.Height + grpLoc.Y + settings.StandartPadding * 2);
            btnCorrelPorCleanAll.Location = new Point(grpLoc.X + grpBCorrelPor.Width - btnCorrelPorCleanAll.Width - padding, grpBCorrelPor.Height + grpLoc.Y + settings.StandartPadding * 2);

            //Выровняем кнопки "Выполнить сравнение" и "Вывести в Word"
            btnCorrelExecute.Location = new Point(grpBCorrelSign.Location.X + grpBCorrelSign.Size.Width - btnIndepExecute.Width, grpBCorrelSign.Size.Height + grpBCorrelSign.Location.Y + settings.StandartPadding * 2);
            btnCorrelReport.Location = new Point(grpBCorrelSign.Location.X, grpBCorrelSign.Size.Height + grpBCorrelSign.Location.Y + settings.StandartPadding * 2);
        }

        private void metroTabPage5_SizeChanged(object sender, EventArgs e)
        {
            //grpBoxes 13 -> [14, 15]
            //10 11 12
            //Bth 300 - 304 ;   305, 306
            //970 -> 480 -> 230

            //Признаки
            //Контейнер для признаков
            Size size = grpBTableSign.Size;
            size.Width = (metroTabPage5.Size.Width - 10) / 2;
            grpBTableSign.Size = size;
            //Номинальные переменные
            size = grpBTableNom.Size;
            size.Width = (grpBTableSign.Size.Width / 2) - 10;
            grpBTableNom.Size = size;
            //Бинарные переменные
            grpBTableBin.Location = new Point(size.Width + settings.PaddingBetweenGroupBoxes, grpBTableBin.Location.Y);
            grpBTableBin.Size = new Size(size.Width, grpBTableBin.Size.Height);

            //Кнопки
            int padding = settings.PaddingFromBtnSelectAll;
            Point grpLoc = grpBTableNom.Location;
            btnTableNomSelectAll.Location = new Point(grpLoc.X + padding, grpBTableNom.Height + grpLoc.Y + settings.StandartPadding * 2);
            btnTableNomCleanAll.Location = new Point(grpLoc.X + grpBTableNom.Width - btnTableNomCleanAll.Width - padding, grpBTableNom.Height + grpLoc.Y + settings.StandartPadding * 2);
            grpLoc = grpBTableBin.Location;
            btnTableBinSelectAll.Location = new Point(grpLoc.X + padding, grpBTableBin.Height + grpLoc.Y + settings.StandartPadding * 2);
            btnTableBinCleanAll.Location = new Point(grpLoc.X + grpBTableBin.Width - btnTableBinCleanAll.Width - padding, grpBTableBin.Height + grpLoc.Y + settings.StandartPadding * 2);

            //Выровняем кнопки "Выполнить сравнение" и "Вывести в Word"
            btnTableExecute.Location = new Point(grpBTableSign.Location.X + grpBTableSign.Size.Width - btnIndepExecute.Width, grpBTableSign.Size.Height + grpBTableSign.Location.Y + settings.StandartPadding * 2);
            btnTableReport.Location = new Point(grpBTableSign.Location.X, grpBTableSign.Size.Height + grpBTableSign.Location.Y + settings.StandartPadding * 2);

        }
    }
}