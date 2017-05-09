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
        /// Ширина и высота формы
        /// </summary>
        private int formWidth;
        private int formHeight;

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
        Properties.Settings settings = new Properties.Settings();

        /// <summary>
        /// Считанные данные
        /// </summary>
        private Data data;

        /// <summary>
        /// Получаем список индексов с Бин и Ном шкалами
        /// </summary>
        private List<int> binList = new List<int>();

        /// <summary>
        /// Получаем список индексов с Ном шкалами
        /// </summary>
        private List<int> nomList = new List<int>();

        /// <summary>
        /// Получаем список индексов с Кол шкалой
        /// </summary>
        private List<int> colList = new List<int>();

        /// <summary>
        /// Получаем список индексов с Пор шкалой
        /// </summary>
        private List<int> porList = new List<int>();

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
        private List<ContingencyTable> contingencyTable = new List<ContingencyTable>();

        #endregion

        public Form1()
        {
            InitializeComponent();
            REngine.SetEnvironmentVariables();

            formWidth = Width;
            formHeight = Height;
        }

        /// <summary>
        /// Загрузка формы. Начальные настройки формы
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {

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
            for (int i = 0; i < records; i++)
                metroGrid1.Rows[0].Cells[i].Value = data.TakeUniqueValuesAtIndex(i);
            for (int i = 0; i < records; i++)
                metroGrid1.Rows[1].Cells[i].Value = data.TakeScaleTypeAtIndex(i);
            for (int i = 0; i < records; i++)
                metroGrid1.Rows[2].Cells[i].Value = data.TakeVariableNameAtIndex(i);
            for (int j = 0; j < data.PatientsCount; j++)
            {
                List<string> person;
                person = data.TakePatientAtIndex(j);
                for (int i = 0; i < records; i++)
                    metroGrid1.Rows[j + 3].Cells[i].Value = person[i];
            }
        }

        /// <summary>
        /// Формирование списков по типу шкал
        /// </summary>
        private void CreateListsByTypeScale()
        {
            colList.Clear();
            porList.Clear();
            binList.Clear();
            nomList.Clear();

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
                metroButton302.Visible = false;

                //Кнопки вывод в Word
                metroButton10.Visible = false;
                metroButton105.Visible = false;
                metroButton205.Visible = false;
                metroButton303.Visible = false;

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
                metroButton302.Visible = false;

                //Кнопки вывод в Word
                metroButton10.Visible = false;
                metroButton105.Visible = false;
                metroButton205.Visible = false;
                metroButton303.Visible = false;

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
                metroButton302.Visible = false;

                //Кнопки вывод в Word
                metroButton10.Visible = false;
                metroButton105.Visible = false;
                metroButton205.Visible = false;
                metroButton303.Visible = false;

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
            }

            #endregion

            #region Таблицы сопряженности

            if (e.TabPageIndex == 4)
            {
                groupBox14.Controls.Clear();
                MetroCheckBox[] rB;

                //Скроем кнопки с других вкладок
                metroButton9.Visible = false;
                metroButton104.Visible = false;
                metroButton204.Visible = false;

                //Кнопки вывод в Word
                metroButton10.Visible = false;
                metroButton105.Visible = false;
                metroButton205.Visible = false;
                metroButton303.Visible = false;

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

                //Выровняем кнопки "Выполнить сравнение" и "Вывести в Word"
                point = new Point(metroButton302.Location.X, groupBox13.Size.Height);
                point.Y += 20;
                metroButton302.Location = point;
                point = new Point(metroButton303.Location.X, groupBox13.Size.Height);
                point.Y += 20;
                metroButton303.Location = point;
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
                    metroButton302.Visible = true;
            }
            else
            {
                groupFactList.Remove(index);
                if (groupFactList.Count == 0)
                    groupBox4.Visible = false;
                if (Convert.ToInt32(mCB.Parent.Tag) == 4 && groupFactList.Count < 2)
                    metroButton302.Visible = false;
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
                tmp.LowerQuintile = quantile[0];
                tmp.TopQuintile = quantile[4];
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
                        CreateGraphic(samples[i], "MW", i);
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
                        CreateGraphic(samples[i], "KW", i);
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
            WordDocument report = new WordDocument(WordDocumentFormat.A4);

            report.SetFont(new Font("Times New Roman", 16, FontStyle.Regular, GraphicsUnit.Pixel));
            report.SetTextAlign(WordTextAlign.Center);
            report.WriteLine("Сравнение средних по 2 независимым выборкам ");

            report = OutResultInTableIndependent(report, "Критерий Манна-Уитни, ", resIndMannaWhitney);
            report = OutResultInTableIndependent(report, "Критерий Краскела-Уоллиса, ", resIndKruskalWallis);

            report.SaveToFile("..\\..\\ResultIndep.doc");
            System.Diagnostics.Process.Start("..\\..\\ResultIndep.doc");
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
            WordTable rt1;
            string grpFact = "";

            var grpRes = Grouping(_resList);
            for (int i = 0; i < grpRes.Count; i++)
            {
                Sample s;
                //Создаем табилицу и заполняем шапку
                rt1 = _wordDocument.NewTable(new Font("Times New Roman", 12, FontStyle.Regular), Color.Black, grpRes[i].Count + 1, grpRes[i][0].SubSampleList.Count + 3, 2);
                _wordDocument.SetTextAlign(WordTextAlign.Left);
                _wordDocument.SetFont(new Font("Times New Roman", 12, FontStyle.Bold, GraphicsUnit.Pixel));
                _wordDocument.SetTextAlign(WordTextAlign.Left);
                _wordDocument.WriteLine();
                _wordDocument.WriteLine(_methodName);
                grpFact = grpRes[i][0].GroupFact;
                _wordDocument.WriteLine("Группирующий фактор: " + grpFact.ToString());
                rt1.Rows[0][0].Write("N");
                rt1.Rows[0][1].Write("Показатель");
                for (int k = 0; k < grpRes[i][0].SubSampleList.Count; k++)
                    rt1.Rows[0][2 + k].Write(grpRes[i][0].SubSampleList[k].UniqueVal + ", n = " + grpRes[i][0].SubSampleList[k].SampleList.Count.ToString());
                rt1.Rows[0][grpRes[i][0].SubSampleList.Count + 2].Write("p-значение");

                //Отрисуем рамки у ячеек
                for (int j = 0; j < grpRes[i][0].SubSampleList.Count + 3; j++)
                    rt1.Rows[0][j].SetBorders(Color.Black, 1, true, true, true, true);

                //Выводим данные в таблицу
                for (int j = 0; j < grpRes[i].Count; j++)
                {
                    s = grpRes[i][j];
                    _wordDocument.SetTextAlign(WordTextAlign.Left);
                    rt1.Rows[j + 1][0].Write((j + 1).ToString());
                    rt1.Rows[j + 1][1].Write(s.NameSign);
                    for (int k = 0; k < grpRes[i][0].SubSampleList.Count; k++)
                    {
                        rt1.Rows[j + 1][2 + k].WriteLine(s.SubSampleList[k].AverageValue.ToString() + " " + (char)177 + " " + Math.Round(s.SubSampleList[k].StandartDeviation, 3).ToString());
                        rt1.Rows[j + 1][2 + k].WriteLine(Math.Round(s.SubSampleList[k].Median, 3).ToString());
                        rt1.Rows[j + 1][2 + k].Write("(" + Math.Round(s.SubSampleList[k].LowerQuintile, 3).ToString() + "; " + Math.Round(s.SubSampleList[k].TopQuintile, 3).ToString() + ")");
                    }

                    double p = grpRes[i][j].PValue;
                    rt1.Rows[j + 1][grpRes[i][0].SubSampleList.Count + 2].Write((Math.Round(p, 3).ToString()));

                    //Отрисуем рамки у ячеек
                    for (int k = 0; k < grpRes[i][0].SubSampleList.Count + 3; k++)
                        rt1.Rows[j + 1][k].SetBorders(Color.Black, 1, true, true, true, true);
                }
                rt1.SaveToDocument(9600, 0);
            }

            //Вставка графиков в отчет
            //Добавим отступ от таблицы
            _wordDocument.WriteLine();
            if (_methodName.Contains("Манна"))
            {
                //Результаты для Манна-Уитни
                string[] dirs = Directory.GetFiles(pathC, "MW*");
                for (int i = 0; i < dirs.Length; i++)
                {
                    _wordDocument.PutImage(dirs[i], 96);
                    _wordDocument.WriteLine();
                }
            }
            else
            {
                //Результаты для Краскела-Уоллиса
                string[] dirs = Directory.GetFiles(pathC, "KW*");
                for (int i = 0; i < dirs.Length; i++)
                {
                    _wordDocument.PutImage(dirs[i], 96);
                    _wordDocument.WriteLine();
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
            WordDocument report = new WordDocument(WordDocumentFormat.A4);

            report.SetFont(new Font("Times New Roman", 16, FontStyle.Regular, GraphicsUnit.Pixel));
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
                rt1 = _wordDocument.NewTable(new Font("Times New Roman", 12, FontStyle.Regular), Color.Black, grpRes[i].Count + 1, grpRes[i][0].SubSampleList.Count + 1, 2);
                _wordDocument.SetTextAlign(WordTextAlign.Left);
                _wordDocument.SetFont(new Font("Times New Roman", 12, FontStyle.Bold, GraphicsUnit.Pixel));
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
            WordDocument report = new WordDocument(WordDocumentFormat.A4);

            report.SetFont(new Font("Times New Roman", 16, FontStyle.Regular, GraphicsUnit.Pixel));
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

                    WordTable wt = _wordDocument.NewTable(new Font("Times New Roman", 12, FontStyle.Regular), Color.Black, widthTable + 1, heightTable + 1, 2);
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
                    _wordDocument.SetFont(new Font("Times New Roman", 12, FontStyle.Regular, GraphicsUnit.Pixel));
                    _wordDocument.SetTextAlign(WordTextAlign.Left);
                    _wordDocument.WriteLine(Environment.NewLine + "Таблица " + (tables.Count + 1).ToString() + " - Корреляционный анализ. " + _methodName);

                    tables.Add(wt);
                    wt.SaveToDocument(9600, 0);

                    //Добавим примечание
                    _wordDocument.SetTextAlign(WordTextAlign.Justified);
                    _wordDocument.SetFont(new Font("Times New Roman", 12, FontStyle.Regular, GraphicsUnit.Pixel));
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
                            w.Rows[i + 1][j + 1].SetFont(new Font("Times New Roman", 12, FontStyle.Bold, GraphicsUnit.Pixel));
                            w.Rows[i + 1][j + 1].WriteLine("r = " + Math.Round(r, 3).ToString());
                            if (p > 0.001)
                                w.Rows[i + 1][j + 1].Write("p = " + Math.Round(p, 3).ToString());
                            else
                                w.Rows[i + 1][j + 1].Write("p < 0.001");
                        }
                        else
                        {
                            w.Rows[i + 1][j + 1].SetFont(new Font("Times New Roman", 12, FontStyle.Regular, GraphicsUnit.Pixel));
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
                                wt.Rows[str + 1][j + 1].SetFont(new Font("Times New Roman", 12, FontStyle.Bold, GraphicsUnit.Pixel));
                                wt.Rows[str + 1][j + 1].WriteLine("r = " + Math.Round(r, 3).ToString());
                                if (p > 0.001)
                                    wt.Rows[str + 1][j + 1].Write("p = " + Math.Round(p, 3).ToString());
                                else
                                    wt.Rows[str + 1][j + 1].Write("p < 0.001");
                            }
                            else
                            {
                                wt.Rows[str + 1][j + 1].SetFont(new Font("Times New Roman", 12, FontStyle.Regular, GraphicsUnit.Pixel));
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
                            wt.Rows[str + 1][j + 1].SetFont(new Font("Times New Roman", 12, FontStyle.Bold, GraphicsUnit.Pixel));
                            wt.Rows[str + 1][j + 1].WriteLine("r = " + Math.Round(r, 3).ToString());
                            if (p > 0.001)
                                wt.Rows[str + 1][j + 1].Write("p = " + Math.Round(p, 3).ToString());
                            else
                                wt.Rows[str + 1][j + 1].Write("p < 0.001");
                        }
                        else
                        {
                            wt.Rows[str + 1][j + 1].SetFont(new Font("Times New Roman", 12, FontStyle.Regular, GraphicsUnit.Pixel));
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
            //Получим список таблиц
            contingencyTable = ContingencyTable.GetTableList(data, groupFactList);
            metroButton303.Visible = true;
        }

        /// <summary>
        /// Выведем таблицы сопряженности в файл
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void metroButton303_Click(object sender, EventArgs e)
        {
            WordDocument report = new WordDocument(WordDocumentFormat.A4);

            report.SetFont(new Font("Times New Roman", 16, FontStyle.Regular, GraphicsUnit.Pixel));
            report.SetTextAlign(WordTextAlign.Center);
            report.WriteLine("Таблицы смежности ");

            report = OutResultInContingencyTable(report, contingencyTable);

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
        private WordDocument OutResultInContingencyTable(WordDocument _wordDocument, List<ContingencyTable> _resList)
        {
            //Идем по спискам тиблиц
            for (int i = 0; i < _resList.Count; i++)
            {
                WordTable rt1 = _wordDocument.NewTable(new Font("Times New Roman", 12, FontStyle.Regular), Color.Black, _resList[i].Variable1.Count + 2, _resList[i].Variable2.Count + 2, 2);
                _wordDocument.SetTextAlign(WordTextAlign.Left);
                _wordDocument.SetFont(new Font("Times New Roman", 12, FontStyle.Bold, GraphicsUnit.Pixel));
                _wordDocument.SetTextAlign(WordTextAlign.Left);
                _wordDocument.WriteLine();

                //Выведем список уникальных значений первой переменной
                for (int j = 0; j < _resList[i].Variable1.Count; j++)
                    rt1.Rows[j + 1][0].Write(_resList[i].Variable1[j]);
                //Выведем список уникальных значений второй переменной
                for (int j = 0; j < _resList[i].Variable2.Count; j++)
                    rt1.Rows[0][j + 1].Write(_resList[i].Variable2[j]);

                //Выведем данные в таблицу
                //Сумма по столбцам
                List<int> sumCol = new List<int>();
                //Сбросим сумму на 0-ли
                for (int j = 0; j < _resList[i].Variable2.Count; j++)
                    sumCol.Add(0);

                //Идем по строкам
                for (int j = 0; j < _resList[i].Variable1.Count; j++)
                {
                    //Посчитаем сумму в строке
                    int sum = 0;
                    //Идем по столбцам
                    for (int k = 0; k < _resList[i].Variable2.Count; k++)
                    {
                        rt1.Rows[j + 1][k + 1].Write(_resList[i].Data[k + j * _resList[i].Variable2.Count].ToString());
                        sum += _resList[i].Data[k + j * _resList[i].Variable2.Count];
                        sumCol[k] += _resList[i].Data[k + j * _resList[i].Variable2.Count];
                    }
                    //Выведем сумму в конец сторки
                    rt1.Rows[j + 1][_resList[i].Variable2.Count + 1].Write(sum.ToString());
                }
                //Выведем сумму по столбцам
                for (int j = 0; j < sumCol.Count; j++)
                    rt1.Rows[_resList[i].Variable1.Count + 1][j + 1].Write(sumCol[j].ToString());

                //Сумма в последней строке. Всего элементов в таблице
                rt1.Rows[_resList[i].Variable1.Count + 1][_resList[i].Variable2.Count + 1].Write(_resList[i].Data.Sum().ToString());

                //Нарисуем рамки у ячеек
                for (int j = 0; j < _resList[i].Variable1.Count + 2; j++)
                {
                    for (int k = 0; k < _resList[i].Variable2.Count + 2; k++)
                        rt1.Rows[j][k].SetBorders(Color.Black, 1, true, true, true, true);
                }

                rt1.SaveToDocument(9600, 0);
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
            int rWidth = Width - formWidth;

            /***************************Внешний вид кладки***************************/
            //GroupBox'ы
            //Группирующие факторы
            Size size = groupBox1.Size;
            size.Width += rWidth / 2;
            groupBox1.Size = size;
            //Бинарные переменные
            size = groupBox2.Size;
            size.Width += rWidth / 4;
            groupBox2.Size = size;
            //Номинальные переменные
            groupBox3.Location = new Point(size.Width + settings.PaddingBetweenGroupBoxes, groupBox3.Location.Y);
            groupBox3.Size = size;

            //Признаки
            size = groupBox1.Size;
            groupBox4.Location = new Point(size.Width + settings.PaddingBetweenGroupBoxes, groupBox4.Location.Y);
            groupBox4.Size = size;
            //Количественные переменные
            size = groupBox5.Size;
            size.Width += rWidth / 4;
            groupBox5.Size = size;
            //Порядковые переменные
            groupBox6.Location = new Point(size.Width + settings.PaddingBetweenGroupBoxes, groupBox6.Location.Y);
            groupBox6.Size = size;

            //Кнопки
            //Подвинем кнопки под groupBox'ами
            //Изменение положения кнопок бинарных переменных
            metroButton1.Location = new Point(groupBox2.Location.X + settings.PaddingFromBtnSelectAll + rWidth / 8, groupBox2.Size.Height + groupBox2.Location.Y + settings.StandartPadding * 2);
            metroButton2.Location = new Point(groupBox2.Location.X + settings.PaddingFromBtnCleanAll + rWidth / 8, groupBox2.Size.Height + groupBox2.Location.Y + settings.StandartPadding * 2);

            //Изменение положения кнопок номинальных переменных
            metroButton3.Location = new Point(groupBox3.Location.X + settings.PaddingFromBtnSelectAll + rWidth / 8, groupBox3.Size.Height + groupBox3.Location.Y + settings.StandartPadding * 2);
            metroButton4.Location = new Point(groupBox3.Location.X + settings.PaddingFromBtnCleanAll + rWidth / 8, groupBox3.Size.Height + groupBox3.Location.Y + settings.StandartPadding * 2);
            
            //Изменение положения кнопок количественных переменных
            metroButton5.Location = new Point(groupBox5.Location.X + settings.PaddingFromBtnSelectAll + rWidth / 8, groupBox5.Size.Height + groupBox5.Location.Y + settings.StandartPadding * 2);
            metroButton6.Location = new Point(groupBox5.Location.X + settings.PaddingFromBtnCleanAll + rWidth / 8, groupBox5.Size.Height + groupBox5.Location.Y + settings.StandartPadding * 2);
            
            //Изменение положения кнопок порядковых переменных            
            metroButton7.Location = new Point(groupBox6.Location.X + settings.PaddingFromBtnSelectAll + rWidth / 8, groupBox6.Size.Height + groupBox6.Location.Y + settings.StandartPadding * 2);
            metroButton8.Location = new Point(groupBox6.Location.X + settings.PaddingFromBtnCleanAll + rWidth / 8, groupBox6.Size.Height + groupBox6.Location.Y + settings.StandartPadding * 2);

            //Выровняем кнопки "Выполнить сравнение" и "Вывести в Word"
            metroButton9.Location = new Point(groupBox1.Location.X + groupBox1.Size.Width - metroButton9.Width, groupBox1.Size.Height + groupBox1.Location.Y + settings.StandartPadding * 2);
            metroButton10.Location = new Point(groupBox4.Location.X, groupBox4.Size.Height + groupBox4.Location.Y + settings.StandartPadding * 2);

            formWidth = Width;
            formHeight = Height;
        }

        /// <summary>
        /// Отрисовка графиков по полученным выборкам
        /// </summary>
        /// <param name="_samples">Список сгруппированных выборок</param>
        private void CreateGraphic(Sample _sample, string _Methodname, int _index)
        {
            //Создадим дирректорию для хранения графиков
            if (!Directory.Exists(pathC))
                Directory.CreateDirectory(pathC); 

            string data = "";
            string names = "";
            for (int i = 0; i < _sample.SubSampleList.Count; i++)
            {
                //Создать вектор
                string name = "group" + i.ToString();
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
            engine.Evaluate("jpeg(\"" + pathR + _Methodname + "graph" + _index.ToString() + ".jpg\")");
            engine.Evaluate("boxplot(" + data + ", main=\"" + _sample.GroupFact.ToString() + "\", names=c(" + names + "), ylab=\"" + _sample.NameSign.ToString()+"\")");
            engine.Evaluate("dev.off()");
        }
    }
}