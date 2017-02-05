using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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
        /// Список с результатами проверок для теста Манна-Уитни
        /// </summary>
        List<Sample> resManna = new List<Sample>();

        /// <summary>
        /// Список с результатами проверок для теста Краскела-Уоллиса
        /// </summary>
        List<Sample> resKruskel = new List<Sample>();

        #endregion

        public Form1()
        {
            InitializeComponent();
            REngine.SetEnvironmentVariables();
        }

        //Загрузка формы. Начальные настройки формы
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
            #region "Сравнение независимых групп"
            if (e.TabPageIndex == 1)
            {
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
                    rB[i].Location = new Point(6, 22 * (i + 1));
                    rB[i].CheckedChanged += CheckedChangedForGroupFac;
                    groupBox2.Controls.Add(rB[i]);
                }

                //Изменение положения кнопок
                Point point = new Point(metroButton1.Location.X, metroButton1.Location.Y);
                point.Y = groupBox2.Size.Height + 19/*groupBox2.Location.Y*/ + 6;
                metroButton1.Location = point;
                point = new Point(metroButton2.Location.X, metroButton2.Location.Y);
                point.Y = groupBox2.Size.Height + 25;
                metroButton2.Location = point;

                //Вывод номинальных переменных
                rB = new MetroCheckBox[nomList.Count];
                for (int i = 0; i < nomList.Count; i++)
                {
                    rB[i] = new MetroCheckBox();
                    rB[i].Text = data.TakeVariableNameAtIndex(nomList[i]);
                    rB[i].Checked = false;
                    rB[i].Tag = nomList[i];
                    rB[i].AutoSize = true;
                    rB[i].Location = new Point(6, 22 * (i + 1));
                    rB[i].CheckedChanged += CheckedChangedForGroupFac;
                    groupBox3.Controls.Add(rB[i]);
                }

                //Изменение положения кнопок
                point = new Point(metroButton3.Location.X, metroButton3.Location.Y);
                point.Y = groupBox3.Size.Height + 25;
                metroButton3.Location = point;
                point = new Point(metroButton4.Location.X, metroButton4.Location.Y);
                point.Y = groupBox3.Size.Height + 25;
                metroButton4.Location = point;

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
                    rB[i].Location = new Point(6, 22 * (i + 1));
                    rB[i].CheckedChanged += CheckedChangedForSigns;
                    groupBox5.Controls.Add(rB[i]);
                }

                //Изменение положения кнопок
                point = new Point(metroButton5.Location.X, metroButton5.Location.Y);
                point.Y = groupBox5.Size.Height + 25;
                metroButton5.Location = point;
                point = new Point(metroButton6.Location.X, metroButton6.Location.Y);
                point.Y = groupBox5.Size.Height + 25;
                metroButton6.Location = point;

                //Вывод списка порядковых переменных
                rB = new MetroCheckBox[porList.Count];
                for (int i = 0; i < porList.Count; i++)
                {
                    rB[i] = new MetroCheckBox();
                    rB[i].Text = data.TakeVariableNameAtIndex(porList[i]);
                    rB[i].Checked = false;
                    rB[i].Tag = porList[i];
                    rB[i].AutoSize = true;
                    rB[i].Location = new Point(6, 22 * (i + 1));
                    rB[i].CheckedChanged += CheckedChangedForSigns;
                    groupBox6.Controls.Add(rB[i]);
                }

                //Изменение положения кнопок
                point = new Point(metroButton7.Location.X, metroButton7.Location.Y);
                point.Y = groupBox6.Size.Height + 25;
                metroButton7.Location = point;
                point = new Point(metroButton8.Location.X, metroButton8.Location.Y);
                point.Y = groupBox6.Size.Height + 25;
                metroButton8.Location = point;

                //Выровним боксы для Груп. факторо и признаков
                if (groupBox1.Height > groupBox4.Height)
                    groupBox4.Height = groupBox1.Height;
                else
                    groupBox1.Height = groupBox4.Height;
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
            }
            else
            {
                groupFactList.Remove(index);
                if (groupFactList.Count == 0)
                    groupBox4.Visible = false;
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
            //Поставили галочку - добавить в список выбранных элементов и отобразить кнопку для сравнения
            //Убрали - убрать из списка и, если сняли все галочки, то скрыть кнопку для сравнения
            if (mCB.Checked == true)
            {
                signsList.Add(index);
                metroButton16.Visible = true;
            }
            else
            {
                signsList.Remove(index);
                if (signsList.Count == 0)
                    metroButton16.Visible = false;
            }
        }

        //Нажата кнопка "Выбрать все"
        private void SelectAll_Click(object sender, EventArgs e)
        {
            //Получим Тэг кнопки
            MetroButton mB = (MetroButton)sender;
            int tag = Convert.ToInt32(mB.Tag);

            //По Тэгу производим действия: добавляем в список выбранных элементов и проставляем галочки
            switch(tag)
            {
                //Бинарные переменные
                case 1: Check(groupBox2); break;
                //Номинальные переменные
                case 3: Check(groupBox3); break;
                //Количественные переменные
                case 5: Check(groupBox5); break;
                //Порядковые переменные
                default: Check(groupBox6); break;
            }
        }

        //Нажата кнопка "Очистить"
        private void CleanAll_Click(object sender, EventArgs e)
        {
            //Получим Тэг кнопки
            MetroButton mB = (MetroButton)sender;
            int tag = Convert.ToInt32(mB.Tag);

            //По Тэгу производим действия: удаляем элементы из списка и снимаем все галочки
            switch (tag)
            {
                //Бинарные переменные
                case 2: Uncheck(groupBox2); break;
                //Номинальные переменные
                case 4: Uncheck(groupBox3); break;
                //Количественные переменные
                case 6: Uncheck(groupBox5); break;
                //Порядковые переменные
                default: Uncheck(groupBox6); break;
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

        //Проставить галочки на выбранных элементах
        private void Check(GroupBox _groupBox)
        {
            Control.ControlCollection controls = _groupBox.Controls;
            foreach(MetroCheckBox control in controls)
                control.Checked = true;       
        }

        //Снять галочки
        private void Uncheck(GroupBox _groupBox)
        {
            Control.ControlCollection controls = _groupBox.Controls;
            foreach (MetroCheckBox control in controls)
                control.Checked = false;
        }

        //Выоплнить "Сравнение независимых групп"
        private void metroButton16_Click(object sender, EventArgs e)
        {
            Sample testRes;
            GenericVector tmp;
            engine = REngine.GetInstance();
            List<Sample> samples = new List<Sample>();
            samples = Sample.GetSample(data, groupFactList, signsList);
            for (int i = 0; i < samples.Count; i++)
            {
                testRes = new Sample();
                testRes = samples[i];
                if (samples[i].SubSampleList.Count == 2)
                {
                    //Вызвать тест Манна-Уитни
                    NumericVector group1 = engine.CreateNumericVector(samples[i].SubSampleList[0].SampleList);
                    engine.SetSymbol("Rgroup1", group1);
                    NumericVector group2 = engine.CreateNumericVector(samples[i].SubSampleList[1].SampleList);
                    engine.SetSymbol("Rgroup2", group2);
                    tmp = engine.Evaluate("wilcox.test(Rgroup1, Rgroup2)").AsList();
                    double p = tmp["p.value"].AsNumeric().First();
                    testRes.PValue = p;

                    //Посчитаем медиану и стандартное отклонение
                    testRes = FillingResults(samples[i].SubSampleList, testRes);
                    resManna.Add(testRes);
                }
                else
                {
                    //Вызвать тест Краскела-Уоллиса
                    //Создадим список выборок для отправки в тест Краскела-Уоллиса
                    GenericVector gV = new GenericVector(engine, samples[i].SubSampleList.Count);
                    NumericVector nV;
                    for (int j = 0; j < samples[i].SubSampleList.Count; j++)
                    {
                        nV = new NumericVector(engine, samples[i].SubSampleList[j].SampleList.Count);
                        for (int k = 0; k < samples[i].SubSampleList[j].SampleList.Count; k++)
                            nV[k] = samples[i].SubSampleList[j].SampleList[k];
                        gV[j] = nV;
                    }

                    engine.SetSymbol("sample", gV);
                    tmp = engine.Evaluate("kruskal.test(sample)").AsList();
                    double p = tmp["p.value"].AsNumeric().First();
                    testRes.PValue = p;

                    //Посчитаем медиану и стандартное отклонение
                    testRes = FillingResults(samples[i].SubSampleList, testRes);
                    resKruskel.Add(testRes);
                }
                metroButton15.Visible = true;
            }
        }

        //Заполнение результата дополнительными данными
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
                tmp.Amount = _s[i].SampleList.Count;
                _t.SubSampleList[i] = tmp;
            }
            return _t;
        }






    }
}
