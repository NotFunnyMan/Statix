﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Statix
{
    /*Класс "Независимые выборки"
     * Хранит информацию о выборках для соответствующей обработки
     */
    class Sample
    {
        #region Поля

        /// <summary>
        ///Группирующий фактор
        /// </summary>
        private string groupFact;

        /// <summary>
        ///Название призанка
        /// </summary>
        private string nameSign;
        
        /// <summary>
        ///Структура подвыборки
        /// </summary>
        public struct SubSample
        {
            #region Поля

            /// <summary>
            ///Уникальное значение
            /// </summary>
            private string uniqueVal;

            /// <summary>
            ///Элементы подвыборки
            /// </summary>
            private List<double> sampleList;

            /// <summary>
            ///Среднее значение
            /// </summary>
            private double averageValue;

            /// <summary>
            ///Количество элементов в подвыборке
            /// </summary>
            private int amount;

            /// <summary>
            ///Медиана выборки
            /// </summary>
            private double median;

            /// <summary>
            ///Стандартное отклонение
            /// </summary>
            private double standartDeviation;

            /// <summary>
            ///Нижняя кванить
            /// </summary>
            private double lowerQuintile;

            /// <summary>
            ///Верхняя квантиль
            /// </summary>
            private double topQuintile;

            #endregion

            #region Свойства

            /// <summary>
            /// Уникальное значение
            /// </summary>
            public string UniqueVal
            {
                set { uniqueVal = value; }
                get { return this.uniqueVal; }
            }

            /// <summary>
            /// Элементы подвыборки
            /// </summary>
            public List<double> SampleList
            {
                set { sampleList = value; }
                get { return this.sampleList; }
            }

            /// <summary>
            /// Среднее значение
            /// </summary>
            public double AverageValue
            {
                set { averageValue = value; }
                get { return this.averageValue; }
            }

            /// <summary>
            /// Количество элементов в подвыборке
            /// </summary>
            public int Amount
            {
                set { amount = value; }
                get { return this.amount; }
            }

            /// <summary>
            ///Медиана выборки
            /// </summary>
            public double Median
            {
                set { median = value; }
                get { return median; }
            }

            /// <summary>
            ///Стандартное отклонение
            /// </summary>
            public double StandartDeviation
            {
                set { standartDeviation = value; }
                get { return standartDeviation; }
            }

            /// <summary>
            ///Нижняя кванить
            /// </summary>
            public double LowerQuintile
            {
                set { lowerQuintile = value; }
                get { return lowerQuintile; }
            }

            /// <summary>
            ///Верхняя квантиль
            /// </summary>
            public double TopQuintile
            {
                set { topQuintile = value; }
                get { return topQuintile; }
            }

            #endregion
        }

        /// <summary>
        ///Подвыборка
        /// </summary>
        private List<SubSample> subSampleList;

        /// <summary>
        /// p-значение
        /// </summary>
        private double pValue;

        #endregion

        #region Конструкторы

        /// <summary>
        /// Обычный конструктор
        /// </summary>
        public Sample()
        {
            groupFact = "";
            nameSign = "";
            subSampleList = new List<SubSample>();
            pValue = 0;
        }

        #endregion

        #region Методы

        /// <summary>
        /// Разбиение на подвыборки признаков согласно группирующим элементам
        /// </summary>
        /// <param name="_data">Считанные данные</param>
        /// <param name="_groupIndex">Индексы группирующих элементов</param>
        /// <param name="_signIndex">Индексы выбранных признаков</param>
        /// <returns>Список разбитых выборок</returns>
        public static List<Sample> GetSample(Data _data, List<int> _groupIndex, List<int> _signIndex)
        {
            //Список индексов NA
            List<int> naList = new List<int>();
            //Выборка без пропущенных значений
            List<Sample> resList = new List<Sample>();

            //Пробежимся по всем данным и составим список индексов с пропущенными значениями (NA)
            //Признаки
            List<string> patient = new List<string>();
            int index = 0;
            for (int i = 0; i < _signIndex.Count; i++)
            {
                index = _signIndex[i];
                for (int j = 0; j < _data.PatientsCount; j++)
                {
                    patient = _data.TakePatientAtIndex(j);
                    if (patient[index] == "NA")
                    {
                        if (naList.IndexOf(j) == -1)
                            naList.Add(j);
                    }
                }
            }
            //Группирующие элементы
            for (int i = 0; i < _groupIndex.Count; i++)
            {
                index = _groupIndex[i];
                for (int j = 0; j < _data.PatientsCount; j++)
                {
                    patient = _data.TakePatientAtIndex(j);
                    if (patient[index] == "NA")
                    {
                        if (naList.IndexOf(j) == -1)
                            naList.Add(j);
                    }
                }
            }

            /*
             * Теперь сформируем выходные данные
             * Выделим все уникальные значения у группирующего фактора
             * И для каждого уникального значения сформируем свою подвыборку
             */
            //Список данных одной выборки
            List<SubSample> oneSample;
            Sample S;
            SubSample sS;
            string tmp;
            string UniqVal = "";
            for (int i = 0; i < _groupIndex.Count; i++)
            {
                //Получим список уникальных значений группирующщего фактора
                List<string> unVal = UniqueValues(_data.Patients, _groupIndex[i]);

                for (int j = 0; j < _signIndex.Count; j++)
                {
                    oneSample = new List<SubSample>();

                    //Данные выборки
                    S = new Sample();
                    S.GroupFact = _data.TakeVariableNameAtIndex(_groupIndex[i]);
                    S.NameSign = _data.TakeVariableNameAtIndex(_signIndex[j]);
                    S.SubSampleList = new List<SubSample>();

                    for (int l = 0; l < unVal.Count; l++)
                    {
                        //Данные подвыборки
                        sS = new SubSample();
                        sS.UniqueVal = unVal[l];
                        sS.SampleList = new List<double>();

                        //Заполнение подвыборки значениями
                        for (int k = 0; k < _data.PatientsCount; k++)
                        {
                            if (naList.IndexOf(k) == -1)
                            {
                                patient = _data.TakePatientAtIndex(k);
                                UniqVal = patient[_groupIndex[i]];
                                if (UniqVal == sS.UniqueVal)
                                {
                                    tmp = patient[_signIndex[j]];
                                    tmp = tmp.Replace('.', ',');
                                    sS.SampleList.Add(Convert.ToDouble(tmp));
                                }
                            }
                        }
                        sS.AverageValue = Math.Round(sS.SampleList.Average(), 1);
                        oneSample.Add(sS);
                    }
                    S.SubSampleList = oneSample;
                    resList.Add(S);
                }
            }
            return resList;
        }

        /// <summary>
        /// Составляем список уникальных значений
        /// </summary>
        /// <param name="_index">Индекс группирующего элемента</param>
        /// <returns></returns>
        private static List<string> UniqueValues(List<List<string>> _patients, int _index)
        {
            List<string> res = new List<string>();
            string data = "";
            for (int i = 0; i < _patients.Count; i++)
            {
                List<string> patient = _patients[i];
                data = patient[_index];
                if (data != "NA" && res.IndexOf(data) == -1)
                {
                    res.Add(data);
                }
            }
            return res;
        }



        #endregion

        #region Свойства

        /// <summary>
        /// Группирующий фактор
        /// </summary>
        public string GroupFact
        {
            set { groupFact = value; }
            get { return groupFact; }
        }

        /// <summary>
        ///Название призанка
        /// </summary>
        public string NameSign
        {
            set { nameSign = value; }
            get { return nameSign; }
        }

        /// <summary>
        ///Подвыборка
        /// </summary>
        public List<SubSample> SubSampleList
        {
            set { subSampleList = value; }
            get { return subSampleList; }
        }

        /// <summary>
        /// p-значение
        /// </summary>
        public double PValue
        {
            set { pValue = value; }
            get { return pValue; }
        }

        #endregion
    }
}
