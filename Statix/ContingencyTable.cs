using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Statix
{
    /*
     * Класс "Таблица сопряженности"
     * Работа с таблицами сопряженности
     */

    class ContingencyTable
    {
        #region Поля

        /// <summary>
        /// Первая переменная
        /// </summary>
        string variable1;

        /// <summary>
        /// Уникальные значения первой переменной
        /// </summary>
        List<string> variable1UniqueValuesList;

        /// <summary>
        /// Вторая переменная
        /// </summary>
        string variable2;

        /// <summary>
        /// Уникальные значения второй переменной
        /// </summary>
        List<string> variable2UniqueValuesList;

        /// <summary>
        /// Количество совпадений. Пересечение переменных
        /// </summary>
        List<double> data;

        /// <summary>
        /// Количество строк в таблице
        /// </summary>
        int rowCount;

        /// <summary>
        /// Количество столбцов в таблице
        /// </summary>
        int columnCount;

        /// <summary>
        /// Число степеней свободы
        /// </summary>
        int df;
        
        #endregion

        #region Конструкторы

        /// <summary>
        /// Обычный конструктор
        /// </summary>
        public ContingencyTable()
        {
            variable1UniqueValuesList = new List<string>();
            variable2UniqueValuesList = new List<string>();
            data = new List<double>();
            rowCount = 0;
            columnCount = 0;
            df = 0;
        }

        #endregion

        #region Методы

        /// <summary>
        /// Получение списка таблиц сопряженности
        /// </summary>
        /// <param name="_data">Считанные данные</param>
        /// <param name="_listIndex">Список индексов переменных</param>
        /// <returns>Спискок таблиц сопряженности</returns>
        public static List<ContingencyTable> GetTableList(Data _data, List<int> _listIndex)
        {
            //Результирующая таблица
            List<ContingencyTable> tableList = new List<ContingencyTable>();

            //Составим список таблиц
            for (int i = 0; i < _listIndex.Count; i++)
                for (int j = i + 1; j < _listIndex.Count; j++)
                    tableList.Add(GetTable(_data, _listIndex[i], _listIndex[j]));

            return tableList;
        }

        /// <summary>
        /// Получение таблицы сопряженности
        /// </summary>
        /// <param name="_data">Считанные данные</param>
        /// <param name="_index1">Индекс первой переменной</param>
        /// <param name="_index2">Индекс второй переменной</param>
        /// <returns>Таблица сопряженности</returns>
        public static ContingencyTable GetTable(Data _data, int _index1, int _index2)
        {
            ContingencyTable table = new ContingencyTable();

            //Уникальные значения первой переменной
            List<string> uniq1 = Sample.UniqueValues(_data.Patients, _index1);
            //Уникальные значения второй переменной
            List<string> uniq2 = Sample.UniqueValues(_data.Patients, _index2);
            
            //Список индексов NA
            List<int> naList = new List<int>();
            naList = Sample.MissingList(_data, new List<int> { _index1, _index2 });

            //Список данных
            List<double> data = new List<double>();
            List<string> patient = new List<string>();
            int count = 0;

            //Идем по первой переменной
            for (int i = 0; i < uniq1.Count; i++)
            {
                //Идем по второй переменной
                for (int j = 0; j < uniq2.Count; j++)
                {
                    //Считаем совпадение первой и второй переменной
                    for (int k = 0; k < _data.Patients.Count; k++)
                    {
                        //Есть ли индекс текущего пациента в списке пропущенных значений?
                        if (naList.IndexOf(k) == -1)
                        {
                            //Информацию о пациенте
                            patient = _data.TakePatientAtIndex(k);
                            //Если информация пациента совпадает с уникальными значениями, то увеличиваем счетчик
                            if (patient[_index1] == uniq1[i] && patient[_index2] == uniq2[j])
                                count++;
                        }
                    }
                    data.Add(count);
                    count = 0;
                }
            }

            //Сформируем выходные данные
            table.variable1 = _data.TakeVariableNameAtIndex(_index1);
            table.variable1UniqueValuesList = uniq1;
            table.variable2 = _data.TakeVariableNameAtIndex(_index2);
            table.variable2UniqueValuesList = uniq2;
            table.data = data;
            table.rowCount = uniq1.Count;
            table.columnCount = uniq2.Count;
            table.df = (table.rowCount - 1) * (table.columnCount - 1);
            return table;
        }

        /// <summary>
        /// Критерий Вулфа
        /// </summary>
        /// <param name="_table">Таблица сопряженности</param>
        /// <returns></returns>
        public static double Wulf(ContingencyTable _table)
        {
            double stat = 0;
            double a = _table.Data[0];
            double b = _table.Data[1];
            double c = _table.Data[2];
            double d = _table.Data[3];
            List<double> buf = new List<double>(new double[] { a, b, c, d });
            double min = buf.Min();
            int indexMin = buf.IndexOf(min);
            if (indexMin == 0 || indexMin == 3)
                { a += 0.5; b -= 0.5; c -= 0.5; d += 0.5; }
            else
                { a -= 0.5; b += 0.5; c += 0.5; d -= 0.5; }
            
            stat = 2 * (a * Math.Log(a) + b * Math.Log(b) + c * Math.Log(c) + d * Math.Log(d) -
                       (a + b) * Math.Log(a + b) - (c + d) * Math.Log(c + d) - (a + c) * Math.Log(a + c) - (b + d) * Math.Log(b + d) +
                       (a + b + c + d) * Math.Log(a + b + c + d));

            return stat;
        }
        
        /// <summary>
        /// Критерий Хи-квадрат Пирсона
        /// </summary>
        /// <param name="_table">Таблица сопряженности</param>
        public static double Hi2Pearson(ContingencyTable _table)
        {
            //Составим списко ожидаемых частот
            List<double> waitFreq = new List<double>();
            List<double> sumInRow = new List<double>();
            List<double> sumInColumn = new List<double>();
            double sum = 0;
            
            //Сумма по строке
            for (int i = 0; i < _table.rowCount; i++)
            {
                sum = 0;
                for (int j = 0; j < _table.columnCount; j++)
                    sum += _table.Data[i * _table.columnCount + j];
                sumInRow.Add(sum);
            }
            //Сумма по столбцу
            for (int i = 0; i < _table.columnCount; i++)
            {
                sum = 0;
                for (int j = 0; j < _table.rowCount; j++)
                    sum += _table.Data[j * _table.columnCount + i];
                sumInColumn.Add(sum);
            }
            //Подсчет ожидаемых значений
            double sumAll = sumInColumn.Sum();
            for (int i = 0; i < _table.rowCount; i++)
                for (int j = 0; j < _table.columnCount; j++)
                    waitFreq.Add(Convert.ToDouble((sumInRow[i] * sumInColumn[j]) / sumAll));

            double stat = 0;
            List<double> observerFreq = _table.Data;
            for (int i = 0; i < _table.rowCount; i++)
                for (int j = 0; j < _table.columnCount; j++)
                    stat += Math.Pow(observerFreq[i * _table.columnCount + j] - waitFreq[i * _table.columnCount + j], 2) / waitFreq[i * _table.columnCount + j];
            return stat;
        }

        #endregion

        #region Свойства

        /// <summary>
        /// Первая переменная
        /// </summary>
        public string Variable1
        {
            get { return variable1; }
        }
        
        /// <summary>
        /// Уникальные значения первой переменной
        /// </summary>
        public List<string> Variable1UniqueValuesList
        {
            get { return variable1UniqueValuesList; }
        }

        /// <summary>
        /// Вторая переменная
        /// </summary>
        public string Variable2
        {
            get { return variable2; }
        }

        /// <summary>
        /// Уникальные значения второй переменной
        /// </summary>
        public List<string> Variable2UniqueValuesList
        {
            get { return variable2UniqueValuesList; }
        }

        /// <summary>
        /// Данные таблицы сопряженности
        /// </summary>
        public List<double> Data
        {
            get { return data; }
        }

        /// <summary>
        /// Количество строк в таблице
        /// </summary>
        public int RowCount
        {
            get { return rowCount; }
        }

        /// <summary>
        /// Количество столбцов в таблице
        /// </summary>
        public int ColumnCount
        {
            get { return columnCount; }
        }

        /// <summary>
        /// Число степеней свободы
        /// </summary>
        public int Df
        {
            get { return df; }
        }


        #endregion
    }
}
