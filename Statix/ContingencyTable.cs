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
        /// Переменная 1
        /// </summary>
        List<string> variable1;

        /// <summary>
        /// Переменная 2
        /// </summary>
        List<string> variable2;

        /// <summary>
        /// Количество совпадений. Пересечение переменных
        /// </summary>
        List<int> data;
        
        #endregion

        #region Конструкторы

        /// <summary>
        /// Обычный конструктор
        /// </summary>
        public ContingencyTable()
        {
            variable1 = new List<string>();
            variable2 = new List<string>();
            data = new List<int>();
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
            List<int> data = new List<int>();
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
                            //Если информация пациета совпадает с уникальными значениями, то увеличиваем счетчик
                            if (patient[_index1] == uniq1[i] && patient[_index2] == uniq2[j])
                                count++;
                        }
                    }
                    data.Add(count);
                    count = 0;
                }
            }

            //Сформируем выходные данные
            table.variable1 = uniq1;
            table.variable2 = uniq2;
            table.data = data;
            return table;
        }

        #endregion

        #region Свойства

        /// <summary>
        /// Уникальные значения первой переменной
        /// </summary>
        public List<string> Variable1
        {
            get { return variable1; }
        }

        /// <summary>
        /// Уникальные значения второй переменной
        /// </summary>
        public List<string> Variable2
        {
            get { return variable2; }
        }

        /// <summary>
        /// Данные таблицы сопряженности
        /// </summary>
        public List<int> Data
        {
            get { return data; }
        }

        #endregion
    }
}
