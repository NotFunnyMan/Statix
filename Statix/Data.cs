using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Statix
{
    class Data
    {

        #region Поля

        //Уникальные значения
        private List<string> uniqueValues;

        //Тип шкалы
        private List<string> scaleType;

        //Название переменной
        private List<string> variableName;

        //Все пациенты
        private List<List<string>> patientsList;

        //Количество записей в строке (Количство переменных)
        private int records;

        //Список ошибок
        //private List<List<string>> errorsList;

        //Временные переменные
        //Список ошибок для добавления пациентов
        //List<string> tmpList;
        //Количество пропущенных значений
        //int tmpMissVals;
        //Количество строк не удовлетворяющих длине
        //int tmpCntStrOthrLen;

        #endregion

        #region Конструкторы

        /// <summary>
        /// Конструктор
        /// </summary>
        public Data()
        {
            patientsList = new List<List<string>>();
            uniqueValues = new List<string>();
            scaleType = new List<string>();
            variableName = new List<string>();
            records = 0;
            //errorsList = new List<List<string>>();
            //tmpList = new List<string>();
            //tmpMissVals = 0;
            //tmpCntStrOthrLen = 0;
            //sample = new Sample();
        }

        /// <summary>
        /// Конструктор по файлу
        /// </summary>
        /// <param name="_fileName">Имя файла</param>
        /// <param name="_encoding">Кодировка файла</param>
        public Data(string _fileName, Encoding _encoding)
        {
            //Чтение 
            var File = new StreamReader(_fileName, _encoding);
            string line = "";

            patientsList = new List<List<string>>();
            uniqueValues = new List<string>();
            scaleType = new List<string>();
            variableName = new List<string>();
            records = 0;
            //errorsList = new List<List<string>>();
            //sample = new Sample();

            line = File.ReadLine();
            AddUniqueValues(line);
            line = File.ReadLine();
            AddScaleType(line);
            line = File.ReadLine();
            AddVariableName(line);

            //В строке первый элемент
            while ((line = File.ReadLine()) != null)
            {
                AddPatients(line);
            }
            File.Close();

            //Список ошибок для добавления пациентов
            //tmpList = new List<string>();
            //Количество пропущенных значений
            //tmpMissVals = 0;
            //Количество строк не удовлетворяющих длине
            //tmpCntStrOthrLen = 0;

            //Формирование сообщения об ошибке
            //List<string> list = new List<string>();
            //list.Add("Пациенты:");
            //if (tmpMissVals != 0)
            //{
            //    list.Add("Было найдено: " + tmpMissVals.ToString() + " пустых значений или значений, которые не удовлетворяют типу шкалы.");
            //    list.Add("Например, в шкале с типом 'Бин' стоит 2 (возможные значения 0 или 1)");
            //    list.Add("Некорректные значения заменены символом: NA.");
            //    list.Add("Проверьте входные данные.");
            //}
            //else
            //    if (tmpCntStrOthrLen != 0)
            //{
            //    list.Add("Было найдено: " + tmpCntStrOthrLen.ToString() + " строк разной длины.");
            //    list.Add("Проверьте входные данные.");
            //}
            //else
            //    list.Add("Ошибок не обнаружено.");
            //errorsList.Add(list);

        }

        #endregion

        #region Методы

        /// <summary>
        ///Добавить пациента
        /// </summary>
        /// <param name="_param">Строка с данными</param>
        public void AddPatients(string _param)
        {
            //Разбиваем входную строку на отдельные элементы
            var param = _param.Split(';');

            //Временные переменные
            string tp;
            int numbOfvar = param.Length;
            //if (numbOfvar != records)
            //    tmpCntStrOthrLen++;

            //Проверка входных данных на корректность
            for (int i = 1; i < numbOfvar - 1; i++)
            {
                tp = TakeScaleTypeAtIndex(i - 1);
                if (tp == "Бин")
                    param[i] = ProcessingBinaryValue(param[i]);
                if (tp == "Кол")
                    param[i] = ProcessingQuantitativeValue(param[i]);
                if (tp == "Ном")
                    param[i] = ProcessingNominalValue(param[i]);
                if (tp == "Пор")
                    param[i] = ProcessingOrdinallValue(param[i]);
                //if (param[i] == "NA")
                //    tmpMissVals++;
            }

            //Удалим первый элемент последовательности и сдвинем все
            for (int i = 0; i < numbOfvar - 1; i++)
                param[i] = param[i + 1];
            Array.Resize(ref param, numbOfvar - 1);

            List<string> tmp = new List<string>();
            tmp.AddRange(param);
            patientsList.Add(tmp);
        }

        /// <summary>
        ///  Добавить уникальные значения
        /// </summary>
        /// <param name="_param">Строка с данными</param>
        public void AddUniqueValues(string _param)
        {
            var param = _param.Split(';');
            int countEmpty = 0;

            for (int i = 1; i < param.Length; i++)
                if (param[i] != "")
                    uniqueValues.Add(param[i]);
                else
                {
                    uniqueValues.Add("NA");
                    countEmpty++;
                }

            //Формирование сообщения об ошибке
            //List<string> list = new List<string>();
            //list.Add("Уникальные значения:");
            //if (countEmpty != 0)
            //{
            //    list.Add("Было найдено: " + countEmpty.ToString() + " пустых значений.");
            //    list.Add("Пустые значения заменены символом: NA");
            //    list.Add("Проверьте входные данные.");
            //}
            //else
            //    list.Add("Ошибок не обнаружено.");
            //errorsList.Add(list);

            //Записываем количество записей в строке
            records = param.Length - 1;
        }

        /// <summary>
        ///Добавить тип шкалы
        /// </summary>
        /// <param name="_param">Строка с данными</param>
        public void AddScaleType(string _param)
        {
            var param = _param.Split(';');
            int countEmpty = 0;

            for (int i = 1; i < param.Length; i++)
                if (param[i] != "")
                    scaleType.Add(param[i]);
                else
                {
                    scaleType.Add("NA");
                    countEmpty++;
                }

            //Формирование сообщения об ошибке
            //List<string> list = new List<string>();
            //list.Add("Тип шкалы:");
            //if (countEmpty != 0)
            //{
            //    list.Add("Было найдено: " + countEmpty.ToString() + " пустых значений.");
            //    list.Add("Пустые значения заменены символом: NA");
            //    list.Add("Проверьте входные данные.");
            //}
            //else
            //    if (param.Length != records)
            //{
            //    list.Add("Количество значений в строке 'Тип шкалы' не совпадает с количеством значений в первой строке: " + records.ToString());
            //    list.Add("Проверьте входные данные.");
            //}
            //else
            //    list.Add("Ошибок не обнаружено.");
            //errorsList.Add(list);
        }

        /// <summary>
        ///Добавить название переменной
        /// </summary>
        /// <param name="_param">Строка с данными</param>
        public void AddVariableName(string _param)
        {
            var param = _param.Split(';');
            int countEmpty = 0;

            for (int i = 1; i < param.Length; i++)
                if (param[i] != "")
                    variableName.Add(param[i]);
                else
                {
                    variableName.Add("NA");
                    countEmpty++;
                }

            //Формирование сообщения об ошибке
            //List<string> list = new List<string>();
            //list.Add("Имя переменной:");
            //if (countEmpty != 0)
            //{
            //    list.Add("Было найдено: " + countEmpty.ToString() + " пустых значений.");
            //    list.Add("Пустые значения заменены символом: NA");
            //    list.Add("Проверьте входные данные.");
            //}
            //else
            //    if (param.Length != records)
            //{
            //    list.Add("Количество значений в строке 'Имя переменной' не совпадает с количеством значений в первой строке: " + records.ToString());
            //    list.Add("Проверьте входные данные.");
            //}
            //else
            //    list.Add("Ошибок не обнаружено.");
            //errorsList.Add(list);
        }

        /// <summary>
        /// Взятие пациента по индексу
        /// </summary>
        /// <param name="_index">Индекс</param>
        /// <returns></returns>
        public List<string> TakePatientAtIndex(int _index)
        {
            return patientsList.ElementAt(_index);
        }

        /// <summary>
        /// Взятие уникального значения по индексу
        /// </summary>
        /// <param name="_index">Индекс</param>
        /// <returns></returns>
        public string TakeUniqueValuesAtIndex(int _index)
        {
            return uniqueValues.ElementAt(_index);
        }

        /// <summary>
        /// Взятие типа шкалы по индексу
        /// </summary>
        /// <param name="_index">Индекс</param>
        /// <returns></returns>
        public string TakeScaleTypeAtIndex(int _index)
        {
            return scaleType.ElementAt(_index);
        }

        /// <summary>
        /// Взятие имени переменной по индексу
        /// </summary>
        /// <param name="_index">Индекс</param>
        /// <returns></returns>
        public string TakeVariableNameAtIndex(int _index)
        {
            return variableName.ElementAt(_index);
        }

        /// <summary>
        /// Проверка, что строка - это число
        /// </summary>
        /// <param name="_str">Строка для проверки</param>
        /// <returns></returns>
        public bool ThisNumber(string _str)
        {
            //Если какой-либо символ строки не цифра, то введено не число
            for (int i = 0; i < _str.Length; i++)
            {
                if (!char.IsDigit(_str[i]))
                {
                    if (_str[i] != '.')
                        if (_str[i] != '-')
                            return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Обработка бинарного значения
        /// </summary>
        /// <param name="_str">Строка для проверки</param>
        /// <returns></returns>
        public string ProcessingBinaryValue(string _str)
        {
            //Проверка строки на пустоту
            if (_str != "")
                //Если введено число. И это либо 0, либо 1, то все хорошо
                if (ThisNumber(_str))
                    if (_str == "0" || _str == "1")
                        return _str;
            return "NA";
        }

        /// <summary>
        /// Обработка количественного значения
        /// </summary>
        /// <param name="_str">Строка для проверки</param>
        /// <returns></returns>
        public string ProcessingQuantitativeValue(string _str)
        {
            //Проверка строки на пустоту
            if (_str != "")
                //Если введено число и оно больше нуля, то все хорошо
                if (ThisNumber(_str))
                    if (_str[0] != '-')
                        return _str;
            return "NA";
        }

        /// <summary>
        /// Обработка номанального значения
        /// </summary>
        /// <param name="_str">Строка для проверки</param>
        /// <returns></returns>
        private string ProcessingNominalValue(string _str)
        {
            if (_str == "")
                return "NA";
            return _str.ToLower();
        }

        /// <summary>
        /// Обработка порядкового значения
        /// </summary>
        /// <param name="_str">Строка для проверки</param>
        /// <returns></returns>
        private string ProcessingOrdinallValue(string _str)
        {
            return _str;
        }

        

        #endregion

        #region Поля

        /// <summary>
        /// Количество переменных
        /// </summary>
        public int Records
        {
            get { return this.records; }
        }

        /// <summary>
        /// Количество пациентов
        /// </summary>
        public int PatientsCount
        {
            get { return this.patientsList.Count; }
        }

        #endregion
    }
}
