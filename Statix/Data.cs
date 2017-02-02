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

        //Количство переменных
        private int records;

        //Список ошибок
        private Errors errors; 

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
            errors = new Errors();
        }

        /// <summary>
        /// Конструктор по файлу
        /// </summary>
        /// <param name="_fileName">Имя файла</param>
        /// <param name="_encoding">Кодировка файла</param>
        public Data(string _fileName, Encoding _encoding)
        {
            //Открытие файла
            var File = new StreamReader(_fileName, _encoding);
            string line = "";

            patientsList = new List<List<string>>();
            uniqueValues = new List<string>();
            scaleType = new List<string>();
            variableName = new List<string>();
            records = 0;
            errors = new Errors();

            //Читаем уникальные значения
            line = File.ReadLine();
            AddUniqueValues(line);
            //Читаем типы шкал
            line = File.ReadLine();
            AddScaleType(line);
            //Читаем названия переменных
            line = File.ReadLine();
            AddVariableName(line);

            //Читаем пациентов
            while ((line = File.ReadLine()) != null)
                AddPatients(line);

            File.Close();

        }

        #endregion

        #region Методы
        
        /// <summary>
        ///  Добавить уникальные значения
        /// </summary>
        /// <param name="_param">Строка с данными</param>
        public void AddUniqueValues(string _param)
        {
            //Количество пропущенных значений
            int tmpMissVals = 0;

            //Разбиваем входную строку на отдельные элементы
            var param = _param.Split(';');

            //Записываем количество переменных
            records = param.Length - 1;

            int numbOfVar = param.Length;

            //Начинаем со второго значения, так как первое - название шкалы
            for (int i = 1; i < numbOfVar; i++)
                if (param[i] != "")
                    uniqueValues.Add(param[i]);
                else
                {
                    uniqueValues.Add("NA");
                    tmpMissVals++;
                }
            
            //Добавляем ошибки в список ошибок
            if (tmpMissVals != 0)
                errors.AddError("Уникальные значения. Количество пропущенных значений: " + tmpMissVals.ToString() + ".");
        }

        /// <summary>
        ///Добавить тип шкалы
        /// </summary>
        /// <param name="_param">Строка с данными</param>
        public void AddScaleType(string _param)
        {
            //Количество пропущенных значений
            int tmpMissVals = 0;

            //Разбиваем входную строку на отдельные элементы
            var param = _param.Split(';');
            int numbOfVar = param.Length;
            if (numbOfVar != records)
                tmpMissVals = records - numbOfVar + 1;

            //Начинаем со второго значения, так как первое - название шкалы
            for (int i = 1; i < param.Length; i++)
                if (param[i] != "")
                    scaleType.Add(param[i]);
                else
                {
                    scaleType.Add("NA");
                    tmpMissVals++;
                }
            
            //Добавляем ошибки в список ошибок
            if (tmpMissVals != 0)
                errors.AddError("Тип шкалы. Количество пропущенных значений: " + tmpMissVals.ToString() + ".");
        }

        /// <summary>
        ///Добавить название переменной
        /// </summary>
        /// <param name="_param">Строка с данными</param>
        public void AddVariableName(string _param)
        {
            //Количество пропущенных значений
            int tmpMissVals = 0;

            //Разбиваем входную строку на отдельные элементы
            var param = _param.Split(';');
            int numbOfVar = param.Length;
            if (numbOfVar != records)
                tmpMissVals = records - numbOfVar + 1;

            //Начинаем со второго значения, так как первое - название шкалы
            for (int i = 1; i < param.Length; i++)
                if (param[i] != "")
                    variableName.Add(param[i]);
                else
                {
                    variableName.Add("NA");
                    tmpMissVals++;
                }

            //Добавляем ошибки в список ошибок
            if (tmpMissVals != 0)
                errors.AddError("Название переменной. Количество пропущенных значений: " + tmpMissVals.ToString() + ".");
        }
        
        /// <summary>
        ///Добавить пациента
        /// </summary>
        /// <param name="_param">Строка с данными</param>
        public void AddPatients(string _param)
        {
            //Количество пропущенных значений
            int tmpMissVals = 0;

            //Разбиваем входную строку на отдельные элементы
            var param = _param.Split(';');

            //Временные переменные
            string typeScale;
            int numbOfVar = param.Length;
            if (numbOfVar != records)
                tmpMissVals = records - numbOfVar + 1;

            //Проверка входных данных на корректность в соответствии с типом шкалы
            //Начинаем со второго значения, так как первое - порядковый номер
            for (int i = 1; i < numbOfVar - 1; i++)
            {
                typeScale = TakeScaleTypeAtIndex(i - 1);
                if (typeScale == "Бин")
                    param[i] = ProcessingBinaryValue(param[i]);
                if (typeScale == "Кол")
                    param[i] = ProcessingQuantitativeValue(param[i]);
                if (typeScale == "Ном")
                    param[i] = ProcessingNominalValue(param[i]);
                if (typeScale == "Пор")
                    param[i] = ProcessingOrdinallValue(param[i]);
                if (param[i] == "NA")
                    tmpMissVals++;
            }

            //Удалим первый элемент последовательности и сдвинем все
            for (int i = 0; i < numbOfVar - 1; i++)
                param[i] = param[i + 1];
            Array.Resize(ref param, numbOfVar - 1);

            //Добавляем пациента в список пациентов
            List<string> tmp = new List<string>();
            tmp.AddRange(param);
            patientsList.Add(tmp);

            //Добавляем ошибки в список ошибок
            if (tmpMissVals != 0)
                errors.AddError("Пациент номер: " + patientsList.Count.ToString() + ". Количество пропущенных значений: " + tmpMissVals.ToString() + ".");
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
        private bool ThisNumber(string _str)
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
        private string ProcessingBinaryValue(string _str)
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
        private string ProcessingQuantitativeValue(string _str)
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
            else
                if (char.IsDigit(_str[0])) return "NA";
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
        
        /// <summary>
        /// Получение списка ошибок
        /// </summary>
        /// <returns></returns>
        public Errors ErrorsList()
        {
            return errors;
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
