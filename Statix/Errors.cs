using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Statix
{
    class Errors
    {
        #region Поля

        //Список ошибок
        private List<string> errorsList;

        //Количество ошибок
        private int countErrors;

        #endregion

        #region Конструкторы

        /// <summary>
        /// Обычный конструктор
        /// </summary>
        public Errors()
        {
            errorsList = new List<string>();
            countErrors = 0;
        }

        #endregion

        #region Методы

        /// <summary>
        /// Добавление ошибки в список
        /// </summary>
        /// <param name="_error">Ошибка</param>
        public void AddError(string _error)
        {
            errorsList.Add(_error);
            countErrors++;
        }

        /// <summary>
        /// Взятие ошибки по индексу
        /// </summary>
        /// <param name="_index">Индекс ошибки</param>
        /// <returns></returns>
        public string TakeErrorOnIndex(int _index)
        {
            return errorsList.ElementAt(_index);
        }

        #endregion

        #region Свойства

        /// <summary>
        /// Количество ошибок
        /// </summary>
        public int CountErrors
        {
            get { return countErrors; }
        }

        /// <summary>
        /// Список ошибок
        /// </summary>
        public List<string> ErrorsList
        {
            get { return errorsList; }
        }

        #endregion
    }
}
