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

namespace Statix
{
    public partial class Form1 : MetroForm
    {
        #region Поля

        //Считанные данные
        private Data data;
        
        #endregion

        public Form1()
        {
            InitializeComponent();
            REngine.SetEnvironmentVariables();
        }

        /// <summary>
        /// Считывание данных из файла. Вывод данных для проверки. Формирование списков по типу переменных.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void metroButton7_Click(object sender, EventArgs e)
        {
            //Считывание данных из файла
            //Далее будет выбор файла
            data = new Data("ExampleCSV.csv", Encoding.Default);

            //Вывод считанной информации для проверки
            OutRedingInformation();

            //Получение ошибок, найденных при считывании
            Errors errors = data.ErrorsList();

            //Вывод сообщения о найденных ошибках
            if (errors.CountErrors != 0)
            {
                string title = "Во время считывания данных из файла были найдены ошибки!";
                string message = "Желаете посмотреть список ошибок?\n" + "\"Yes\" - Да, \"No\" - Нет.\n";
                if (MetroMessageBox.Show(this, message, title, MessageBoxButtons.YesNo, MessageBoxIcon.Error) == DialogResult.Yes)
                {
                    //Вывод списка ошибок в новом окне
                    FormForErrors Form2 = new FormForErrors(data.ErrorsList().ErrorsList);
                    Form2.Show();
                }
            }

            //signList.Clear();
            //ordinList.Clear();
            //groupSignIndex.Clear();
            //for (int i = 0; i < data.Records; i++)
            //{
            //    string tp = data.TakeScaleTypeAtIndex(i);
            //    if (tp == "Бин" || tp == "Ном")
            //        groupSignIndex.Add(i);
            //    else
            //        if (tp == "Пор")
            //        ordinList.Add(i);
            //    else
            //        signList.Add(i);
            //}
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
        

    }
}
