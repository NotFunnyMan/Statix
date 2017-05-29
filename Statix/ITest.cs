using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Statix
{
    //Интерфейс для критериев
    interface ITest
    {
        //Вычислить значение статистики
        double Execute(ContingencyTable _table);
    }
}
