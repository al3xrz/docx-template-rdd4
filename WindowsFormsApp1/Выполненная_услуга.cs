using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    class Выполненная_услуга
    {
        public int id = 0;
        public DateTime дата = new DateTime();
        public int id_договора = 0;
        public string наименование = "";
        public int id_услуги = 0;
        public int количество = 0;
        public string ед_изм = "";
        public double цена = 0.0d;
        public double сумма = 0.0d;
        public bool удален = false;
        public string комментарии = "";
        public int скидка = 0;
        public double цена_со_скидкой = 0.0d;

    }
}
