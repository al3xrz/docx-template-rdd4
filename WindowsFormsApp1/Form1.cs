using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TemplateEngine.Docx;


namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }



        private void button1_Click(object sender, EventArgs e)
        {
            Договор_ФОМС_WORD df = new Договор_ФОМС_WORD();
            df.номер_договора = "1";
            df.дата_договора = "«30» сентября 2019";
            df.фио_пациента = "<Имя Пациентки>";
            df.дата_рождения_пациента = "01.01.1901";
            df.адрес_пациента = "<Тестовый адрес пациента>";
            df.номер_страхового = "<111-111111111>";
            df.паспортные_данные = "<код 8201 серия 111111>";
            df.телефон_пациента = "<89281111111>";

            df.template_file = "templates\\договор_доп_услуг_ФОМС.docx";
            df.output_file = "output\\Договор_№1_ФОМС.docx";

            df.CreateFile();




            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Договор_платный_WORD df = new Договор_платный_WORD();
            df.template_file = "templates\\договор_об_оказании_платных_ услуг.docx";
            df.output_file = "output\\Договор_№1_платный.docx";

            df.CreateFile();


        }

        private void button3_Click(object sender, EventArgs e)
        {
            Акт_платный_WORD af = new Акт_платный_WORD();

            List<Выполненная_услуга> услуги = new List<Выполненная_услуга>();
            for (int i = 0; i <= 4; i++)
            {
                Выполненная_услуга у = new Выполненная_услуга();
                у.наименование = $"Услуга{i}";
                у.дата = new DateTime().Date;
                у.комментарии = $"Комментарий {i}";
                услуги.Add(у);
            }


            af.номеракта = "№1";
            af.потребитель = "[ФИО пациента]";
            af.услуги = услуги;

            af.template_file = "templates\\акт_об_оказании_платных_услуг.docx";
            af.output_file = "output\\акт_№1_платный.docx";

            af.CreateFile();
        }
    }

}
    
