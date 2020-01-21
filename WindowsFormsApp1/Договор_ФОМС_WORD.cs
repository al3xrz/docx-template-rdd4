using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TemplateEngine.Docx;

namespace WindowsFormsApp1
{
    class Договор_ФОМС_WORD
    {

        public string номер_договора = "";
        public string дата_договора = "";
        public string фио_пациента = "";
        public string дата_рождения_пациента = "";
        public string адрес_пациента = "";
        public string номер_страхового = "";
        public string паспортные_данные = "";
        public string телефон_пациента = "";
        public string template_file = "";
        public string output_file = "";


        public Договор_ФОМС_WORD()
        {

        }

        public void CreateFile()
        {
            File.Delete(output_file);
            File.Copy(template_file, output_file);

            string dt = DateTime.Now.ToString("D");
            Console.WriteLine(dt);
            дата_договора = dt;

            var valuesToFill = new Content(
              new FieldContent("номер договора", номер_договора),
              new FieldContent("дата", дата_договора),
              new FieldContent("ФИО", фио_пациента),
              new FieldContent("дата рождения", дата_рождения_пациента),
              new FieldContent("адрес пациента", адрес_пациента),
              new FieldContent("номер страхового", номер_страхового),
              new FieldContent("паспортные данные", паспортные_данные),
              new FieldContent("телефон пациента", телефон_пациента)
              );

            using (var outputDocument = new TemplateProcessor(output_file).SetRemoveContentControls(true))
            {
                outputDocument.FillContent(valuesToFill);
                outputDocument.SaveChanges();
            }




        }






    }
}
