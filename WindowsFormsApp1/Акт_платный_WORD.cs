using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TemplateEngine.Docx;

namespace WindowsFormsApp1
{
    class Акт_платный_WORD
    {
        public string номеракта = "";
        public string дд = "";
        public string мм = "";
        public string гггг = "";
        public string дд1 = "";
        public string мм1 = "";
        public string гггг1 = "";
        public string дд2 = "";
        public string мм2 = "";
        public string гггг2 = "";
        public string номердоговора = "";
        public string потребитель = "ФИО";
        public string иитого = "";
        public string предоплата = "";
        public string к_оплате = "";
        public string нндс = "";
        public string ввсего = "";
        public string всегонасумму = "";
        public string всегондс = "";
        public string фио1 = "";
        public string фио2 = "";
        public string адрес1 = "";
        public string адрес2 = "";
        public string адрес3 = "";
        public string адрес4 = "";
        public string адрес5 = "";
        public string адрес6 = "";
        public string паспорт1 = "";
        public string паспорт2 = "";
        public string паспорт3 = "";
        public string паспорт4 = "";
        public List<Выполненная_услуга> услуги = new List<Выполненная_услуга>();
        public string телефон = "";

        public string template_file = "";
        public string output_file = "";

        private TableContent CreateTableContent(List<Выполненная_услуга> услуги)
        {
            TableContent tc = new TableContent("_строка");
            int counter = 0;
            foreach (Выполненная_услуга у in услуги)
            {

                tc.AddRow(new FieldContent("№", (++counter).ToString()),
                            new FieldContent("наименование", у.наименование),
                            new FieldContent("количество", у.количество.ToString()),
                            new FieldContent("едизм", у.ед_изм),
                            new FieldContent("цена", у.цена.ToString()),
                            new FieldContent("сумма", у.сумма.ToString()));


            }

            return tc;
            
        }

        public void CreateFile()
        {
            File.Delete(output_file);
            File.Copy(template_file, output_file);



            var valuesToFill = new Content(
                new FieldContent("номеракта", номеракта),
                new FieldContent("дд", дд),
                new FieldContent("мм", мм),
                new FieldContent("гггг", гггг),
                new FieldContent("дд1", дд1),
                new FieldContent("мм1", мм1),
                new FieldContent("гггг1", гггг1),
                new FieldContent("дд2", дд2),
                new FieldContent("мм2", мм2),
                new FieldContent("гггг2", гггг2),
                new FieldContent("номердоговора", номердоговора),
                new FieldContent("потребитель", потребитель),
                new FieldContent("иитого", иитого),
                //new FieldContent("предоплата", "предоплата"),
                //new FieldContent("к_оплате", "к_оплате")//,
                new FieldContent("нндс", нндс),
                new FieldContent("ввсего", ввсего),
                new FieldContent("всегонасумму", всегонасумму),
                new FieldContent("всегондс", всегондс),
                new FieldContent("фио1", фио1),
                new FieldContent("фио2", фио2),
                new FieldContent("адрес1", адрес1),
                new FieldContent("адрес2", адрес2),
                new FieldContent("адрес3", адрес3),
                new FieldContent("адрес4", адрес4),
                new FieldContent("адрес5", адрес5),
                new FieldContent("адрес6", адрес6),
                new FieldContent("паспорт1", паспорт1),
                new FieldContent("паспорт2", паспорт2),
                new FieldContent("паспорт3", паспорт3),
                new FieldContent("паспорт4", паспорт4),
                new FieldContent("телефон", телефон),
                CreateTableContent(this.услуги)
                );

            
            using (var outputDocument = new TemplateProcessor(output_file).SetRemoveContentControls(true))
            {
                outputDocument.FillContent(valuesToFill);
                outputDocument.SaveChanges();
            }
        }




    }



}
