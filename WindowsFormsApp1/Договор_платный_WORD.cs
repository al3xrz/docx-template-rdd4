using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TemplateEngine.Docx;

namespace WindowsFormsApp1
{
    class Договор_платный_WORD
    {
        public string _номер_договора = "";
        public string ЧИСЛО = "";
        public string МЕСЯЦ = "";
        public string ГОД = "";

        public string _предоплата1 = "";
        public string _предоплата1_прописью = "";

        public string _предоплата2 = "";
        public string _предоплата2_прописью = "";


        public string _фио_потребителя = "";
        public string _сл = "";
        public string _рр = "";

        public string _ппсл = "";
        public string _ппрр = "";
        public string _потребитель_1_строка_ = "";
        public string _потребитель_2_строка_ = "";
        public string дд = "";
        public string мм = "";
        public string гггг = "";


        public string дд1 = "";
        public string мм1 = "";
        public string гггг1 = "";


        public string _место_регистрации_пациента_1 = "";
        public string _место_регистрации_пациента_2 = "";
        public string _место_регистрации_пациента_3 = "";
        public string _адрес_проживания_пациента_1 = "";
        public string _адрес_проживания_пациента_2 = "";
        public string _адрес_проживания_пациента_3 = "";
        public string _паспортные_данные_пациента_1 = "";
        public string _паспортные_данные_пациента_2 = "";
        public string _паспортные_данные_пациента_3 = "";
        public string _паспортные_данные_пациента_4 = "";
        public string _телефон_пациента = "";

        public string _представитель_1_строка = "";
        public string _представитель_2_строка = "";
        public string _дата_рождения_представителя = "";
        public string _место_регистрации_представителя_1 = "";
        public string _место_регистрации_представителя_2 = "";
        public string _место_регистрации_представителя_3 = "";
        public string _адрес_проживания_представителя_1 = "";
        public string _адрес_проживания_представителя_2 = "";
        public string _адрес_проживания_представителя_3 = "";
        public string _паспортные_данные_представителя_1 = "";
        public string _паспортные_данные_представителя_2 = "";
        public string _паспортные_данные_представителя_3 = "";
        public string _паспортные_данные_представителя_4 = "";
        public string _телефон_представителя = "";
        public string template_file = "";
        public string output_file = "";


        public Договор_платный_WORD()
        {

        }



        public void CreateFile()
        {
            File.Delete(output_file);
            File.Copy(template_file, output_file);

            var valuesToFill = new Content(
                new FieldContent("_номер_договора", "_номер_договора"),
                new FieldContent("ЧИСЛО", "ЧИСЛО"),
                new FieldContent("МЕСЯЦ", "МЕСЯЦ"),
                new FieldContent("ГОД", "ГОД"),
                new FieldContent("_предоплата1", "_предоплата1"),
                new FieldContent("_предоплата1_прописью", "_предоплата1_прописью"),
                new FieldContent("_предоплата2", "_предоплата2"),
                new FieldContent("_предоплата2_прописью", "_предоплата2_прописью"),
                new FieldContent("_фио_потребителя", "_фио_потребителя"),
                new FieldContent("_сл", "_сл"),
                new FieldContent("_рр", "_рр"),
                new FieldContent("_ппсл", "_ппсл"),
                new FieldContent("_ппрр", "_ппрр"),
                new FieldContent("_потребитель_1_строка_", "_потребитель_1_строка_"),
                new FieldContent("_потребитель_2_строка_", "_потребитель_2_строка_"),
                new FieldContent("дд", "дд"),
                new FieldContent("мм", "мм"),
                new FieldContent("гггг", "гггг"),
                //new FieldContent("дд1", "дд1"),
                //new FieldContent("мм1", "мм1"),
                //new FieldContent("гггг1", "гггг1"),
                new FieldContent("_место_регистрации_пациента_1", "_место_регистрации_пациента_1"),
                new FieldContent("_место_регистрации_пациента_2", "_место_регистрации_пациента_2"),
                new FieldContent("_место_регистрации_пациента_3", "_место_регистрации_пациента_3"),
                new FieldContent("_адрес_проживания_пациента_1", "_адрес_проживания_пациента_1"),
                new FieldContent("_адрес_проживания_пациента_2", "_адрес_проживания_пациента_2"),
                new FieldContent("_адрес_проживания_пациента_3", "_адрес_проживания_пациента_3"),
                new FieldContent("_паспортные_данные_пациента_1", "_паспортные_данные_пациента_1"),
                new FieldContent("_паспортные_данные_пациента_2", "_паспортные_данные_пациента_2"),
                new FieldContent("_паспортные_данные_пациента_3", "_паспортные_данные_пациента_3"),
                new FieldContent("_паспортные_данные_пациента_4", "_паспортные_данные_пациента_4"),
                
                
                new FieldContent("_телефон_пациента", "_телефон_пациента"),
                new FieldContent("_представитель_1_строка", "_представитель_1_строка"),
                new FieldContent("_представитель_2_строка", "_представитель_2_строка"),
                new FieldContent("_дата_рождения_представителя", "_дата_рождения_представителя"),
                new FieldContent("_место_регистрации_представителя_1", "_место_регистрации_представителя_1"),
                new FieldContent("_место_регистрации_представителя_2", "_место_регистрации_представителя_2"),
                new FieldContent("_место_регистрации_представителя_3", "_место_регистрации_представителя_3"),
              
                new FieldContent("_адрес_проживания_представителя_1", "_адрес_проживания_представителя_1"),
                new FieldContent("_адрес_проживания_представителя_2", "_адрес_проживания_представителя_2"),
                new FieldContent("_адрес_проживания_представителя_3", "_адрес_проживания_представителя_3"),
                new FieldContent("_паспортные_данные_представителя_1", "_паспортные_данные_представителя_1"),
                new FieldContent("_паспортные_данные_представителя_2", "_паспортные_данные_представителя_2"),
                new FieldContent("_паспортные_данные_представителя_3", "_паспортные_данные_представителя_3"),
                new FieldContent("_паспортные_данные_представителя_4", "_паспортные_данные_представителя_4"),
              
                new FieldContent("_телефон_представителя", "_телефон_представителя")
                            );

            using (var outputDocument = new TemplateProcessor(output_file).SetRemoveContentControls(true))
            {
                outputDocument.FillContent(valuesToFill);
                outputDocument.SaveChanges();
            }


        }




    }
}
