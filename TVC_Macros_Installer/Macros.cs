using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace TVC_Macros_Installer
{
    class Macros
    {
        private string _Descr;
        private string Code { get; set; }
        public string FileName { get; private set; }
        /// <summary>
        /// Имя макроса
        /// </summary>
        public string Name { get; private set; }
        /// <summary>
        /// Описание макроса
        /// </summary>
        public string Description
        {
            get { return _Descr; }
            set
            {   
                //Удаление текущего описание
                _Descr = string.Empty;                
                foreach (var s in value.Split('\n')) { _Descr +=$"\'{s.Replace("\r", string.Empty)}\n"; }
                using (var sw = new StreamWriter(this.FileName))
                {
                    sw.Write(_Descr);
                    sw.WriteLine(this.Code);
                    sw.Close();
                }
            }
        }
        /// <summary>
        /// Создание экземпляра класса
        /// </summary>
        /// <param name="filename">Путь к файлу</param>
        /// <param name="name">Имя файла</param>
        public Macros(string filename,string name)
        {            
            File.Copy(filename, $"{Directory.GetCurrentDirectory()}\\Macros\\{filename.Split('\\')[filename.Split('\\').Length-1]}");
            Constructor(name);
        }
        /// <summary>
        /// Создание экземпляра класса
        /// </summary>
        /// <param name="name">Имя файла в папке программы</param>
        public Macros(string name)
        {
            Constructor(name);    
        }
        /// <summary>
        /// Конструктор класса
        /// </summary>
        /// <param name="name"></param>
        private void Constructor(string name)
        {
            this.Name = name;
            this.FileName = $"{Directory.GetCurrentDirectory()}\\Macros\\{name}";
            using (var sr = new StreamReader(this.FileName, Encoding.UTF8))
            {
                while (sr.Peek() == '\'') { sr.Read(); _Descr += sr.ReadLine() + System.Environment.NewLine; }
                while (!sr.EndOfStream) { this.Code += sr.ReadLine() + Environment.NewLine; }
                sr.Close();
            }
        }
        public override string ToString()
        {
            return this.Name;
        }
        /// <summary>
        /// Добавляет макрос в персональную книгу
        /// </summary>
        public void AddMacro(Excel.Workbook wb)
        {            
            wb.VBProject.VBComponents.Import(FileName);
            wb.Save();
        }
        /// <summary>
        /// Удаляет файл макроса из папки программы
        /// </summary>
        public void Delete()
        {
            File.Delete(this.FileName);
        }
        /// <summary>
        /// Закрытие все процессов Excel
        /// </summary>
    }
}
