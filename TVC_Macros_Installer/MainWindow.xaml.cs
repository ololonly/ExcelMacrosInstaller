using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace TVC_Macros_Installer
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        public MainWindow()
        {
            InitializeComponent();
            var MacrosDataBase = new List<Macros>();
            MacrosListBox.ItemsSource = MacrosDataBase;
            string PersonaWorkBookPath = $"{System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData)}\\Microsoft\\Excel\\XLSTART\\PERSONAL.XLSB";
            MacrosInstaller.Loaded += (s, e) =>
              {
                  //Загрузка файлов из папки программы
                  CreateMacrosDirectory();
                  GetMacrosFiles(ref MacrosDataBase);
                  MacrosListBox.Items.Refresh();
              };
            //Событие смены выделенного пункта в Listbox
            this.MacrosListBox.SelectionChanged += (s, e) => { try { MacrosDescriptionTextBox.Text = MacrosDataBase[MacrosListBox.SelectedIndex].Description; } catch { MacrosDescriptionTextBox.Text = string.Empty; } };
            //Кнопка "Добавить макрос"
            this.MacrosFileSearchButton.Click += (s, e) => { MacrosFileSearchDialog(ref MacrosDataBase); MacrosListBox.Items.Refresh(); };
            //Кнопка "Редактировать описание"
            this.EditMacrosDescriptionButton.Click += (s, e) =>
              {
                  if (MacrosListBox.SelectedIndex == -1) { MessageBox.Show("Выберите файл макроса для редактрования описания", "Внимание!", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.None); return; }
                  MacrosDescriptionTextBox.IsEnabled = true;
                  this.EditMacrosDescriptionButton.Visibility = Visibility.Collapsed;
                  this.SaveMacrosDescriptionButton.Visibility = Visibility.Visible;
              };
            //Кнопка "Сохранить"
            this.SaveMacrosDescriptionButton.Click += (s, e) =>
                {
                    if (MacrosDataBase[MacrosListBox.SelectedIndex].Description != MacrosDescriptionTextBox.Text) MacrosDataBase[MacrosListBox.SelectedIndex].Description = MacrosDescriptionTextBox.Text;
                    this.EditMacrosDescriptionButton.Visibility = Visibility.Visible;
                    this.SaveMacrosDescriptionButton.Visibility = Visibility.Collapsed;
                    this.MacrosDescriptionTextBox.IsEnabled = false;
                };
            //Кнопка "Удалить макрос"
            this.DeleteMacrosButton.Click += (s, e) =>
            {
                try
                {
                    foreach (Macros item in MacrosListBox.SelectedItems) { item.Delete(); MacrosDataBase.Remove(item); }
                    MacrosListBox.Items.Refresh();
                    MacrosListBox.SelectedIndex = 0;
                }
                catch (Exception ex) { System.Windows.MessageBox.Show(ex.Message); };
            };
            //Кнопка "Добавить в Excel"
            this.ExecuteMacrosButton.Click += (s, e) =>
            {
                if (Process.GetProcessesByName("EXCEL").Length != 0) { MessageBox.Show("Закройте все окна Excel перед импортом!", "Внимание!", MessageBoxButton.OK, MessageBoxImage.Warning); return; }
                this.IsEnabled = false;
                Excel.Workbook wb;
                ExcelWorkBookOpen(out wb, PersonaWorkBookPath);
                try { foreach (Macros item in MacrosListBox.SelectedItems) { item.AddMacro(wb); } MessageBox.Show("Success"); }
                catch (Exception ex) { System.Windows.MessageBox.Show(ex.Message); };
                ExcelClose();
                this.IsEnabled = true;
            };
            this.MacrosListBox.Drop += (s, e) =>
            {
                //
                //
                //РАСШИРЕНИЕ ФУНКЦИОНАЛА
                //
                //
            };
        }
        void ExcelWorkBookOpen(out Excel.Workbook wb, string path)
        {
            var App = new Excel.Application { Visible = false };
            wb = App.Workbooks.Open(path);
        }
        void ExcelClose()
        {
            Process[] List;
            List = Process.GetProcessesByName("EXCEL");
            foreach (var proc in List)
            {
                proc.Kill();
            }
        }

        /// <summary>
        /// Добавление макросов в коллекцию из папки программы
        /// </summary>
        /// <param name="MacrosDataBase"></param>
        void GetMacrosFiles(ref List<Macros> MacrosDataBase)
        {
            var dir = new DirectoryInfo($"{Directory.GetCurrentDirectory()}//Macros");
            foreach (var f in dir.GetFiles("*.bas"))
            {
                MacrosDataBase.Add(new Macros(f.Name));
            }
        }
        /// <summary>
        /// Создание папки для макросов
        /// </summary>
        void CreateMacrosDirectory()
        {
            Directory.CreateDirectory($"{Directory.GetCurrentDirectory()}//Macros");
        }
        /// <summary>
        /// Добавление файла макроса через проводник
        /// </summary>
        /// <param name="MacrosDataBase">Коллекция макросов</param>
        void MacrosFileSearchDialog(ref List<Macros> MacrosDataBase)
        {
            var MacrosFileDialog = new OpenFileDialog() { Filter = "Файлы макросов|*.bas", Title = "Выберите макрос" };
            if (MacrosFileDialog.ShowDialog().Value)
            {
                MacrosAdd(ref MacrosDataBase, MacrosFileDialog.FileName, MacrosFileDialog.SafeFileName);
            }
        }
        /// <summary>
        /// Добавление макроса в коллекцию
        /// </summary>
        /// <param name="MacrosDataBase">Коллекция макросов</param>
        /// <param name="name">Имя файла</param>
        /// <param name="path">Путь к файлу</param>
        void MacrosAdd(ref List<Macros> MacrosDataBase, string name, string path)
        {
            Macros mc;
            try { mc = new Macros(name, path); }
            catch (Exception ex) { System.Windows.MessageBox.Show(ex.Message); return; }
            MacrosDataBase.Add(mc);
        }

    }
    
}
