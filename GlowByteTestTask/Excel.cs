using System;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;

namespace GlowByteTestTask
{
    /// <summary>
    /// Объект для работы с файлами Excel
    /// </summary>

    public class Excel
    {
        /// <summary>
        /// Имя обрабатываемого файла
        /// </summary>
        public string Filename = string.Empty;

        /// <summary>
        /// Ссылка на объект-приложение для управления файлом
        /// </summary>
        private Application _application = null;

        /// <summary>
        /// Ссылка на активную книгу Excel
        /// </summary>
        public Workbook Workbook = null;

        /// <summary>
        /// Ссылка на активный лист Excel
        /// </summary>
        public Worksheet Worksheet = null;

        /// <summary>
        /// Создание нового объекта Excel-файла с открытием файла для записи
        /// </summary>
        public Excel(string filename, bool visible = true)
        {
            Filename = filename;
            try
            {
                _application = new Application
                {
                    Visible = visible,
                    DisplayAlerts = false
                };
                Workbook = _application.Workbooks.Open(Filename);
                Worksheet = (Worksheet)Workbook.Worksheets.get_Item(1);
            }
            catch { }
        }

        /// <summary>
        /// Чтение строки из заданной ячейки
        /// </summary>
        public string ReadCell(string addr)
        {
            if (Worksheet is null)
                return string.Empty;
            try
            {
                Microsoft.Office.Interop.Excel.Range cell = Worksheet.get_Range(addr);
                return cell.Value.ToString();
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Закрытие Excel-файла
        /// </summary>
        public void Close()
        {
            try
            {
                if (Workbook != null)
                    Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
                if (_application != null)
                    _application.Quit();
            }
            catch { }
            _application = null;
            Workbook = null;
            Worksheet = null;
        }

        /// <summary>
        /// Закрывает все открытые окна Excel
        /// </summary>
        public static void CloseAll()
        {
            try
            {
                foreach (Process process in Process.GetProcesses())
                    if (process.ProcessName == "EXCEL")
                        process.Kill();
            }
            catch { }
        }
    }
}