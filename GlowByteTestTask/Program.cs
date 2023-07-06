using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Diagnostics;
using System.Reflection;

namespace GlowByteTestTask 
{
    /// <summary>
    /// Тестовое задание для GlowByte
    /// </summary>
    /// Разработчик: Якунин Валерий
    internal class Program
    {
        // Рабочие директории
        private static string RobotTempDir;
        private static string DownloadsPath;
        // Объекты для работы с Excel и браузером
        private static Chrome Browser;
        private static Excel ExcelFile;
        // Списки с данными из Excel файла
        private static List<string> FirstName = new List<string>();
        private static List<string> LastName = new List<string>();
        private static List<string> CompanyName = new List<string>();
        private static List<string> Role = new List<string>();
        private static List<string> Address = new List<string>();
        private static List<string> Email = new List<string>();
        private static List<string> PhoneNumber = new List<string>();

        static void Main(string[] args)
        {
            try
            {
                Stopwatch stopwatch = Stopwatch.StartNew();
                Console.WriteLine("Робот начал работу");
                LoadSettings();
                Browser.Load("https://www.rpachallenge.com/", 5000);
                Browser.GetElement(Chrome.Node.Tag, "a", "Download Excel").Click();
                Thread.Sleep(1000);
                ReadRecords();
                Browser.GetElement(Chrome.Node.Tag, "button", "Start").Click();
                Thread.Sleep(250);
                for (int i = 0; i < FirstName.Count; i++)
                {
                    FillData(FirstName[i], LastName[i], CompanyName[i], Role[i], Address[i], Email[i], PhoneNumber[i]);
                    Browser.GetElement(Chrome.Node.Tag, "input", "type", "submit").Click();
                }
                Browser.SaveScreenshot();
                Console.WriteLine("Скриншот сохранен по пути: " + Browser.ScreenshotPath + "\\" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".jpg");
                Chrome.CloseAll();
                Excel.CloseAll();
                stopwatch.Stop();
                Console.WriteLine("Робот завершил работу за: " + stopwatch.Elapsed.Seconds + " секунд");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        /// <summary>
        /// Загрузка настроек Робота
        /// </summary>
        static void LoadSettings()
        {
            RobotTempDir = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
            Chrome.ChromeDriverPath = RobotTempDir;
            Browser = new Chrome();
            DownloadsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),"Downloads");
        }

        /// <summary>
        /// Чтение данных из Excel файла
        /// </summary>
        static void ReadRecords()
        {
            int counter = 2;
            ExcelFile = new Excel(DownloadsPath + "\\challenge.xlsx", false);
            while (true)
            {
                if (ExcelFile.ReadCell("A" + counter).Length > 1)
                {
                    FirstName.Add(ExcelFile.ReadCell("A" + counter));
                    LastName.Add(ExcelFile.ReadCell("B" + counter));
                    CompanyName.Add(ExcelFile.ReadCell("C" + counter));
                    Role.Add(ExcelFile.ReadCell("D" + counter));
                    Address.Add(ExcelFile.ReadCell("E" + counter));
                    Email.Add(ExcelFile.ReadCell("F" + counter));
                    PhoneNumber.Add(ExcelFile.ReadCell("G" + counter));
                    counter++;
                    continue;
                }
                else
                {
                    break;
                }
            }
            ExcelFile.Close();
        }

        /// <summary>
        /// Завполнение данных в браузере
        /// </summary>
        static void FillData(string FirstName, string LastName, string CompanyName, string Role, string Address, string Email, string PhoneNumber)
        {
            Browser.GetElement(Chrome.Node.Tag, "input", "ng-reflect-name", "labelFirstName").Send(FirstName);
            Browser.GetElement(Chrome.Node.Tag, "input", "ng-reflect-name", "labelLastName").Send(LastName);
            Browser.GetElement(Chrome.Node.Tag, "input", "ng-reflect-name", "labelCompanyName").Send(CompanyName);
            Browser.GetElement(Chrome.Node.Tag, "input", "ng-reflect-name", "labelRole").Send(Role);
            Browser.GetElement(Chrome.Node.Tag, "input", "ng-reflect-name", "labelAddress").Send(Address);
            Browser.GetElement(Chrome.Node.Tag, "input", "ng-reflect-name", "labelEmail").Send(Email);
            Browser.GetElement(Chrome.Node.Tag, "input", "ng-reflect-name", "labelPhone").Send(PhoneNumber);
        }
    }
}