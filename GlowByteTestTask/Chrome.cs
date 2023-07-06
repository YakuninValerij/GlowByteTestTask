using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Threading;

namespace GlowByteTestTask
{
    /// <summary>
    /// Объект для работы с html-страницами в браузере Chrome
    /// </summary>
    public class Chrome
    {
        /// <summary>
        /// Каталог размещения фала chromedriver.exe
        /// </summary>
        public static string ChromeDriverPath = "c:\\Robot";

        /// <summary>
        /// Время ожидания отклика браузера
        /// </summary>
        public int Timeout = 3000;

        /// <summary>
        /// Каталог размещения скриншотов
        /// </summary>
        public string ScreenshotPath = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) + "\\Screenshots";

        /// <summary>
        /// Объект для взаимодействия с браузером Chrome
        /// </summary>
        public ChromeDriver Browser = null;

        /// <summary>
        /// Селекторы элементов
        /// </summary>
        public enum Node
        {
            Tag,
            Class,
            Id,
            XPath
        }

        /// <summary>
        /// Создание экземпляра объекта, открытие браузера и переход по указаному адресу
        /// </summary>
        public Chrome(string url = "", bool visible = true)
        {
            ChromeOptions options;
            ChromeDriverService service;
            try
            {
                if (Browser is null)
                {
                    service = ChromeDriverService.CreateDefaultService(ChromeDriverPath);
                    options = new ChromeOptions();
                    options.AddArgument("--disable-gpu");
                    options.AddArgument("--log-level=3");
                    options.AddArguments("--test-type");
                    options.AddArguments("--allow-running-insecure-content");
                    options.AddArguments("disable-infobars");
                    options.AddUserProfilePreference("profile.default_content_setting_values.automatic_downloads", 1);
                    if (!visible)
                    {
                        service.HideCommandPromptWindow = true;
                        options.AddArgument("--window-position=-32000,-32000");
                    }
                    Browser = new ChromeDriver(service, options);
                }
                if (url.Length > 0)
                    Browser.Navigate().GoToUrl(url);
            }
            catch { }
        }

        /// <summary>
        /// Загрузить страницу по указаному адресу
        /// </summary>
        public bool Load(string url, int timeout = -1)
        {
            if (timeout < 0)
                timeout = Timeout;
            try
            {
                Browser.Navigate().GoToUrl(url);
                return WaitLoad(timeout);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Ожидание загрузки страницы
        /// </summary>
        public bool WaitLoad(int timeout = 1)
        {
            if (timeout < 0)
                timeout = Timeout;
            DateTime timer = DateTime.Now.AddMilliseconds(timeout);
            while (DateTime.Now < timer)
            {
                try
                {
                    if ((string)((IJavaScriptExecutor)Browser).
                        ExecuteScript("return document.readyState") == "complete")
                        return true;
                    Thread.Sleep(100);
                }
                catch
                {
                    break;
                }
            }
            return false;
        }

        /// <summary>
        /// Поиск HTML-элемента на загруженной странице
        /// </summary>
        public Elememt GetElement(Node node, string name, string text = "", int count = 1)
        {
            return GetElement(node, name, "innerHTML", text, count);
        }

        /// <summary>
        /// Поиск HTML-элемента на загруженной странице
        /// </summary>
        public Elememt GetElement(Node node, string name, string attribute, string text, int count = 1)
        {
            try
            {
                return Find(Browser.FindElements(GetBy(node, name)), attribute, text, count);
            }
            catch
            {
                return new Elememt(null);
            }
        }

        /// <summary>
        /// Сохранение скриншота окна браузера в файл
        /// </summary>
        public string SaveScreenshot(string filename = "")
        {
            if (string.IsNullOrEmpty(filename))
                filename = ScreenshotPath + "\\" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".jpg";
            try
            {
                if (!Directory.Exists(Path.GetDirectoryName(filename)))
                    Directory.CreateDirectory(Path.GetDirectoryName(filename));
                Screenshot screenshot = Browser.GetScreenshot();
                screenshot.SaveAsFile(filename);
            }
            catch
            {
                return string.Empty;
            }
            if (File.Exists(filename))
                return filename;
            return string.Empty;
        }

        /// <summary>
        /// Закрывает все открытые окна браузера Chrome
        /// </summary>
        public static void CloseAll()
        {
            try
            {
                foreach (Process process in Process.GetProcesses())
                    if (process.ProcessName == "chrome")
                        process.Kill();
            }
            catch { }
        }

        private static By GetBy(Node node, string name)
        {
            switch (node)
            {
                case Node.Tag:
                    return By.TagName(name);
                case Node.Class:
                    return By.ClassName(name);
                case Node.Id:
                    return By.Id(name);
                case Node.XPath:
                    return By.XPath(name);
                default:
                    return null;
            }
        }

        private static Elememt Find(ReadOnlyCollection<IWebElement> collection,
            string attribute, string text, int count = 1)
        {
            if (collection is null)
                return new Elememt(null);
            try
            {
                if (string.IsNullOrEmpty(text))
                {
                    foreach (IWebElement element in collection)
                        if (--count == 0)
                            return new Elememt(element);
                }
                else
                {
                    foreach (IWebElement element in collection)
                        if (element.GetAttribute(attribute).Trim() == text)
                            if (--count == 0)
                                return new Elememt(element);
                    foreach (IWebElement element in collection)
                        if (element.GetAttribute(attribute).IndexOf(text) >= 0)
                            if (--count == 0)
                                return new Elememt(element);
                }
            }
            catch { }
            return new Elememt(null);
        }

        /// <summary>
        /// Объект-элемент HTML
        /// </summary>
        public class Elememt
        {
            private IWebElement WebElement;

            public Elememt(IWebElement webElement)
            {
                WebElement = webElement;
            } 

            /// <summary>
            /// Клик по элементу
            /// </summary>
            public bool Click(int timeout = 0)
            {
                if (WebElement is null)
                    return false;
                DateTime timer = DateTime.Now.AddMilliseconds(timeout);
                while (true)
                {
                    try
                    {
                        WebElement.Click();
                        return true;
                    }
                    catch
                    {
                        if (DateTime.Now > timer)
                            break;
                    }
                    Thread.Sleep(250);
                }
                return false;
            }

            /// <summary>
            /// Ввод текста в элемент
            /// </summary>
            public void Send(string keys)
            {
                if (WebElement is null)
                    return;
                try
                {
                    WebElement.SendKeys(keys);
                }
                catch { }
            }
        }
    }
}
