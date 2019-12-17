using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using System;
using System.IO;

namespace BT_Automatic_Test_Selenium
{
    public class Program
    {
        public static void Main(string[] args)
        {
            //TestAutomationByChrome();

            //ChangeRoleOnChrome();

            TestAutomationByFireFox();
        }

        private static void TestAutomationByChrome()
        {
            IWebDriver driver = new ChromeDriver();
            int waitTime = 10;

            driver.Url = "http://localhost:8086/mantisbt/login_page.php";
            driver.Manage().Window.Maximize();
            for (int i = 1; i <= 10; i++)
            {
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(waitTime);
                driver.FindElement(By.XPath("//a[@class='back-to-login-link pull-left']")).Click();

                string username = "161260" + i.ToString();
                string email = "162360" + i.ToString() + "@gmail.com";

                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(waitTime);
                driver.FindElement(By.XPath("//form[@id='signup-form']")).Click();

                driver.FindElement(By.Id("username")).SendKeys(username);
                driver.FindElement(By.Id("email-field")).SendKeys(email + Keys.Enter);

                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(waitTime);
                driver.FindElement(By.XPath("//a[@class='width-40 btn btn-inverse bigger-110 btn-success']")).Click();
            }

            driver.Close();
        }

        private static void ChangeRoleOnChrome()
        {
            IWebDriver driver = new ChromeDriver();
            int waitTime = 10;

            driver.Url = "http://localhost:8086/mantisbt/login_page.php";
            driver.Manage().Window.Maximize();

            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(waitTime);
            driver.FindElement(By.Id("username")).SendKeys("administrator" + Keys.Enter);

            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(waitTime);
            driver.FindElement(By.Id("password")).SendKeys("1234" + Keys.Enter);
            
 
            for (int i = 1; i <= 10; i++)
            {
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(waitTime);
                driver.FindElement(By.XPath("//a[contains(text(),'Manage Users')]")).Click();

                string username = "161260" + i.ToString();

                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(waitTime);
                driver.FindElement(By.XPath("//input[@id='username']")).SendKeys(username + Keys.Enter);

                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(waitTime);
                driver.FindElement(By.XPath("//select[@id='edit-access-level']"));
                var element = driver.FindElement(By.Id("edit-access-level"));
                var select = new SelectElement(element);

                if (i % 2 == 0)
                {
                    select.SelectByText("manager");
                }
                else
                {
                    select.SelectByText("developer");
                }

                driver.FindElement(By.XPath("//form[@id='edit-user-form']//input[@class='btn btn-primary btn-white btn-round']")).Click();
            }

            driver.Close();
        }

        private static void TestAutomationByFireFox()
        {
            IWebDriver driver = new FirefoxDriver();

            int waitTime = 10;
            driver.Url = "http://localhost:8086/mantisbt/login_page.php";
            driver.Manage().Window.Maximize();

            // mở file excel
            var package = new ExcelPackage(new FileInfo(@"C:\Users\Admin\Desktop\test_data.xlsx"));

            // lấy ra sheet đầu tiên để thao tác
            ExcelWorksheet workSheet = package.Workbook.Worksheets[1];

            int rowIndex = 1;
            while (rowIndex <= 10)
            {
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(waitTime);
                driver.FindElement(By.XPath("//a[@class='back-to-login-link pull-left']")).Click();

                // vì ta dùng 3 cột A, B, C => data của ta sẽ như sau
                var username = workSheet.Cells[rowIndex, 1].Value.ToString();
                var email = workSheet.Cells[rowIndex, 2].Value.ToString();

                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(waitTime);
                driver.FindElement(By.XPath("//form[@id='signup-form']")).Click();

                driver.FindElement(By.Id("username")).SendKeys(username);
                driver.FindElement(By.Id("email-field")).SendKeys(email + Keys.Enter);

                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(waitTime);
                driver.FindElement(By.XPath("//a[@class='width-40 btn btn-inverse bigger-110 btn-success']")).Click();

                // Xuất ra thông tin lên màn hình
                Console.WriteLine("Username: {0} | Email: {1}", username, email);

                // tăng index khi lấy xong
                rowIndex++;
            }
            driver.Close();
        }
    }
}
