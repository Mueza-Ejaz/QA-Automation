using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OfficeOpenXml;

public class LoginTest
{
    private IWebDriver _driver = null!;

    [SetUp]
    public void Setup()
    {
        _driver = new ChromeDriver();
        _driver.Manage().Window.Maximize();
        _driver.Navigate().GoToUrl("https://www.saucedemo.com/");
    }

    [TearDown]
    public void TearDown()
    {
        if (_driver != null)
        {
            _driver.Quit();
            _driver.Dispose();
        }
    }

    private void LoginWithCredentials(string username, string password)
    {
        _driver.FindElement(By.Id("user-name")).Clear();
        _driver.FindElement(By.Id("password")).Clear();
        _driver.FindElement(By.Id("user-name")).SendKeys(username);
        _driver.FindElement(By.Id("password")).SendKeys(password);
        _driver.FindElement(By.Id("login-button")).Click();
    }

    [Test]
    public void LoginFromExcel()
    {
        string path = Path.Combine(
            @"C:\Users\Yousuf Traders\OneDrive\Desktop\QA-Automation",
            "logins.xlsx");

        var rows = ReadExcel(path);

        foreach (var (user, pass) in rows)
        {
            _driver.Navigate().GoToUrl("https://www.saucedemo.com/");
            LoginWithCredentials(user, pass);

            bool success = _driver.Url.Contains("inventory")
                        || _driver.FindElements(By.CssSelector(".error-message-container")).Count > 0;

            Assert.IsTrue(success, $"Login attempt failed for user: {user}");
            Console.WriteLine($"Tried user: {user} — OK");
        }
    }

    private List<(string user, string pass)> ReadExcel(string path)
    {
        var list = new List<(string, string)>();
        ExcelPackage.License.SetNonCommercialPersonal("Student");

        using var pkg = new ExcelPackage(new FileInfo(path));
        var ws = pkg.Workbook.Worksheets[0];

        for (int i = 2; i <= ws.Dimension.Rows; i++)
        {
            var u = ws.Cells[i, 1].Text?.Trim();
            var p = ws.Cells[i, 2].Text?.Trim();

            if (!string.IsNullOrEmpty(u) && !string.IsNullOrEmpty(p))
                list.Add((u, p));
        }

        return list;
    }
}

