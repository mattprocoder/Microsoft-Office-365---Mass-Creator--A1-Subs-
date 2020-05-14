using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Threading;
using OpenQA.Selenium.Support.UI;
using Leaf.xNet;
using System.Text.RegularExpressions;
using System.Net;
using System.Security.Cryptography;
using System.IO;

namespace MS365_Creator
{
    class Program
    {
        static string cMail = "";
        static int toWait = 0;
        static void Main(string[] args)
        {
        restart:
            Console.Clear();
            Console.Title = "MS 365 Account Creator";
            Console.Write("Waiting time for email (in s) (leave null for default): ");
            string tempAmount = Console.ReadLine();
            if (tempAmount != "") toWait = int.Parse(tempAmount) * 1000;
            else toWait = 20000;
            Console.Write("Do you want to use custom details? (Yes/No): ");
            string useCustom = Console.ReadLine().ToLower();
            if (useCustom == "") useCustom = "no";
            try
            {
                if (useCustom == "yes")
                {
                    Console.Title = "MS 365 Account Creator | Custom Mode";
                    bool usingCustom = true;
                    Console.Write("First Name: ");
                    string fName = Console.ReadLine();
                    Console.Write("Last Name: ");
                    string lName = Console.ReadLine();
                    Console.Write("Custom email (no special chars): ");
                    cMail = Console.ReadLine();
                    if (cMail == "") usingCustom = false;
                    if (lName == "") lName = "‎";
                    string finalFolder = $"Accounts/{DateTime.Now.ToString("MM.dd  H.mm")} - {fName} {lName}.txt";
                    File.AppendAllText(finalFolder, CreateAccount(fName, lName,usingCustom));
                    Colorful.Console.WriteLine($"Generating of custom account ({fName} | {lName} | {cMail}) done...",System.Drawing.Color.Yellow);
                }
                else
                {
                    Console.Title = "MS 365 Account Creator | Mass Creation";
                    Console.Write("How many accounts should I prepare?: ");
                    int toMake = int.Parse(Console.ReadLine());
                    string finalFolder = $"Accounts/{DateTime.Now.ToString("MM.dd  H.mm")} - {toMake} accounts.txt";
                    int aGenerated = 0;
                    Console.Clear();
                    Colorful.Console.WriteAscii("     Generating");
                    Colorful.Console.WriteLine($"                                            Generated [{aGenerated}/{toMake}]", System.Drawing.Color.Yellow);
                    for (int i = 0; i < toMake; i++)
                    {
                        File.AppendAllText(finalFolder, CreateAccount(GetName(), "‎",false));
                        ++aGenerated;
                        Console.Clear();
                        Colorful.Console.WriteAscii("     Generating");
                        Colorful.Console.WriteLine($"                                            Generated [{aGenerated}/{toMake}]", System.Drawing.Color.Yellow);

                    }
                }
            }
            catch(Exception e)
            {
                Console.Clear();
                Colorful.Console.Write("ERROR OCCURRED! Error message:", System.Drawing.Color.Red); Colorful.Console.Write($" {e.Message}", System.Drawing.Color.White);
                Console.ReadKey();
            }
            Console.Clear();
            Colorful.Console.Write("Do you want to continue generating? (Yes/No): ",System.Drawing.Color.Yellow);
            string output = Console.ReadLine().ToLower();
            if (output == "yes") goto restart;
            else
            {
                Console.Clear();
                Colorful.Console.Write("Created by matt", System.Drawing.Color.Lime); Colorful.Console.Write(" | ", System.Drawing.Color.Gray); Colorful.Console.Write("Press any key to exit", System.Drawing.Color.Yellow);
                Console.ReadKey();
            }
        }
        static string CreateAccount(string fName, string lName,bool customMail)
        {
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--window-size=380,700");
            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true; //Disable the logging feature
            IWebDriver driver = new ChromeDriver(service,options);
            driver.Navigate().GoToUrl("https://xkx.me");
            Thread.Sleep(2500);
            string email = "";
            IWebElement element = driver.FindElement(By.ClassName("copyable"));
            if(!customMail) email = element.GetAttribute("data-clipboard-text");
            else
            {
                element = driver.FindElement(By.Id("customShortid"));
                IJavaScriptExecutor executor = driver as IJavaScriptExecutor;
                executor.ExecuteScript("arguments[0].click();", element);
                element = driver.FindElement(By.Id("shortid"));
                element.SendKeys(cMail);
                element = driver.FindElement(By.Id("customShortid"));
                executor.ExecuteScript("arguments[0].click();", element);
                email = cMail + "@xkx.me";
            }
            ((IJavaScriptExecutor)driver).ExecuteScript("window.open();");
            driver.SwitchTo().Window(driver.WindowHandles.Last());
            driver.Navigate().GoToUrl("https://signup.microsoft.com/signup?sku=student");
            Thread.Sleep(2500);
            element = driver.FindElement(By.Id("StepsData_Email"));
            element.SendKeys(email);
            element = driver.FindElement(By.ClassName("mpl-button-next-text"));
            element.Click();
            driver.SwitchTo().Window(driver.WindowHandles.First());
            Thread.Sleep(toWait);
            element = driver.FindElement(By.Id("emaillist"));
            Match verCode = Regex.Match(element.Text, "([0-9]{6})", RegexOptions.IgnoreCase);
            driver.SwitchTo().Window(driver.WindowHandles.Last());
            element = driver.FindElement(By.Id("FirstName"));
            element.SendKeys(fName);
            element = driver.FindElement(By.Id("LastName"));
            element.SendKeys(lName);
            element = driver.FindElement(By.Id("Password"));
            string password = GeneratePassword(10);
            element.SendKeys(password);
            element = driver.FindElement(By.Id("RePassword"));
            element.SendKeys(password);
            element = driver.FindElement(By.Id("SignupCode"));
            element.SendKeys(verCode.Value);
            element = driver.FindElement(By.ClassName("mpl-button-next-text"));
            element.Click();
            driver.Quit();
            return $"Username : {email}\nPassword : {password}\n************************************\n";
        }
        static string GetName()
        {
            string[] names = { "Jennifer", "Patrick", "Carola", "Frank", "Aaron", "Willie", "Joseph" };
            Random random = new Random();
            int num = random.Next(0, names.Length-1);
            return names[num];
        }
        static string GeneratePassword(int length)
        {
            const string valid = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890";
            StringBuilder res = new StringBuilder();
            Random rnd = new Random();
            while (0 < length--)
            {
                res.Append(valid[rnd.Next(valid.Length)]);
            }
            return res.ToString();
        }
    }
}
