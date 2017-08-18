using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
namespace Process_6_v1._0
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            List<States> states = new List<States>();
            Console.Write("System starting up");
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            try
            {
                string line = null;
                System.IO.StreamReader file =
                   new System.IO.StreamReader(basePath + "States.csv");
                while ((line = file.ReadLine()) != null)
                {
                    string[] lines = line.Split(',');
                    States a = new States();
                    a.state = lines[0];
                    a.code = lines[1];
                    states.Add(a);
                }

                file.Close();

            }
            catch
            {
                Console.WriteLine("Reading state file failed...");
                goto step00;
            }
            for (int i = 0; i < 3; i++)
            {
                System.Threading.Thread.Sleep(1000);
                Console.Write(".");
            }
            System.Threading.Thread.Sleep(1000);

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("\nProcess-6 v1.0 running now...!");
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine("System will process for");
            Console.WriteLine("\t Client : 000377 - Miller Brothers Staffing - 1st");
            Console.WriteLine("\t Client : 041001 - Access Labor - 2nd");
            Console.WriteLine("\t Client : 011164 - Advantage Logistics - 3rd");
            Console.WriteLine("\t Client : 000375 - Total Workforce-4th");
            Console.WriteLine("on following Times :\n\tMonday @11:59PM\n\t Wednesdays @1:00AM \n\t Friday @11:59PM");
            Console.WriteLine("\n\n\n");
            Console.ForegroundColor = ConsoleColor.White;
            var chromeOptions = new ChromeOptions();
            chromeOptions.AddUserProfilePreference("download.default_directory", basePath);
            chromeOptions.AddUserProfilePreference("intl.accept_languages", "nl");
            chromeOptions.AddUserProfilePreference("disable-popup-blocking", "true");
            string[] clients = { "000377" ,"041001","011164","000375"};
            string filename = null;
            int id = 0;
            string eid = "21";
            wait:
            Console.WriteLine("System on wait...");
            while (true)
            {
                if (DateTime.Now.DayOfWeek.ToString().ToUpper() == "MONDAY" || DateTime.Now.DayOfWeek.ToString().ToUpper() == "WEDNESDAY" || DateTime.Now.DayOfWeek.ToString().ToUpper() == "FRIDAY")
                {
                    string time = DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString();
                    if (time == "100" && DateTime.Now.DayOfWeek.ToString().ToUpper() == "WEDNESDAY")
                    {
                        break;
                    }
                    else
                        if (DateTime.Now.DayOfWeek.ToString().ToUpper() != "WEDNESDAY" && time == "2359")
                    {
                        break;
                    }
                }
            }
            Console.WriteLine("wait ended...!");
            IWebDriver gc = new ChromeDriver(chromeOptions);
            Start:
            //..login in to prism system.
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("System processin for Client#" + clients[id] + " on " + eid);
            
            try
            {
                gc.Navigate().GoToUrl("https://ctw.prismhr.com/ctw/dbweb.asp?dbcgm=1");
                System.Threading.Thread.Sleep(2000);
                gc.FindElement(By.XPath("//*[@id='text4v1']")).SendKeys("lightbot");
                gc.FindElement(By.XPath("//*[@id='password6v1']")).SendKeys("RPAuser1!" + Keys.Enter);
                System.Threading.Thread.Sleep(1000);
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Magenta;
                Console.WriteLine("Login to prism HR failed...!");
                Console.ForegroundColor = ConsoleColor.White;
            }

            //switch client account
            try
            {
                System.Threading.Thread.Sleep(1000);
                gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("c:");
                string check = "c:"+clients[id];
                foreach (char i in check)
                {
                    System.Threading.Thread.Sleep(1500);
                    gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys(i.ToString());
                }
                System.Threading.Thread.Sleep(3500);
                gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys(Keys.Enter);
            }
            catch
            {

                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Switicing client account failed..!");
                Console.ForegroundColor = ConsoleColor.White;
                goto step0;
            }
            Mid:
            //going to desired report
            try
            {
                gc.Navigate().GoToUrl("https://informer4w15c1.prismhr.com/ctw/Informer.html?locale=en_US#action=ReportRun&reportId=9666560&launch=false");
                System.Threading.Thread.Sleep(1000);
                gc.FindElement(By.XPath("//*[@id='LoginHeader']/center/div[2]/table/tbody/tr/td/input[1]")).SendKeys("lightbot");
                gc.FindElement(By.XPath("//*[@id='LoginHeader']/center/div[2]/table/tbody/tr/td/input[2]")).SendKeys("RPAuser1!" + Keys.Enter);
                System.Threading.Thread.Sleep(3000);
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                //Console.WriteLine("Report login failed..!");
                Console.ForegroundColor = ConsoleColor.White;
            }

            //generating the report
            try
            {
                System.Threading.Thread.Sleep(2000);
                gc.FindElement(By.ClassName("gwt-TextBox")).SendKeys(clients[id]);
                gc.FindElements(By.ClassName("gwt-TextBox"))[1].SendKeys(eid);
                gc.FindElement(By.XPath("/html/body/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td/table/tbody/tr/td[2]/table/tbody/tr/td/div/div/div[4]/div/div/table[2]/tbody/tr/td/div/div/table/tbody/tr/td")).Click();
                System.Threading.Thread.Sleep(10000);
            }
            catch
            {

                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Report Generation failed..!");
                Console.ForegroundColor = ConsoleColor.White;
                goto step0;
            }
            //exporting the data into csv
            try
            {
                gc.FindElement(By.XPath("/html/body/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/table/tbody/tr[1]/td/table/tbody/tr/td[3]/div/table/tbody/tr/td[1]/div/table/tbody/tr[2]/td/div")).Click();
                gc.FindElement(By.XPath("/html/body/div[3]/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[1]/td/div/div[1]/div[1]/table/tbody/tr[1]/td/div/div/table/tbody/tr/td[2]")).Click();
                filename = clients[id]+"_"+eid + "_" + DateTime.Now.Date.ToString("MMMM dd, yyyy");
                gc.FindElement(By.XPath("/html/body/div[3]/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[1]/td/div/div[1]/div[2]/table/tbody/tr[1]/td[2]/input")).SendKeys(Keys.Control + "a" + Keys.Control + filename);

                gc.FindElement(By.XPath("/html/body/div[3]/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/div/div")).Click();
                System.Threading.Thread.Sleep(1000);
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Report Downloading failed..!");
                Console.ForegroundColor = ConsoleColor.White;
                goto step0;
            }

            //...load the downloaded data back to system.


            yz://reading the state file
            

            //updating the data
            try
            {
                gc.Navigate().GoToUrl("https://ctw.prismhr.com/ctw/dbweb.asp?dbcgm=1");
                try
                {
                    System.Threading.Thread.Sleep(2000);
                    gc.FindElement(By.XPath("//*[@id='text4v1']")).SendKeys("lightbot");
                    gc.FindElement(By.XPath("//*[@id='password6v1']")).SendKeys("RPAuser1!" + Keys.Enter);
                    System.Threading.Thread.Sleep(1000);

                }
                catch
                {

                //does nothing
                }
                gc.FindElement(By.XPath("//*[@id='tdxspan3v1x1z1']")).Click();
                System.Threading.Thread.Sleep(200);
                gc.FindElement(By.XPath("//*[@id='tdxspan15v1x26z1']")).Click();
                System.Threading.Thread.Sleep(200);
                
            }
            catch
            {
                Console.WriteLine("Opening Employee reassignment failed..!");
            }
            //inputing data
            try
            {
                filename = clients[id]+"_"+eid + "_" + DateTime.Now.Date.ToString("MMMM dd, yyyy");

                // Read the file and display it line by line.
                int i = 0;
                string line = null;
               
                System.IO.StreamReader file =
                   new System.IO.StreamReader(basePath + filename + ".csv");
                line = file.ReadLine();
                while ((line = file.ReadLine()) != null)
                {
                    
                    line = line.Replace(", ", " ");
                    line = line.Replace('"', ' ');
                    line = line.Replace(" ", "");
                    string[] lines = line.Split(',');
                    try
                    {
                        gc.FindElement(By.XPath("//*[@id='text37v1']")).SendKeys(Keys.Backspace);
                        gc.FindElement(By.XPath("//*[@id='text37v1']")).SendKeys(Keys.Backspace);
                        gc.FindElement(By.XPath("//*[@id='text37v1']")).SendKeys(Keys.Backspace);
                        gc.FindElement(By.XPath("//*[@id='text37v1']")).SendKeys(Keys.Backspace);
                        gc.FindElement(By.XPath("//*[@id='text37v1']")).SendKeys(Keys.Backspace);
                        gc.FindElement(By.XPath("//*[@id='text37v1']")).SendKeys(Keys.Backspace);

                        Console.WriteLine("\nTrying for : " + lines[3]);
                        if(lines[12]=="NY")
                        {
                            throw new Exception();
                        }
                        foreach ( char s in lines[3])
                        {
                            gc.FindElement(By.XPath("//*[@id='text37v1']")).SendKeys(s.ToString());
                            System.Threading.Thread.Sleep(200);
                        }
                        System.Threading.Thread.Sleep(1000);
                        gc.FindElement(By.XPath("//*[@id='text37v1']")).SendKeys(Keys.Tab);
                        System.Threading.Thread.Sleep(300);
                        gc.FindElement(By.XPath("//*[@id='tdxspan41v1x1z1']")).Click();
                        System.Threading.Thread.Sleep(200);
                        States data = states.Where(s => s.state.Equals(lines[12])).Single();
                        gc.FindElement(By.XPath("//*[@id='text27v2']")).SendKeys(data.code);
                        gc.FindElement(By.XPath("//*[@id='button30v2']")).Click();
                        data = null;

                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("Success :"+lines[3] + "\t" + lines[12]);
                        Console.ForegroundColor = ConsoleColor.White;
                        i++;
                        System.Threading.Thread.Sleep(500);
                    }
                    catch
                    {
                        string current = gc.CurrentWindowHandle;
                        try
                        {
                            gc.SwitchTo().Alert().Accept();
                        }
                        catch
                        {

                        }
                        try
                        {
                            gc.SwitchTo().Window(current);
                        }
                        catch
                        {

                        }

                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Failed  :"+lines[3]);
                        Console.ForegroundColor = ConsoleColor.White;
                    }
                }
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine("Number of records processed for company " + clients[id] + " are :" + i);
                Console.ForegroundColor = ConsoleColor.White;
                file.Close();
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("File opening failed failed..!");
                Console.ForegroundColor = ConsoleColor.White;
            }

            step0:
            
            if(eid=="21")
            {
                eid = "22";
                goto Start;
            }
            
            eid = "21";
            id++;
            if(id>=4)
            {
                goto end;
            }
            goto Start;
            end:
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Process Completed..!");
            Console.ForegroundColor = ConsoleColor.White;
            goto wait;
            step00:
            Console.WriteLine("Bot terminating...!");
            Console.WriteLine("Press any keys to exit....");
            Console.ReadLine();
        }
    }
}
