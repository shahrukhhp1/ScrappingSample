using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support;
using OpenQA.Selenium.Support.UI;

namespace WebScrapper_autorialad
{
    public partial class Form1 : Form
    {
        IWebDriver driver = null;
        // 9
        public Form1()
        {
            InitializeComponent();
            try
            {
                driver = new ChromeDriver(@"C:\SeleniumWebDriver\chromedriver");
            }
            catch
            {
                MessageBox.Show("There was an error.Program will now close");
                System.Windows.Forms.Application.Exit();
            }
        }

        private void SomeMethod()
        {
            label2.Text = "Started";
            label2.Refresh();
            driver.Navigate().GoToUrl("{your URL here}"); // could be something like http://test.com

            var company = extensionss.FindElement(driver, "//*[@id=\"inst\"]", 5);
            if (company == null)
            {
                driver.Quit();
                System.Windows.Forms.Application.Exit();
            }

            DateTime end = DateTime.Now;
            //DateTime start = DateTime.Parse(textBox1.Text);
            if (textBox2.Text == "")
            {
            }
            else
            {
                end = DateTime.Parse(textBox2.Text);
            }

            var comp = new SelectElement(company);
            var max = comp.Options.Count;
            int stat = Convert.ToInt32(textBox4.Text);
            for (int cps = stat; cps < max; cps++)
            {
                label3.Text = cps + 1 + "/" + max;
                label3.Refresh();

                comp.SelectByIndex(cps);
                var nameofComp = comp.SelectedOption.Text;
                label2.Text = "Working";
                label2.Refresh();

                var csv = new StringBuilder();


                DateTime nEnd = end;
                DateTime nStart = nEnd.AddMonths(-1).AddDays(1);

                int line = 0;
                int monthlyTry = 0;
                while (monthlyTry < 6 || nEnd > end.AddYears(-4))
                {
                    var d1 = driver.FindElement(By.XPath("//*[@id=\"feini\"]"));
                    d1.SendKeys(nStart.ToString("dd/MM/yyyy").Replace('-', '/'));

                    var d2 = driver.FindElement(By.XPath("//*[@id=\"fefin\"]"));
                    d2.SendKeys(nEnd.ToString("dd/MM/yyyy").Replace('-', '/'));

                    var butt = driver.FindElement(By.XPath("//*[@id=\"enviar\"]"));
                    butt.Click();
                    line = GetTableContents(line, ref csv);
                    if (File.Exists(textBox3.Text + "//" + GetSafeFilename(nameofComp) + ".csv") && csv.Length > 0)
                    {
                        File.AppendAllText(textBox3.Text + "//" + GetSafeFilename(nameofComp) + ".csv", csv.ToString());
                        monthlyTry = 0;
                    }
                    else if (csv.Length > 0)
                    {
                        File.WriteAllText(textBox3.Text + "//" + GetSafeFilename(nameofComp) + ".csv", csv.ToString());
                        monthlyTry = 0;
                    }
                    else if (csv.Length == 0)
                        monthlyTry += 1;

                    csv = new StringBuilder();
                    while (line != 0)
                    {
                        line = NextPageExtraction(line, ref csv);
                    }
                    if (csv.Length > 0)
                        File.AppendAllText(textBox3.Text + "//" + GetSafeFilename(nameofComp) + ".csv", csv.ToString());
                    csv = new StringBuilder();

                    driver.Navigate().GoToUrl("http://autorialadi.extra.bcv.org.ve/operadorAladi/busqfecautorizacionaladi.jsp");
                    company = extensionss.FindElement(driver, "//*[@id=\"inst\"]", 5);
                    comp = new SelectElement(company);
                    comp.SelectByIndex(cps);

                    nEnd = nEnd.AddMonths(-1);
                    nStart = nEnd.AddMonths(-1).AddDays(1);
                }

                listBox1.Items.Add(nameofComp);

                listBox1.Refresh();
                driver.Navigate().GoToUrl("http://autorialadi.extra.bcv.org.ve/operadorAladi/busqfecautorizacionaladi.jsp");
                GoWait();

                company = driver.FindElement(By.XPath("//*[@id=\"inst\"]"));
                comp = new SelectElement(company);
            }
            this.Focus();
            label2.Text = "Done";
            label2.Refresh();
            MessageBox.Show("Done");
        }

        public string GetSafeFilename(string filename)
        {

            return string.Join("_", filename.Split(Path.GetInvalidFileNameChars()));

        }

        private void GoWait()
        {
            Random r = new Random();
            int stat = Convert.ToInt32(textBox5.Text);
            int nRan = (r.Next(1, stat)) * 1000;
            label2.Text = "Waiting " + nRan / 1000 + " sec";
            label2.Refresh();
            System.Threading.Thread.Sleep(nRan);
            label2.Text = "Working";
            label2.Refresh();
        }

        private int NextPageExtraction(int line, ref StringBuilder csv)
        {
            var page = driver.FindElement(By.XPath("//*[@id=\"pagineo\"]"));
            var nn = page.GetAttribute("value");
            var pagNo = nn.Substring(nn.LastIndexOf(" ") + 1);
            if (Convert.ToInt32(pagNo) > line)
            {
                driver.FindElement(By.XPath("//*[@id=\"listDetalle\"]/div/img[2]")).Click();
                GoWait();
                if (line <= 50)
                {
                    try
                    {
                        driver.SwitchTo().Window(driver.WindowHandles[1]);
                    }
                    catch
                    {
                        GoWait();
                        driver.SwitchTo().Window(driver.WindowHandles[1]);
                    }
                }

                line = GetTableContents(line, ref csv);
                return line;
            }
            else
            {
                return 0;
            }
        }

        private int GetTableContents(int line, ref StringBuilder csv)
        {
            var table = extensionss.FindElement(driver, "/html/body/table[2]/tbody/tr[2]/td/div/table[1]/tbody", 1);
            try
            {
                if (table == null || !table.Displayed)
                {
                    for (int j = 0; j < 10; j++)
                    {
                        if (extensionss.FindElement(driver, "/html/body/table[1]/tbody/tr[3]/td[2]", 5) != null)
                        {
                            if (extensionss.FindElement(driver, "/html/body/table[1]/tbody/tr[3]/td[2]", 5).Text.Trim() == "Información no disponible")
                            {
                                table = null;
                                break;
                            }
                        }
                        else
                        {
                            table = extensionss.FindElement(driver, "/html/body/table[2]/tbody/tr[2]/td/div/table[1]/tbody", 5);
                            if (table != null && table.Displayed)
                                break;
                        }
                         
                        GoWait();
                    }
                }
            }
            catch
            {
                for (int j = 0; j < 10; j++)
                {
                    if (extensionss.FindElement(driver, "/html/body/table[1]/tbody/tr[3]/td[2]", 5) != null)
                    {
                        if (extensionss.FindElement(driver, "/html/body/table[1]/tbody/tr[3]/td[2]", 5).Text.Trim() == "Información no disponible")
                        {
                            table = null;
                            break;
                        }
                    }
                    else
                    {
                        table = extensionss.FindElement(driver, "/html/body/table[2]/tbody/tr[2]/td/div/table[1]/tbody", 5);
                        if (table != null && table.Displayed)
                            break;
                    }

                    GoWait();
                }
            }
            if (table != null)
            {
                IList<IWebElement> allRows = table.FindElements(By.TagName("tr"));
                for (int i = 2; i < allRows.Count() - 1; i++)
                {

                    var row = allRows[i];
                    IList<IWebElement> cells = row.FindElements(By.TagName("td"));
                    string cz = cells[0] != null ? cells[0].Text : "";
                    string cf = cells[1] != null ? cells[1].Text : "";
                    string cs = cells[2] != null ? cells[2].Text : "";
                    //string ct = cells[3] != null ? cells[3].Text : "";
                    //string cfo = cells[4] != null ? cells[4].Text : "";
                    //string cfi = cells[5] != null ? cells[5].Text : "";
                    //string csi = cells[6] != null ? cells[6].Text : "";
                    //string cse = cells[7] != null ? cells[7].Text : "";
                    // string ce = cells[8] != null ? cells[8].Text : "";
                    string newLine = string.Format("{0},{1},{2},{3}", cz, cf, cs, Environment.NewLine);
                    //string newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8}", cz, cf, cs, ct, cfo, cfi, csi, cse, Environment.NewLine);
                    //foreach (var cell in cells)
                    //{
                    //    newLine += string.Format("{0},", cell.Text);
                    //}
                    //newLine += string.Format("{0}", Environment.NewLine);
                    csv.Append(newLine);

                    line += 1;
                }
            }
            return line;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            driver.Quit();
            System.Windows.Forms.Application.Exit();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            SomeMethod();
        }
    }
    public static class extensionss
    {
        public static IWebElement FindElement(this IWebDriver driver, string xname, int timeoutInSeconds)
        {
            try
            {
                if (timeoutInSeconds > 0)
                {
                    var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeoutInSeconds));
                    return wait.Until(drv => drv.FindElement(By.XPath(xname)));
                }
                return driver.FindElement(By.XPath(xname));
            }
            catch
            {
                return null;
            }

        }
    }
}
