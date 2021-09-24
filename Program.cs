using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;

namespace rediff_minni_project
{
	class Program
	{
		static void Main(string[] args)
		{	
			
			//Lauch Chrome
			IWebDriver driver = new ChromeDriver("C:\\Users\\mdraf\\Desktop\\rediff minni project");

			driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(300);
			//Maximize the browser

			driver.Manage().Window.Maximize();
			//to launch the website
			// To read url from text file

			TextFileWriteRead Twr = new TextFileWriteRead();
			Twr.DirectoryOperation();
			string url = System.IO.File.ReadAllText(@"C:\Users\mdraf\Desktop\rediff minni project\data1.text");
			Console.WriteLine(url);
			string[] web_url = System.IO.File.ReadAllLines(@"C:\Users\mdraf\Desktop\rediff minni project\data1.text");
			driver.Url = web_url[0];
			//driver.Url = "https://money.rediff.com/index.html";

			string emails = "saifeefarzana549@gmail.com";
			string password = "Smrsf549";
            
			//To click on sign in
			driver.FindElement(By.LinkText("Sign In")).Click();
			//To Enter the registered Email
			driver.FindElement(By.XPath("/html/body/div[2]/div[5]/div[3]/form/div/input")).SendKeys(emails);
			//To Enter the Password
			driver.FindElement(By.XPath("/html/body/div[2]/div[5]/div[3]/form/div/div[3]/input[1]")).SendKeys(password);
			// print the login to console 
			Console.WriteLine(emails);
			Console.WriteLine(password);

			//To click on checkbox Remember me
			IWebElement check_box = driver.FindElement(By.XPath("/html/body/div[2]/div[5]/div[3]/form/div/div[3]/div[2]/input"));
			//To Click on SignIn Button
			driver.FindElement(By.Name("loginsubmit")).Click();
			//Click to add portfolio

			driver.FindElement(By.XPath("/html/body/div[4]/div[2]/div/b/div[2]/div[1]/div[1]/div/input")).Click();
			//To ReadExcel File
			string path= @"C:\Users\mdraf\Desktop\rediff minni project\data1.xlsx";
			
			// Instantiate a Workbook object that represents Excel file.
			Workbook wb = new Workbook(path);

			// Access "Sheet1" from the workbook.
			Worksheet sheet = wb.Worksheets[0];

//For addint portfolio 
			#region First Portfolio Value

			// Access the "A1" cell in the sheet.
			// To pass the Stock name
			Cell cell11 = sheet.Cells.GetCell(1, 0);
			String value11 = cell11.Value.ToString();
			driver.FindElement(By.Id("addstockname")).SendKeys(value11);

			ReadOnlyCollection<IWebElement> companies = driver.FindElements(By.XPath("/html/body/div[5]"));

			foreach (IWebElement  company in companies)
			{
				Console.WriteLine(value11);
			}
			driver.FindElement(By.XPath("/html/body/div[5]/div[1]")).Click();



			//To pass the date value
			var dateTime = DateTime.Now;
			String date = dateTime.ToString("dd-MM-yyyy");
			driver.FindElement(By.Id("stockAddDate")).SendKeys(date);

			//To pass the Quantity
			Cell cell13 = sheet.Cells.GetCell(1, 1);
			var value13 = cell13.Value.ToString();

			driver.FindElement(By.XPath("/html/body/b/div[6]/form/div[2]/div/div[1]/div[3]/input")).SendKeys(value13);

			//To pass the Total Amount
			Cell cell14 = sheet.Cells.GetCell(1, 2);
			var value14 = cell14.Value.ToString();

			driver.FindElement(By.XPath("/html/body/b/div[6]/form/div[2]/div/div[1]/div[4]/input")).SendKeys(value14);

			// To click the Exchange Type



			//To click on Add Stock Button
			driver.FindElement(By.XPath("/html/body/b/div[6]/form/div[2]/div/div[1]/div[6]/div/input")).Click();
			Console.WriteLine("PORTFOLIO HAS BEEN CREATED");
			Console.WriteLine(value11);
			Console.WriteLine(date);
			Console.WriteLine(value13);
			Console.WriteLine(value14);
			#endregion
			#region Second Portfolio value
			driver.FindElement(By.XPath("/html/body/div[4]/div[2]/div/div[3]/div[1]/div[1]/div/input")).Click();
			// To pass the Stock name
			Cell cell21 = sheet.Cells.GetCell(2, 0);
			String value21 = cell21.Value.ToString();
			driver.FindElement(By.Id("addstockname")).SendKeys(value21);

			driver.FindElement(By.XPath("/html/body/div[19]/div")).Click();



			//To pass the date value
			var dateTime1 = DateTime.Now;
			String date1 = dateTime.ToString("dd-MM-yyyy");
			driver.FindElement(By.Id("stockAddDate")).SendKeys(date1);

			//To pass the Quantity
			Cell cell23 = sheet.Cells.GetCell(2, 1);
			var value23 = cell13.Value.ToString();

			driver.FindElement(By.XPath("/html/body/div[10]/form/div[2]/div/div[1]/div[3]/input")).SendKeys(value23);

			//To pass the Total Amount
			Cell cell24 = sheet.Cells.GetCell(2, 2);
			var value24 = cell24.Value.ToString();

			driver.FindElement(By.XPath("/html/body/div[10]/form/div[2]/div/div[1]/div[4]/input")).SendKeys(value24);

			// To click the Exchange Type



			//To click on Add Stock Button
			driver.FindElement(By.XPath("/html/body/div[10]/form/div[2]/div/div[1]/div[6]/div/input")).Click();
			Console.WriteLine("PORTFOLIO HAS BEEN CREATED");
			Console.WriteLine(value21);
			Console.WriteLine(date1);
			Console.WriteLine(value23);
			Console.WriteLine(value24);
			#endregion
			#region Third Portfolio value

			driver.FindElement(By.XPath("/html/body/div[4]/div[2]/div/b/div[2]/div[1]/div[1]/div/input")).Click();
			// To pass the Stock name
			Cell cell31 = sheet.Cells.GetCell(3, 0);
			String value31 = cell31.Value.ToString();
			driver.FindElement(By.Id("addstockname")).SendKeys(value31);

			driver.FindElement(By.XPath("/html/body/div[6]/div")).Click();



			//To pass the date value
			var dateTime2 = DateTime.Now;
			String date2 = dateTime.ToString("dd-MM-yyyy");
			driver.FindElement(By.Id("stockAddDate")).SendKeys(date2);
			//To pass the Quantity
			Cell cell33 = sheet.Cells.GetCell(3, 1);
			var value33 = cell33.Value.ToString();

			driver.FindElement(By.XPath("/html/body/b/div[6]/form/div[2]/div/div[1]/div[3]/input")).SendKeys(value33);

			//To pass the Total Amount
			Cell cell34 = sheet.Cells.GetCell(3, 2);
			var value34 = cell34.Value.ToString();

			driver.FindElement(By.XPath("/html/body/b/div[6]/form/div[2]/div/div[1]/div[4]/input")).SendKeys(value34);

			// To click the Exchange Type



			//To click on Add Stock Button
			driver.FindElement(By.XPath("/html/body/b/div[6]/form/div[2]/div/div[1]/div[6]/div/input")).Click();
			Console.WriteLine("PORTFOLIO HAS BEEN CREATED");
			Console.WriteLine(value31);
			Console.WriteLine(date2);
			Console.WriteLine(value33);
			Console.WriteLine(value34);
			#endregion
			// for Deleting portfolio
			Thread.Sleep(3000);
            Console.WriteLine("PORTFOLIO DELETED");
			driver.FindElement(By.XPath("/html/body/div[4]/div[2]/div/div[3]/div[3]/table/tbody/tr[1]/td[1]/input[1]")).Click();
			driver.FindElement(By.XPath("/html/body/div[4]/div[2]/div/div[3]/div[3]/table/tbody/tr[1]/td[3]/div/input[2]")).Click();
			driver.SwitchTo().Alert().Accept();

			
			
		



			//Close the Browser
			//driver.Close();
			//driver.Quit();
		}


    }
}