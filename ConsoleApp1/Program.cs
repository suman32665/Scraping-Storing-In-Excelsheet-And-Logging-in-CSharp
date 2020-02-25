using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

namespace ConsoleApp1
{
	class Program
	{
		string TeamName = "//*[@class='p0c-competition-tables__team--short-name']";
		string MatchesPlayed = "//*[@class='p0c-competition-tables__matches-played']";
		string MatchesWon = "//*[@class='p0c-competition-tables__matches-won']";
		string MatchesDrawn = "//*[@class='p0c-competition-tables__matches-drawn']";
		string MatchesLost = "//*[@class='p0c-competition-tables__matches-lost']";
		string GoalsFor = "//*[@class='p0c-competition-tables__goals-for']";
		string GoalsAgainst = "//*[@class='p0c-competition-tables__goals-against']";
		string GoalsDiff = "//*[@class='p0c-competition-tables__goals-diff']";
		string Points = "//*[@class='p0c-competition-tables__pts']";

		string TeamNameValue, MatchesPlayedValue, MatchesWonValue, MatchesDrawnValue,
				MatchesLostValue, GoalsForValue, GoalsAgainstValue, GoalsDiffValue, PointsValue;

		public IWebDriver Webdriver;
		static void Main(string[] args)
		{
			Program obj = new Program();
			obj.SetLogger();
			NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
			try
			{
				ExcelPackage package = new ExcelPackage();
				ExcelWorksheet worksheet = CreateSheet(package, "Football Sheet");

				obj.OpenBrowser();
				obj.Scrap(worksheet);
				obj.SetHeader(worksheet);
				obj.AddFormula(worksheet);
				obj.AutoFitColumn(worksheet);
				obj.WriteExcelFile(package);
			}
			catch (Exception e)
			{
				logger.Error(e);
			}
		}

		public void SetLogger()
		{
			var config = new NLog.Config.LoggingConfiguration();

			// Targets where to log to: File and Console
			var logfile = new NLog.Targets.FileTarget("logfile") { FileName = System.IO.Directory.GetCurrentDirectory() + "\\Logs\\log.txt" };
			var logconsole = new NLog.Targets.ConsoleTarget("logconsole");

			// Rules for mapping loggers to targets            
			config.AddRule(NLog.LogLevel.Info, NLog.LogLevel.Fatal, logconsole);
			config.AddRule(NLog.LogLevel.Debug, NLog.LogLevel.Fatal, logfile);

			// Apply config           
			NLog.LogManager.Configuration = config;
		}

		public void OpenBrowser()
		{
			ChromeOptions chromeBrowserOptions = new ChromeOptions();
			chromeBrowserOptions.AddArgument("--start-maximized");
			chromeBrowserOptions.AddArgument("--test-type");
			chromeBrowserOptions.AddArgument("--silent");
			chromeBrowserOptions.AddArgument("--disable-plugins");
			chromeBrowserOptions.AddArgument("--disable-infobars");
			chromeBrowserOptions.AddArgument("--incognito");
			Webdriver = new ChromeDriver(".", chromeBrowserOptions);
			Webdriver.Navigate().GoToUrl("https://www.goal.com/en/tables");
		}

		public void Scrap(ExcelWorksheet ws)
		{
			string Tr = "(//*[@class='p0c-competition-tables__table'])[1]/tbody/tr";

			for (int i = 1; i <= 5; i++)
			{
				TeamNameValue = Webdriver.FindElement(By.XPath(Tr + '[' + i + ']' + TeamName)).Text;
				MatchesPlayedValue = Webdriver.FindElement(By.XPath(Tr + '[' + i + ']' + MatchesPlayed)).Text;
				MatchesWonValue = Webdriver.FindElement(By.XPath(Tr + '[' + i + ']' + MatchesWon)).Text;
				MatchesDrawnValue = Webdriver.FindElement(By.XPath(Tr + '[' + i + ']' + MatchesDrawn)).Text;
				MatchesLostValue = Webdriver.FindElement(By.XPath(Tr + '[' + i + ']' + MatchesLost)).Text;
				GoalsForValue = Webdriver.FindElement(By.XPath(Tr + '[' + i + ']' + GoalsFor)).Text;
				GoalsAgainstValue = Webdriver.FindElement(By.XPath(Tr + '[' + i + ']' + GoalsAgainst)).Text;
				GoalsDiffValue = Webdriver.FindElement(By.XPath(Tr + '[' + i + ']' + GoalsDiff)).Text;
				PointsValue = Webdriver.FindElement(By.XPath(Tr + '[' + i + ']' + Points)).Text;

				WriteData(ws, i + 2);
				Console.WriteLine(TeamNameValue + ' ' + MatchesPlayedValue + ' ' + MatchesWonValue + ' ' + MatchesDrawnValue + ' ' +
					MatchesLostValue + ' ' + GoalsForValue + ' ' + GoalsAgainstValue + ' ' + GoalsDiffValue + ' ' + PointsValue);
			}


			Webdriver.Quit();
		}

		private static ExcelWorksheet CreateSheet(ExcelPackage p, string sheetName)
		{
			ExcelWorksheet ws = p.Workbook.Worksheets.Add(sheetName);
			ws.Name = sheetName; //Setting Sheet's name
			ws.Cells.Style.Font.Size = 11; //Default font size for whole sheet
			ws.Cells.Style.Font.Name = "Calibri"; //Default Font name for whole sheet

			return ws;
		}

		public void SetHeader(ExcelWorksheet ws)
		{
			int colLength = ws.Dimension.End.Column;
			ws.Cells[1, 1].Value = "Football Data";
			ws.Cells[1, 1].Style.Font.Size = 20;
			ws.Cells[1, 1, 1, colLength].Merge = true;
			ws.Cells[1, 1, 1, colLength].Style.Font.Bold = true;
			ws.Cells[1, 1, 1, colLength].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

			ws.Cells[2, 1].Value = "Team Name";
			ws.Cells[2, 2].Value = "Matches Played";
			ws.Cells[2, 3].Value = "Matches Won";
			ws.Cells[2, 4].Value = "Matches Drawn";
			ws.Cells[2, 5].Value = "Matches Lost";
			ws.Cells[2, 6].Value = "Goals For";
			ws.Cells[2, 7].Value = "Goals Against";
			ws.Cells[2, 8].Value = "Goals Diff";
			ws.Cells[2, 9].Value = "Points";
			ws.Cells[2, 1, 2, colLength].Style.Font.Bold = true;
		}

		public void WriteData(ExcelWorksheet ws, int i)
		{
			ws.Cells[i, 1].Value = TeamNameValue;
			ws.Cells[i, 2].Value = Int32.Parse(MatchesPlayedValue);
			ws.Cells[i, 3].Value = Int32.Parse(MatchesWonValue);
			ws.Cells[i, 4].Value = Int32.Parse(MatchesDrawnValue);
			ws.Cells[i, 5].Value = Int32.Parse(MatchesLostValue);
			ws.Cells[i, 6].Value = Int32.Parse(GoalsForValue);
			ws.Cells[i, 7].Value = Int32.Parse(GoalsAgainstValue);
			ws.Cells[i, 8].Value = Int32.Parse(GoalsDiffValue);
			ws.Cells[i, 9].Value = Int32.Parse(PointsValue);
		}

		public void AddFormula(ExcelWorksheet ws)
		{
			int LastRow = ws.Dimension.End.Row;
			int LastColumn = ws.Dimension.End.Column;
			ws.Cells[LastRow + 1, 1].Value = "Total";

			for (int colIndex = 2; colIndex <= LastColumn; colIndex++)
			{
				ws.Cells[LastRow + 1, colIndex].Formula = "Sum(" + ws.Cells[3, colIndex].Address + ":" + ws.Cells[LastRow, colIndex].Address + ")";
			}
			ws.Cells[ws.Dimension.End.Row, 1, ws.Dimension.End.Row, LastColumn].Style.Font.Bold = true;
		}

		public void AutoFitColumn(ExcelWorksheet workSheet)
		{
			workSheet.Column(1).AutoFit();
			workSheet.Column(2).AutoFit();
			workSheet.Column(3).AutoFit();
			workSheet.Column(4).AutoFit();
			workSheet.Column(5).AutoFit();
			workSheet.Column(6).AutoFit();
			workSheet.Column(7).AutoFit();
			workSheet.Column(8).AutoFit();
			workSheet.Column(9).AutoFit();
		}

		public void WriteExcelFile(ExcelPackage package)
		{
			string timeStamp = DateTime.Now.ToString("yyyyMMdd");
			string path = System.IO.Directory.GetCurrentDirectory() + "\\Output\\output" + timeStamp + ".xlsx";
			if (File.Exists(path))
			{
				File.Delete(path);
			}

			FileStream fileStream = File.Create(path);
			fileStream.Close();
			File.WriteAllBytes(path, package.GetAsByteArray());
			package.Dispose();
		}
	}
}
