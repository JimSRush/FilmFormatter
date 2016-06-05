using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace FilmFormatter
{
	public partial class Form1 : Form
	{

		List<Tuple<string, int>> titlesToRunTime = new List<Tuple<string, int>>();

		public Form1()
		{
			InitializeComponent();
		}

		private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
		{

		}

		private void button1_Click(object sender, EventArgs e)
		{
			DialogResult result = openFileDialog1.ShowDialog();
			Console.WriteLine("I've just clicked a button");
			String file = openFileDialog1.FileName;
			Console.WriteLine("File is called " + file);
			loadFile(file);
		}

		private void loadFile(String fileName)
		{
			using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				using (SpreadsheetDocument myDoc = SpreadsheetDocument.Open(fs, false))
				{
					WorkbookPart workbookPart = myDoc.WorkbookPart;
					titlesToRunTime = GetTitlesFromRuntime(GetWorkSheetFromSheetName(workbookPart, "MAIN"), workbookPart);
					//parse main sheet
					WorksheetPart worksheetPart = GetWorkSheetFromSheetName(workbookPart, "SCREENING INFO");
					SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().Last();
					Console.WriteLine("Opened sheet");
					
					//parse by title -- List<Dictionary<title, titleSession>() 
					List<Dictionary<string, List<TitleSessionInfo>>> filmsByTitle = new List<Dictionary<string, List<TitleSessionInfo>>>();					//
					
					printSheetToConsole(sheetData, workbookPart);

					//parse

				}
			}
		}

		private List<Dictionary<string, List<TitleSessionInfo>>> GetFilmsByTitle(SheetData sheetData, WorkbookPart workbookpart) {
			List<Dictionary<string, List<TitleSessionInfo>>> filmsByTitle = new List<Dictionary<string, List<TitleSessionInfo>>>();
			return filmsByTitle;
		}

		private void printSheetToConsole(SheetData sheetData, WorkbookPart workbookPart)
		{
			foreach (Row r in sheetData.Elements<Row>())
			{
				foreach (Cell c in r.Elements<Cell>())
				{   //If it's not null, it's a string
					if (c != null)
					{
						if (c.DataType != null)
						{
							if (c.DataType == CellValues.SharedString)
							{
								String text = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(Convert.ToInt32(c.CellValue.Text)).InnerText;
								Console.WriteLine(text);
							}
						}
						else
						{
							int value;
							if (int.TryParse(c.InnerText, out value))
							{
								if (value != 0)
								{
									Console.WriteLine(c.InnerText);
									DateTime newDate = DateTime.FromOADate(value + 1462);
									Console.WriteLine(newDate);
								}
							}
						}
					}
				}
			}
		}

		private WorksheetPart GetWorkSheetFromSheetName(WorkbookPart workbookpart, String sheetName)
		{
			Sheet sheet = workbookpart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);
			if (sheet == null) throw new Exception(string.Format("Could not find sheet with name {0}", sheetName));
			else return workbookpart.GetPartById(sheet.Id) as WorksheetPart;
		}

		private List<Tuple<String, int>> GetTitlesFromRuntime(WorksheetPart worksheetpart, WorkbookPart workbookpart)
		{
			SheetData sheetData = worksheetpart.Worksheet.Elements<SheetData>().Last();
			List<Tuple<string, int>> ttrt = new List<Tuple<string, int>>();
			Console.WriteLine("The size of the list is: " + sheetData.Elements<Row>().Count());

			foreach (Row r in sheetData.Elements<Row>())
			{
				Console.WriteLine("This this ain't all that big " + r.Elements<Cell>().Count());
				Cell titleCell = r.Elements<Cell>().ElementAtOrDefault(2);
				Cell runningTimeCell = r.Elements<Cell>().ElementAtOrDefault(10);

				if (titleCell != null)
				{
					if (titleCell.DataType != null)
					{
						String title = "";
						if (titleCell.DataType == CellValues.SharedString)
						{
							title = workbookpart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(Convert.ToInt32(titleCell.CellValue.Text)).InnerText;
							Console.WriteLine(title);
							Console.WriteLine("Some random execution");
						}

						int runningTime;
						if (int.TryParse(runningTimeCell.InnerText, out runningTime))
						{
							Console.WriteLine(runningTime);
						}
						ttrt.Add(System.Tuple.Create(title, runningTime));
					}
				}
			}
			return ttrt;
		}
	}
}
