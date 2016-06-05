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
					//titlesToRunTime = GetTitlesFromRuntime(GetWorkSheetFromSheetName(workbookPart, "MAIN"), workbookPart);
					//parse main sheet
					WorksheetPart worksheetPart = GetWorkSheetFromSheetName(workbookPart, "SCREENING INFO");
					SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().Last();
					//parse by title -- List<Dictionary<title, titleSession>() 
					List<Dictionary<string, List<TitleSessionInfo>>> filmsByTitle = new List<Dictionary<string, List<TitleSessionInfo>>>();					//

					filmsByTitle = GetFilmsByTitle(sheetData, workbookPart);


				}
			}
		}

		private List<Dictionary<string, List<TitleSessionInfo>>> GetFilmsByTitle(SheetData sheetData, WorkbookPart workbookpart)
		{

			List<Dictionary<string, List<TitleSessionInfo>>> filmsByTitle = new List<Dictionary<string, List<TitleSessionInfo>>>();

			int titlePosition = 3;
			int datePosition = 6;
			int timePosition = 7;
			int venuePosition = 9;
			int cityPosition = 10;

			foreach (Row r in sheetData.Elements<Row>())
			{
				Cell titleCell = r.Elements<Cell>().ElementAtOrDefault(titlePosition);
				Cell dateCell = r.Elements<Cell>().ElementAtOrDefault(datePosition);
				Cell timeCell = r.Elements<Cell>().ElementAtOrDefault(timePosition);
				Cell venueCell = r.Elements<Cell>().ElementAtOrDefault(venuePosition);
				Cell cityCell = r.Elements<Cell>().ElementAtOrDefault(cityPosition);

				String title = "";
				String venue = "";
				String city = "";
				DateTime newDate = new DateTime();
				TimeSpan ts = new TimeSpan();

				if (titleCell != null && venueCell != null && cityCell != null)
				{
					//Time to pluck out the title, venue and city.
					if (titleCell.DataType != null && venueCell.DataType != null && cityCell.DataType != null)
					{
						if (titleCell.DataType == CellValues.SharedString && venueCell.DataType == CellValues.SharedString && cityCell.DataType == CellValues.SharedString)
						{

							title = workbookpart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(Convert.ToInt32(titleCell.CellValue.Text)).InnerText;
							venue = workbookpart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(Convert.ToInt32(venueCell.CellValue.Text)).InnerText;
							city = workbookpart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(Convert.ToInt32(cityCell.CellValue.Text)).InnerText;
							String formattedValue = timeCell.InnerText;

							Decimal timeAsDecimal = Convert.ToDecimal(formattedValue) * 24;
							ts = TimeSpan.FromHours(Decimal.ToDouble(timeAsDecimal));
							int value;
							if (int.TryParse(dateCell.InnerText, out value))
							{
								if (value != 0)
								{
									newDate = DateTime.FromOADate(value + 1462);
								}
							}
						}
					}

					if (title!=null || title != "")
					{
						Console.WriteLine("Title: " + title);
						Console.WriteLine("Date: " + newDate);
						Console.WriteLine("Time: " + ts);
						Console.WriteLine("Venue: " + venue);
						Console.WriteLine("City: " + city + "\n");
					}
					
				}
				
			}
			//for each row
			//parse venue, date, time, title
			//check list to see if key exists
			//if doesn't exist, create new + insert in list
			//if exists, add to list associated with title key



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

			foreach (Row r in sheetData.Elements<Row>())
			{
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
