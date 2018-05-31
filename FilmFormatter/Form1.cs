using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Entity;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Linq.Dynamic;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Threading;


using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace FilmFormatter
{
	public partial class Form1 : Form
	{

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
			String file = openFileDialog1.FileName;
			loadFile(file);
		}




		private void loadFile(String fileName)
		{
			var watch = new System.Diagnostics.Stopwatch();
			watch.Start();
			using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				using (SpreadsheetDocument myDoc = SpreadsheetDocument.Open(fs, false))
				{
					WorkbookPart workbookPart = myDoc.WorkbookPart;
					
					FilmFormatter.Tools.SpreadSheetWorkers.titlesToRunTime = GetTitlesFromRuntime(GetWorkSheetFromSheetName(workbookPart, "MAIN"), workbookPart);
			
					//parse main sheet
					WorksheetPart worksheetPart = GetWorkSheetFromSheetName(workbookPart, "SCREENING INFO");
					Console.WriteLine("Reference");
					Console.WriteLine(worksheetPart.Worksheet.SheetDimension.Reference);
					SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().Last();
					List<TitleSessionInfo> rawFilms = new List<TitleSessionInfo>();	

					rawFilms = parseFilms(sheetData, workbookPart);
					Console.WriteLine("Size of films");
					Console.WriteLine(rawFilms.Count);
					//parse the films and write to file
					//List<String> cities = new List<String> { "AUCKLAND", "CHRISTCHURCH", "DUNEDIN", "GORE", "HAMILTON", "HAVELOCK NORTH", "NAPIER", "MASTERTON", "NELSON", "NEW PLYMOUTH", "PALMERSTON NORTH", "TAURANGA", "TIMARU", "WELLINGTON", "HAWKE'S BAY" };
					List<String> c = cities(rawFilms);
					List<String> p = programs(rawFilms);

					List<TitleSessionInfo> rawFilmsForOrderByDate = rawFilms.OrderBy(x => x.getDateTimeAsDate()).ThenBy(y => y.getTimeSpan()).ToList();
					foreach (String city in c)
					{
						var t = System.Tuple.Create(city, rawFilmsForOrderByDate);	
						FilmFormatter.Tools.SpreadSheetWorkers.threadFilmsByDate(t);
						FilmFormatter.Tools.SpreadSheetWorkers.threadFilmsByTitle(t);
					}

	

					Application.Exit();

				}
			}
		}

		private List<string> programs(List<TitleSessionInfo> films)
		{
			List<string> programs = new List<string>();

			foreach (TitleSessionInfo f in films)
			{
				if (!programs.Contains(f.getProgram()))
				{
					programs.Add(f.getProgram());
				}
			}
			return programs;
		}

		private List<string> cities(List<TitleSessionInfo> films) 
		{ 
			List<string> cities = new List<string>();

			foreach (TitleSessionInfo f in films) {
				if (!cities.Contains(f.getCity()))
				{ 
					cities.Add(f.getCity());
				}
			}
			return cities;
		}

		private List<TitleSessionInfo> parseFilms(SheetData sheetData, WorkbookPart workbookpart)
		{
			List<TitleSessionInfo> rawSchedule = new List<TitleSessionInfo>();

			int titlePosition = FilmFormatter.Tools.SpreadsheetHelpers.ColumnLetterToColumnIndex("D");
			int datePosition = FilmFormatter.Tools.SpreadsheetHelpers.ColumnLetterToColumnIndex("I");
			int timePosition = FilmFormatter.Tools.SpreadsheetHelpers.ColumnLetterToColumnIndex("J");
			int venuePosition = FilmFormatter.Tools.SpreadsheetHelpers.ColumnLetterToColumnIndex("L");
			int cityPosition = FilmFormatter.Tools.SpreadsheetHelpers.ColumnLetterToColumnIndex("N");
			int shortPosition = FilmFormatter.Tools.SpreadsheetHelpers.ColumnLetterToColumnIndex("G"); //this is empty in the case of INWARDS/OUTWARDS, so need this to check against.
			int pagePosition = FilmFormatter.Tools.SpreadsheetHelpers.ColumnLetterToColumnIndex("BI");
			int programPosition = FilmFormatter.Tools.SpreadsheetHelpers.ColumnLetterToColumnIndex("BG");
		

			SharedStringItem[] sharedStringItemsArray = workbookpart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ToArray<SharedStringItem>();
			//Get each row
			foreach (Row r in sheetData.Elements<Row>())
			{
				//Pluck out the individual row
				List<Cell> row = FilmFormatter.Tools.SpreadsheetHelpers.getExcelRowCells(r);

				Cell titleCell = row.ElementAtOrDefault(titlePosition);
				Cell dateCell = row.ElementAtOrDefault(datePosition);
				Cell timeCell = row.ElementAtOrDefault(timePosition);
				Cell venueCell = row.ElementAtOrDefault(venuePosition);
				Cell cityCell = row.ElementAtOrDefault(cityPosition);
				Cell shortCell = row.ElementAtOrDefault(shortPosition);
				Cell pageCell = row.ElementAtOrDefault(pagePosition);
				Cell programCell = row.ElementAtOrDefault(programPosition);

				String title = "";
				String venue = "";
				String city = "";
				String program = "";
				DateTime newDate = new DateTime();
				TimeSpan ts = new TimeSpan();
				String shortFilm = "";
				int pageNumber = 1;
				
				if (titleCell != null && venueCell != null && cityCell != null)
				{
					//Time to pluck out the title, venue and city.
					if (titleCell.DataType != null && venueCell.DataType != null && cityCell.DataType != null)
					{
						if (titleCell.DataType == CellValues.SharedString && venueCell.DataType == CellValues.SharedString && cityCell.DataType == CellValues.SharedString && timeCell.InnerText != "")
						{
							//Here, we have to get the text for each. TODO put this in a method.
							title = sharedStringItemsArray[int.Parse(titleCell.CellValue.Text)].InnerText;
							venue = sharedStringItemsArray[int.Parse(venueCell.CellValue.Text)].InnerText;
							city = sharedStringItemsArray[int.Parse(cityCell.CellValue.Text)].InnerText;
							shortFilm = sharedStringItemsArray[int.Parse(shortCell.CellValue.Text)].InnerText;
							program = sharedStringItemsArray[int.Parse(programCell.CellValue.Text)].InnerText;

							//And the time
							String formattedValue = timeCell.InnerText;
							Decimal timeAsDecimal = Convert.ToDecimal(formattedValue) * 24;
							ts = TimeSpan.FromHours(Decimal.ToDouble(timeAsDecimal));

							if (pageCell != null)
							{
								int v;
								if (int.TryParse(pageCell.CellValue.Text, out pageNumber)) ;
							}
							//AAAAaaand the date

							int value;

							if (int.TryParse(dateCell.InnerText, out value))
							{
								if (value != 0)
								{
									newDate = DateTime.FromOADate(value + 1462);
								}
							}
							if (!shortFilm.Equals("INWARDS") && !shortFilm.Equals("OUTWARDS"))
							{
								TitleSessionInfo sessionInfo = new TitleSessionInfo(title, venue, city, newDate, ts, shortFilm, pageNumber, program);
								//Gotta ignore the blank cells
								if (sessionInfo.getCity() != "")
								{
									rawSchedule.Add(sessionInfo);
								}
							}

						}
					}
				}
			}
			FilmFormatter.Tools.SpreadSheetWorkers.setCitiesToVenues(rawSchedule);

			return rawSchedule;
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
			var columnLetters = new List<string>() { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

			foreach (Row r in sheetData.Elements<Row>())
			{
				List<Cell> row = FilmFormatter.Tools.SpreadsheetHelpers.GetCellsForRow(r).ToList();

				Cell titleCell = row.ElementAtOrDefault(FilmFormatter.Tools.SpreadsheetHelpers.ColumnLetterToColumnIndex("C"));
				Cell runningTimeCell = row.ElementAtOrDefault(FilmFormatter.Tools.SpreadsheetHelpers.ColumnLetterToColumnIndex("L"));

				if (titleCell != null)
				{
					if (titleCell.DataType != null)
					{
						String title = "";
						if (titleCell.DataType == CellValues.SharedString)
						{
							title = workbookpart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(Convert.ToInt32(titleCell.CellValue.Text)).InnerText;
						}

						int runningTime;
						if (int.TryParse(runningTimeCell.InnerText, out runningTime))
						{

						}
						ttrt.Add(System.Tuple.Create(title, runningTime));
					}
				}
			}

			return ttrt;
		}

		private void progressBar1_Click(object sender, EventArgs e)
		{

		}
	}
}
