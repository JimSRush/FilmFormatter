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

		List<Tuple<string, int>> titlesToRunTime = new List<Tuple<string, int>>();
		Dictionary<String, List<String>> citiesToVenues = new Dictionary<String, List<String>>();

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


		private void threadFilmsByTitle(Tuple<string, List<TitleSessionInfo>> t) 
		{
			var city = t.Item1;
			var films = t.Item2;
			List<Dictionary<String, List<TitleSessionInfo>>> filmsByTitle = parseFilmsByTitleForCity(city, films);
			writeOutTitlesToFile(filmsByTitle, city);

		}

		private void threadFilmsByDate(Tuple<string, List<TitleSessionInfo>> t)
		{
			var city = t.Item1;
			var films = t.Item2;
			List<Dictionary<DateTime, List<TitleSessionInfo>>> filmsByDate = parseFilmsByDateByCity(city, films);
			writeOutDatesToFile(filmsByDate, city);
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
					titlesToRunTime = GetTitlesFromRuntime(GetWorkSheetFromSheetName(workbookPart, "MAIN"), workbookPart);
					//parse main sheet
					WorksheetPart worksheetPart = GetWorkSheetFromSheetName(workbookPart, "SCREENING INFO");
					SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().Last();
					List<TitleSessionInfo> rawFilms = new List<TitleSessionInfo>();	

					rawFilms = parseFilms(sheetData, workbookPart);
					Console.WriteLine("Size of films");
					Console.WriteLine(rawFilms.Count);
					//parse the films and write to file
					List<String> cities = new List<String> { "AUCKLAND", "CHRISTCHURCH", "DUNEDIN", "GORE", "HAMILTON", "HAVELOCK NORTH", "NAPIER", "MASTERTON", "NELSON", "NEW PLYMOUTH", "PALMERSTON NORTH", "TAURANGA", "TIMARU", "WELLINGTON", "HAWKE'S BAY" };
					
					List<TitleSessionInfo> rawFilmsForOrderByDate = rawFilms.OrderBy(x => x.getDateTimeAsDate()).ThenBy(y => y.getTimeSpan()).ToList();
					foreach (String city in cities)
					{
						var t = System.Tuple.Create(city, rawFilmsForOrderByDate);

						//Thread th = new Thread(threadFilmsByDate);

						//Thread th = new Thread(new ParameterizedThreadStart(threadFilmsByTitle));
						//th.Start(t);


						threadFilmsByTitle(t);
						threadFilmsByDate(t);
					}
					watch.Stop();

					Console.WriteLine("Execution Time: {0} ms", watch.ElapsedMilliseconds);
					Application.Exit();

				}
			}
		}


		private List<Dictionary<DateTime, List<TitleSessionInfo>>> parseFilmsByDateByCity(String city, List<TitleSessionInfo> rawFilms)
		{
			List<Dictionary<DateTime, List<TitleSessionInfo>>> filmsByCityByDate = new List<Dictionary<DateTime, List<TitleSessionInfo>>>();
			foreach (TitleSessionInfo session in rawFilms)
			{
				if (session.getCity() == city)
				{
					if (!filmsByCityByDate.Any(dic => dic.ContainsKey(session.getDateTimeAsDate())))
					{
						Dictionary<DateTime, List<TitleSessionInfo>> toAdd = new Dictionary<DateTime, List<TitleSessionInfo>>() 
						{
							{session.getDateTimeAsDate(), new List<TitleSessionInfo>(){session}}
						};
						filmsByCityByDate.Add(toAdd);
					}
					else
					{
						foreach (Dictionary<DateTime, List<TitleSessionInfo>> dict in filmsByCityByDate)
						{
							if (dict.ContainsKey(session.getDateTimeAsDate()))
							{
								List<TitleSessionInfo> toUpdate = dict[session.getDateTimeAsDate()];
								toUpdate.Add(session);
								dict[session.getDateTimeAsDate()] = toUpdate;
							}
						}
					}
				}
			}

			return filmsByCityByDate;
		}



		private void writeOutDatesToFile(List<Dictionary<DateTime, List<TitleSessionInfo>>> filmsByDate, String city)
		{
			String outPutFolder = @"C:\Temp\";
			System.IO.Directory.CreateDirectory(outPutFolder);
			using (System.IO.StreamWriter file = new System.IO.StreamWriter(outPutFolder + city + "filmsByDate.txt"))
			{
				foreach (Dictionary<DateTime, List<TitleSessionInfo>> date in filmsByDate)
				{
					String newDate = setDateAsString(date.Keys.First());

					file.WriteLine(newDate);
					foreach (List<TitleSessionInfo> value in date.Values)
					{
						foreach (TitleSessionInfo cs in value)
						{
							//find runtime
							String shortRunTime = "";
								//For the bigger cities
							if (citiesToVenues[city].Count > 1) {
								if (cs.getShort().Equals("NO SHORT", StringComparison.InvariantCultureIgnoreCase))
								{
									String toWrite = cs.getSessionType() + "\t" + cs.getTime() + "\t" + cs.getTitle() + " (" + cs.getVenue() + ") " + getRunTimeFromTitle(cs.getTitle()) + "\t" + cs.getPageNumber();
									file.WriteLine(toWrite);
								}
								else if (!cs.getShort().Equals("INTERMISSION", StringComparison.InvariantCultureIgnoreCase)
								  && !cs.getShort().Equals("OUTWARDS", StringComparison.InvariantCultureIgnoreCase)
								  && !cs.getShort().Equals("INWARDS", StringComparison.InvariantCultureIgnoreCase)
								  && !cs.getShort().Equals("FILMMAKER PRESENT", StringComparison.InvariantCultureIgnoreCase)
								  && !cs.getShort().Equals("INTERMISSION", StringComparison.InvariantCultureIgnoreCase))
								{
									String toWrite = cs.getSessionType() + "\t" + cs.getTime() + "\t" + cs.getTitle() + " (" + cs.getVenue() + ") " + getRunTimeFromTitle(cs.getTitle()) + " + " + getRunTimeFromTitle(cs.getShort()) + "\t" + cs.getPageNumber();
									file.WriteLine(toWrite);
									
								}
								else
								{
									String toWrite = cs.getSessionType() + "\t" + cs.getTime() + "\t" + cs.getTitle() + " (" + cs.getVenue() + ") " + getRunTimeFromTitle(cs.getTitle()) + " + " + getRunTimeFromTitle(cs.getShort()) + "\t" + cs.getPageNumber();
									file.WriteLine(toWrite);
									
								}
								
							}
							else { //this is for the single venuie cities
								if (cs.getShort().Equals("NO SHORT", StringComparison.InvariantCultureIgnoreCase))
								{
									String toWrite = cs.getSessionType() + "\t" + cs.getTime() + "\t" + cs.getTitle() + " (" + getRunTimeFromTitle(cs.getTitle()) + ") " + "\t" + cs.getPageNumber();
									
								}
								else if (!cs.getShort().Equals("INTERMISSION", StringComparison.InvariantCultureIgnoreCase)
								  && !cs.getShort().Equals("OUTWARDS", StringComparison.InvariantCultureIgnoreCase)
								  && !cs.getShort().Equals("INWARDS", StringComparison.InvariantCultureIgnoreCase)
								  && !cs.getShort().Equals("FILMMAKER PRESENT", StringComparison.InvariantCultureIgnoreCase)
								  && !cs.getShort().Equals("INTERMISSION", StringComparison.InvariantCultureIgnoreCase))
								{
									String toWrite = cs.getSessionType() + "\t" + cs.getTime() + "\t" + cs.getTitle() + " (" + getRunTimeFromTitle(cs.getTitle()) + " + " + getRunTimeFromTitle(cs.getShort()) + ") " + "\t" + cs.getPageNumber();
									file.WriteLine(toWrite);
								}
								else
								{
									String toWrite = cs.getSessionType() + "\t" + cs.getTime() + "\t" + cs.getTitle() + " (" + getRunTimeFromTitle(cs.getTitle()) + ") " + "\t" + cs.getPageNumber();
									file.WriteLine(toWrite);
								}
							}
						}
					}
				}
			}
		}

		private String setDateAsString(DateTime filmDate)
		{
			return String.Format("{0:dddd d MMMM}", filmDate);
		}


		private void writeOutTitlesToFile(List<Dictionary<String, List<TitleSessionInfo>>> filmsByTitle, String city)
		{
			String outPutFolder = @"C:\Temp\";
			System.IO.Directory.CreateDirectory(outPutFolder);
			using (System.IO.StreamWriter file = new System.IO.StreamWriter(outPutFolder + city + "filmsByTitle.txt"))
			{
				foreach (Dictionary<String, List<TitleSessionInfo>> film in filmsByTitle)
				{
					String title = film.Keys.First();
					file.WriteLine(title);

					//Iterate the values -- have to get the values by the ey
					foreach (List<TitleSessionInfo> value in film.Values)
					{
						foreach (TitleSessionInfo currentSession in value)
						{
							//finally
							String toWrite = currentSession.getSessionType() + "\t" + currentSession.getVenue() + "\t" + currentSession.getDate() + ", " + currentSession.getTime();
							file.WriteLine(toWrite);
						}
					}
				}
			}
		}

		private List<Dictionary<String, List<TitleSessionInfo>>> parseFilmsByTitleForCity(String city, List<TitleSessionInfo> rawFilms)
		{
			List<Dictionary<String, List<TitleSessionInfo>>> filmByCity = new List<Dictionary<String, List<TitleSessionInfo>>>();
			foreach (TitleSessionInfo session in rawFilms)
			{
				if (session.getCity() == city)
				{
					if (!filmByCity.Any(dic => dic.ContainsKey(session.getTitle())))
					{
						Dictionary<String, List<TitleSessionInfo>> toAdd = new Dictionary<string, List<TitleSessionInfo>>() 
						{
							{session.getTitle(), new List<TitleSessionInfo>(){session}}
						};
						filmByCity.Add(toAdd);
					}
					else
					{
						foreach (Dictionary<String, List<TitleSessionInfo>> dict in filmByCity)
						{
							if (dict.ContainsKey(session.getTitle()))
							{
								List<TitleSessionInfo> toUpdate = dict[session.getTitle()];
								toUpdate.Add(session);
								dict[session.getTitle()] = toUpdate;

							}
						}

					}

				}
			}
			return filmByCity;
		}



		private List<TitleSessionInfo> parseFilms(SheetData sheetData, WorkbookPart workbookpart)
		{
			List<TitleSessionInfo> rawSchedule = new List<TitleSessionInfo>();

			int titlePosition = FilmFormatter.Tools.SpreadsheetHelpers.ColumnLetterToColumnIndex("D");
			int datePosition = FilmFormatter.Tools.SpreadsheetHelpers.ColumnLetterToColumnIndex("I");//6
			int timePosition = FilmFormatter.Tools.SpreadsheetHelpers.ColumnLetterToColumnIndex("J");//7
			int venuePosition = FilmFormatter.Tools.SpreadsheetHelpers.ColumnLetterToColumnIndex("L");//9
			int cityPosition = FilmFormatter.Tools.SpreadsheetHelpers.ColumnLetterToColumnIndex("N");//10
			int shortPosition = FilmFormatter.Tools.SpreadsheetHelpers.ColumnLetterToColumnIndex("G"); //this is empty in the case of INWARDS/OUTWARDS, so need this to check against.
			int pagePosition = FilmFormatter.Tools.SpreadsheetHelpers.ColumnLetterToColumnIndex("BG"); ///AB column
			var columnLetters = new List<string>() { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

			foreach (Row r in sheetData.Elements<Row>())
			{
				List<Cell> row = FilmFormatter.Tools.SpreadsheetHelpers.GetCellsForRow(r, columnLetters).ToList();

				Cell titleCell = row.ElementAtOrDefault(titlePosition);
				Cell dateCell = row.ElementAtOrDefault(datePosition);
				Cell timeCell = row.ElementAtOrDefault(timePosition);
				Cell venueCell = row.ElementAtOrDefault(venuePosition);
				Cell cityCell = row.ElementAtOrDefault(cityPosition);
				Cell shortCell = row.ElementAtOrDefault(shortPosition);
				Cell pageCell = row.ElementAtOrDefault(pagePosition);

				String title = "";
				String venue = "";
				String city = "";
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
							title = workbookpart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(Convert.ToInt32(titleCell.CellValue.Text)).InnerText;
							venue = workbookpart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(Convert.ToInt32(venueCell.CellValue.Text)).InnerText;
							city = workbookpart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(Convert.ToInt32(cityCell.CellValue.Text)).InnerText;
							shortFilm = workbookpart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(Convert.ToInt32(shortCell.CellValue.Text)).InnerText;

							//And the time
							String formattedValue = timeCell.InnerText;
							Decimal timeAsDecimal = Convert.ToDecimal(formattedValue) * 24;
							ts = TimeSpan.FromHours(Decimal.ToDouble(timeAsDecimal));
							if (pageCell != null)
							{
								int v;
								if (Int32.TryParse(pageCell.InnerText, out v))
								{
									pageNumber = v;
								}
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
								TitleSessionInfo sessionInfo = new TitleSessionInfo(title, venue, city, newDate, ts, shortFilm, pageNumber);
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
			setCitiesToVenues(rawSchedule);

			return rawSchedule;
		}

		private void setCitiesToVenues (List<TitleSessionInfo> rawSchedule) 
		{
		foreach (TitleSessionInfo session in rawSchedule)
			{
				if (!citiesToVenues.ContainsKey(session.getCity()))
				{
					List<String> values = new List<String>();
					values.Add(session.getVenue());
					citiesToVenues.Add(session.getCity(), values);
				}
				else
				{ //the city exists, so we need to check for the value in the values 
					if (!citiesToVenues[session.getCity()].Contains(session.getVenue()))
					{
						citiesToVenues[session.getCity()].Add(session.getVenue());
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

		private int getRunTimeFromTitle(String title)
		{
			foreach (Tuple<String, int> tuple in titlesToRunTime)
			{
				if (tuple.Item1 == title)
				{
					return tuple.Item2;
				}
			}
			return 0;
		}

		private List<Tuple<String, int>> GetTitlesFromRuntime(WorksheetPart worksheetpart, WorkbookPart workbookpart)
		{
			SheetData sheetData = worksheetpart.Worksheet.Elements<SheetData>().Last();
			List<Tuple<string, int>> ttrt = new List<Tuple<string, int>>();
			var columnLetters = new List<string>() { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

			foreach (Row r in sheetData.Elements<Row>())
			{
				List<Cell> row = FilmFormatter.Tools.SpreadsheetHelpers.GetCellsForRow(r, columnLetters).ToList();

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
