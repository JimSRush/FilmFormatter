﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
					//parse by title -- List<Dictionary<title, titleSession>() 
					List<TitleSessionInfo> rawFilms = new List<TitleSessionInfo>();					//

					rawFilms = parseFilms(sheetData, workbookPart);

					//OK, now we have the films in memory. Now what?

					//GIve me everything in Auckland
					List<Dictionary<String, List<TitleSessionInfo>>> aucklandFilmsByTitle = parseFilmsByTitleForCity("AUCKLAND", rawFilms);
					List<Dictionary<String, List<TitleSessionInfo>>> aucklandFilmsByDate = parseFilmsByDateByCity("AUCKLAND", rawFilms);

					List<Dictionary<String, List<TitleSessionInfo>>> wellingtonFilmsByTitle = parseFilmsByTitleForCity("WELLINGTON", rawFilms);
					List<Dictionary<String, List<TitleSessionInfo>>> wellingtonFilmsByDate = parseFilmsByDateByCity("WELLINGTON", rawFilms);
					//parse the films and write to file
					List<String> cities = new List<String> { "AUCKLAND", "CHRISTCHURCH", "DUNEDIN", "GORE", "HAMILTON", "NAPIER", "MASTERTON", "NELSON", "NEW PLYMOUTH", "PALMERSTON NORTH", "TAURANGA", "TIMARU", "WELLINGTON" };

					foreach (String city in cities)
					{
						List<Dictionary<String, List<TitleSessionInfo>>> filmsByTitle = parseFilmsByTitleForCity(city, rawFilms);
						List<Dictionary<String, List<TitleSessionInfo>>> filmsByDate = parseFilmsByDateByCity(city, rawFilms);
						writeOutTitlesToFile(filmsByTitle, city);
						writeOutDatesToFile(filmsByDate, city);
					}
					Application.Exit();

				}
			}
		}

		private static string GetColumnAddress(string cellReference)
		{
			//Create a regular expression to get column address letters.
			Regex regex = new Regex("[A-Za-z]+");
			Match match = regex.Match(cellReference);
			return match.Value;
		}  

		private static IEnumerable<Cell> GetCellsForRow(Row row, List<string> columnLetters)
		{
			int workIdx = 0;
			foreach (var cell in row.Descendants<Cell>())
			{
				//Get letter part of cell address
				var cellLetter = GetColumnAddress(cell.CellReference);

				//Get column index of the matched cell  
				int currentActualIdx = columnLetters.IndexOf(cellLetter);

				//Add empty cell if work index smaller than actual index
				for (; workIdx < currentActualIdx; workIdx++)
				{
					var emptyCell = new Cell() { DataType = null, CellValue = new CellValue(string.Empty) };
					yield return emptyCell;
				}

				//Return cell with data from Excel row
				yield return cell;
				workIdx++;

				//Check if it's ending cell but there still is any unmatched columnLetters item   
				if (cell == row.LastChild)
				{
					//Append empty cells to enumerable 
					for (; workIdx < columnLetters.Count(); workIdx++)
					{
						var emptyCell = new Cell() { DataType = null, CellValue = new CellValue(string.Empty) };
						yield return emptyCell;
					}
				}
			}
		}

		private List<Dictionary<String, List<TitleSessionInfo>>> parseFilmsByDateByCity(String city, List<TitleSessionInfo> rawFilms)
		{
			//sort rawfilms
			//List<TitleSessionInfo> newRawFilms = rawFilms.OrderBy(x => x.getTimeSpan()).ToList();

			List<Dictionary<String, List<TitleSessionInfo>>> filmsByCityByDate = new List<Dictionary<String, List<TitleSessionInfo>>>();
			foreach (TitleSessionInfo session in rawFilms)
			{
				if (session.getCity() == city)
				{
					if (!filmsByCityByDate.Any(dic => dic.ContainsKey(session.getDate())))
					{
						Dictionary<String, List<TitleSessionInfo>> toAdd = new Dictionary<string, List<TitleSessionInfo>>() 
						{
							{session.getDate(), new List<TitleSessionInfo>(){session}}
						};
						filmsByCityByDate.Add(toAdd);
					}
					else
					{
						foreach (Dictionary<String, List<TitleSessionInfo>> dict in filmsByCityByDate)
						{
							if (dict.ContainsKey(session.getDate()))
							{
								List<TitleSessionInfo> toUpdate = dict[session.getDate()];
								toUpdate.Add(session);
								dict[session.getDate()] = toUpdate;
							}
						}
					}
				}
			}
			//sort here
			////List<TitleSessionInfo> newRawFilms = rawFilms.OrderBy(x => x.getTimeSpan()).ToList();
			//sortedFilmsByDate List<Dictionary<String, List<TitleSessionInfo>>> = 
			return filmsByCityByDate;
		}



		private void writeOutDatesToFile(List<Dictionary<String, List<TitleSessionInfo>>> filmsByDate, String city)
		{
			String outPutFolder = @"C:\Temp\";
			System.IO.Directory.CreateDirectory(outPutFolder);
			using (System.IO.StreamWriter file = new System.IO.StreamWriter(outPutFolder + city + "filmsByDate.txt"))
			{
				foreach (Dictionary<String, List<TitleSessionInfo>> date in filmsByDate)
				{
					String newDate = date.Keys.First();
					file.WriteLine(newDate);
					foreach (List<TitleSessionInfo> value in date.Values)
					{
						foreach (TitleSessionInfo cs in value)
						{
							//find runtime
							String shortRunTime = "";

							String toWrite = cs.getSessionType() + "\t" + cs.getTime() + "\t" + cs.getTitle() + "\t(" + cs.getVenue() + ") " + getRunTimeFromTitle(cs.getTitle()) + "		p" + cs.getPageNumber();
							file.WriteLine(toWrite);
						}
					}
				}
			}
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
					//if(list.Any(dic => dic.ContainsKey(item.Name)))
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
			//TODO: It should really be just a list of session objects
			List<TitleSessionInfo> rawSchedule = new List<TitleSessionInfo>();

			int titlePosition = 3; //3
			int datePosition = 8;//6
			int timePosition = 9;//7
			int venuePosition = 11;//9
			int cityPosition = 12;//10
			int shortPosition = 6; //this is empty in the case of INWARDS/OUTWARDS, so need this to check against.
			int pagePosition = 25; ///AU column
			var columnLetters = new List<string>() { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
			
			foreach (Row r in sheetData.Elements<Row>())
			{
			List<Cell> row = GetCellsForRow(r, columnLetters).ToList();


				Cell titleCell = row.ElementAtOrDefault(titlePosition);
				Cell dateCell = row.ElementAtOrDefault(datePosition);
				Cell timeCell = row.ElementAtOrDefault(timePosition);
				Cell venueCell = row.ElementAtOrDefault(venuePosition);
				Cell cityCell = row.ElementAtOrDefault(cityPosition);
				Cell shortCell = row.ElementAtOrDefault(shortPosition);
				Cell pageCell = row.ElementAtOrDefault(pagePosition);
				//Console.WriteLine("The size of R is " + r.Count());
				String title = "";
				String venue = "";
				String city = "";
				DateTime newDate = new DateTime();
				TimeSpan ts = new TimeSpan();
				String shortFilm = "";
				int pageNumber = -1;
				

				if (titleCell != null && venueCell != null && cityCell != null)
				{
					//Time to pluck out the title, venue and city.
					if (titleCell.DataType != null && venueCell.DataType != null && cityCell.DataType != null)
					{
						if (titleCell.DataType == CellValues.SharedString && venueCell.DataType == CellValues.SharedString && cityCell.DataType == CellValues.SharedString)
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

								Console.WriteLine("Title: " + title + "Pagecell: " + pageCell);
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
			return rawSchedule;
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
