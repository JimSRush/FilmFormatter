using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;

namespace FilmFormatter.Tools
{
	class SpreadSheetWorkers
	{

		public static List<Tuple<string, int>> titlesToRunTime = new List<Tuple<string, int>>();
		public static  Dictionary<String, List<String>> citiesToVenues = new Dictionary<String, List<String>>();

		public static List<Dictionary<String, List<TitleSessionInfo>>> parseFilmsByTitleForCity(String city, List<TitleSessionInfo> rawFilms)
		{
			List<Dictionary<String, List<TitleSessionInfo>>> filmByCity = new List<Dictionary<String, List<TitleSessionInfo>>>();
			foreach (TitleSessionInfo session in rawFilms)
			{
				if (session.getCity() == city)
				{//if it doesn't exiat, add it
					if (!filmByCity.Any(dic => dic.ContainsKey(session.getTitle())))
					{
						Dictionary<String, List<TitleSessionInfo>> toAdd = new Dictionary<string, List<TitleSessionInfo>>() 
						{
							{session.getTitle(), new List<TitleSessionInfo>(){session}}
						};
						filmByCity.Add(toAdd);
					}
					else
					{//otherwise add it
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

		public static List<Dictionary<DateTime, List<TitleSessionInfo>>> parseFilmsByDateByCity(String city, List<TitleSessionInfo> rawFilms)
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

		private static void DateStart(Tuple<string, List<TitleSessionInfo>> films)
		{
			var city = films.Item1;
			var f = films.Item2;
			var filmsByDate = parseFilmsByDateByCity(city, f);
			writeOutDatesToFile(filmsByDate, city);
		}
		
		private static void TitleStart(Tuple<string, List<TitleSessionInfo>>  films)
		{
			var city = films.Item1;
			var f = films.Item2;
			List<Dictionary<String, List<TitleSessionInfo>>> filmsByTitle = parseFilmsByTitleForCity(city, f);
			writeOutTitlesToFile(filmsByTitle, city);
		}

		public static Thread threadFilmsByDate(Tuple<string, List<TitleSessionInfo>> films)
		{
			var f = films;
			var t = new Thread(() => DateStart(f));
			t.Start();
			return t;
		}

		public static Thread threadFilmsByTitle(Tuple<string, List<TitleSessionInfo>> films)
		{
			var f = films;
			var t = new Thread(() => TitleStart(f));
			t.Start();
			return t;
		}


		public static String setDateAsString(DateTime filmDate)
		{
			return String.Format("{0:dddd d MMMM}", filmDate);
		}

		public static void setCitiesToVenues(List<TitleSessionInfo> rawSchedule)
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

		public static int getRunTimeFromTitle(String title)
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


		public static void writeOutDatesToFile(List<Dictionary<DateTime, List<TitleSessionInfo>>> filmsByDate, String city)
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
		
		public static void writeOutTitlesToFile(List<Dictionary<String, List<TitleSessionInfo>>> filmsByTitle, String city)
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


	}
}
