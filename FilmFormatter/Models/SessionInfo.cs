using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FilmFormatter
{
	class TitleSessionInfo
	{
		private string sessionType;
		private string venue;
		private string date;
		private string city;
		private string time;
		private string filmTitle;
		private string shortFilm;
		private TimeSpan screeningTimeAsTimeSpan;
		private int pageNumber;
		private DateTime dateTimeDate;


		public string getShort()
		{ 
			return this.shortFilm;
		}
		public String getSessionType()
		{
			return this.sessionType;
		}

		public DateTime getDateTimeAsDate()
		{ 
			return this.dateTimeDate;
		}
		public String getVenue()
		{
			return this.venue;
		}

		public String getDate()
		{
			return this.date;
		}

		public TimeSpan getTimeSpan()
		{ 
			return this.screeningTimeAsTimeSpan;
		} 
		public String getCity()
		{
			return this.city;
		}

		public String getTime()
		{
			return this.time;
		}

		public int getPageNumber() 
		{ 
			return this.pageNumber;
		}
		public string getTitle() 
		{ 
			return this.filmTitle;
		}
		private static readonly Dictionary<string, string> venuesToAbbreviations = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
		{
			{"ACADEMY", "AC"},
			{"BERGMAN", "PB"},
			{"CINEMA GOLD", "Havelock Nth"},
			{"CITY GALLERY", "CG"},
			{"CIVIC", "CIVIC"},
			{"DELUXE", "ED"},
			{"DOWNTOWN CINEMA", "P. North"},
			{"EMBASSY", "EMB"},
			{"EVENT (New Plymouth)", "N. Ply, Event"},
			{"QUEEN ST", "QSt"},
			{"MANAKAU", "MK"},
			{"WESTGATE", "WGATE"},
			{"HOLLYWOOD", "HWOOD"},
			{"ASB WATERFRONT THEATRE", "AWT"},
			{"HOYTS NORTHLAND 3", "Northlands"},
			{"HOYTS NORTHLAND 4", "Northlands"},
			{"HOYTS NORTHLAND 2", "Northlands"},
			{"ISAAC THEATRE ROYAL (Christchurch)", "Isaac Theatre"},
			{"LEN LYE (New Plymouth)", "N. Ply, Len Lye"},
			{"LIDO", "Hamilton"},
			{"LIGHTHOUSE", "LHP"},
			{"MOVIEMAX (Timaru)", "MM Timaru"},
			{"NGA TAONGA", "NT"},
			{"PARAMOUNT", "PAR"},
			{"PENTHOUSE", "PH"},
			{"REGENT (Dunedin)", "Regent"},
			{"REGENT (Masterton)", "Masterton"},
			{"RIALTO (Auckland)", "Rialto"},
			{"RIALTO (Dunedin)", "Rialto"},
			{"RIALTO (Tauranga)", "Tauranga"},
			{"ROXY", "RX"},
			{"SKY CITY", "SCT"},
			{"ST JAMES (Gore)", "SJ Gore"},
			{"STATE CINEMA (Nelson)", "STATE"},
			{"SUTER (Nelson)", "SUTER"},
			{"TEPAPA", "TP"}
		};

		public TitleSessionInfo(String title, String venue, string city, DateTime filmDate, TimeSpan screeningTime, String shortFilm, int pageNumber)
		{
			setScreeningTimeAsALetter(screeningTime, filmDate);
			formatSessionTime(screeningTime);
			setDateAsString(filmDate);
			setAbbreviatedVenue(venue);
			this.filmTitle = title;
			this.city = city;
			this.shortFilm = shortFilm;
			this.screeningTimeAsTimeSpan = screeningTime;
			this.pageNumber = pageNumber;
			this.dateTimeDate = filmDate;
		}
		private void setDateAsString(DateTime filmDate)
		{
			this.date = String.Format("{0:ddd d MMM}", filmDate);
		}

		private void setAbbreviatedVenue(String venue)
		{
			if (!venuesToAbbreviations.ContainsKey(venue)) 
			{
				this.venue = venue;
				return;
			}
			this.venue = venuesToAbbreviations[venue];
		}

		private void formatSessionTime(TimeSpan screeningTime)
		{
			DateTime dateTime = new DateTime(screeningTime.Ticks); //back to datetime we go
			String formattedTime = dateTime.ToString("h.mm tt", System.Globalization.CultureInfo.InvariantCulture);
			
			this.time = formattedTime.ToLower();
		}

		private void setScreeningTimeAsALetter(TimeSpan ts, DateTime date)
		{
			//5pm!
			TimeSpan toCompareAgainst = TimeSpan.FromHours(17);

			if (date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday)
			{
				this.sessionType = "A";
				return;
			}
			if (ts.CompareTo(toCompareAgainst) == -1)
			{
				this.sessionType = "B";
				return;
			}
			if (ts.CompareTo(toCompareAgainst) == 0 || ts.CompareTo(toCompareAgainst) == 1)
			{
				this.sessionType = "A";
				return;
			}
			return;
		}
	}
}
