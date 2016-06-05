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

		public String getSessionType()
		{
			return this.sessionType;
		}

		public String getVenue()
		{
			return this.venue;
		}

		public String getDate()
		{
			return this.date;
		}

		public String getCity()
		{
			return this.city;
		}

		public String getTime()
		{
			return this.time;
		}

		public string getTitle() 
		{ 
			return this.filmTitle;
		}
		private static readonly Dictionary<string, string> venuesToAbbreviations = new Dictionary<string, string>
		{
			{"ACADEMY", "AC"},
			{"BERGMAN", "PB"},
			{"CINEMA GOLD", "Cinema Gold"},
			{"CITY GALLERY", "CG"},
			{"CIVIC", "CIVIC"},
			{"DELUXE", "ED"},
			{"DOWNTOWN CINEMA", "DOWNTOWN"},
			{"EMBASSY", "EMB"},
			{"EVENT", "QSt"},
			{"FILM ARCHIVE", "NT"},
			{"HOYTS NORTHLAND 3", "HOYTS"},
			{"HOYTS NORTHLAND 4", "HOYTS"},
			{"HOYTS NORTHLAND 2", "HOYTS"},
			{"LIDO", "LIDO"},
			{"LIGHTHOUSE", "LHP"},
			{"PARAMOUNT", "PAR"},
			{"PENTHOUSE", "PH"},
			{"REGENT", "REGENT"},
			{"RIALTO", "RIALTO"},
			{"RIALTO D", "RIALTO"},
			{"ROXY", "RX"},
			{"SKYCITY", "SCT"},
			{"TEPAPA", "TP"}
		};

		public TitleSessionInfo(String title, String venue, String city, DateTime filmDate, TimeSpan screeningTime, String shortFilm)
		{
			setScreeningTimeAsALetter(screeningTime, filmDate);
			formatSessionTime(screeningTime);
			setDateAsString(filmDate);
			setAbbreviatedVenue(venue);
			this.filmTitle = title;
			this.city = city;
		}
		private void setDateAsString(DateTime filmDate)
		{
			this.date = String.Format("{0:dddd d MMMM}", filmDate);
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
			String formattedTime = dateTime.ToString("h:mm tt", System.Globalization.CultureInfo.InvariantCulture);
			this.time = formattedTime;
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
