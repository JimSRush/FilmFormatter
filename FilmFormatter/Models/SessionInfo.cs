﻿using System;
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
        private int sortOrder;
        private string programmeNumber;


		public string getProgram()
		{
			return this.programmeNumber;
		}

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

		public override string ToString()
		{
			return "Title: " + filmTitle + " " + getTime();
		}

        public int getSortOrder()
        {
            return this.sortOrder;
        }

		public TitleSessionInfo(String title, String venue, string city, DateTime filmDate, TimeSpan screeningTime, String shortFilm, int pageNumber)
		{
			setScreeningTimeAsALetter(screeningTime, filmDate);
			formatSessionTime(screeningTime);
			setDateAsString(filmDate);
			setAbbreviatedVenue(venue);
            setProgrammeAndSortOrder(city);
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
			if (!FilmFormatter.Tools.SpreadSheetWorkers.vmappings.ContainsKey(venue)) 
			{
				this.venue = venue;
				return;
			}
			this.venue = FilmFormatter.Tools.SpreadSheetWorkers.vmappings[venue];
		}

        //find the city in the spreadsheet workers
        //parse out the programme info
        //set the program string
        //set the order in program
        private void setProgrammeAndSortOrder(string city)
        {
            if (FilmFormatter.Tools.SpreadSheetWorkers.programToSortOrder.ContainsKey(city))
                {
                    var programmeInfo = FilmFormatter.Tools.SpreadSheetWorkers.programToSortOrder[city];
                    string[] words = programmeInfo.Split('-');
                    this.programmeNumber = words[0];
                    if (words.Length > 1)
                    {
                    System.Int32.TryParse(words[1], out this.sortOrder);
                    }
                }      
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
