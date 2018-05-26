using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;

namespace FilmFormatter.Tools
{
	class SpreadsheetHelpers
	{

		public static int ColumnLetterToColumnIndex(string columnLetter)
		{
			columnLetter = columnLetter.ToUpper();
			int sum = 0;

			for (int i = 0; i < columnLetter.Length; i++)
			{
				sum *= 26;
				sum += (columnLetter[i] - 'A' + 1);
			}
			return sum-1;
		}


		private static string GetColumnAddress(string cellReference)
		{
			//Create a regular expression to get column address letters.
			Regex regex = new Regex("[A-Za-z]+");
			Match match = regex.Match(cellReference);
			return match.Value;
		}

		public static IEnumerable<Cell> GetCellsForRow(Row row, List<string> columnLetters)
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
	}
}
