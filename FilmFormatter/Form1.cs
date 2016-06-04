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

		Dictionary<string, int> titlesToRunTime = new Dictionary<string, int>();

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
					printSheetToConsole(sheetData, workbookPart);

				}
			}
		}

		private void printSheetToConsole(SheetData sheetData, WorkbookPart workbookPart)
		{
			foreach (Row r in sheetData.Elements<Row>())
			{
				foreach (Cell c in r.Elements<Cell>())
				{   //If it's not null, it's a string
					Console.WriteLine("CellValue is: " + c.CellValue);
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

		private Dictionary<String, int> GetTitlesFromRuntime(WorksheetPart worksheetpart, WorkbookPart workbookpart)
		{
			SheetData sheetData = worksheetpart.Worksheet.Elements<SheetData>().Last();
			Dictionary<string, int> ttrt = new Dictionary<string, int>();

			foreach (Row r in sheetData.Elements<Row>())
			{
				if (!r.Elements<Cell>().Any()) { Console.WriteLine("Found a small row"); }
				if (r.Elements<Cell>().Any())
				{
					Cell titleCell = r.Elements<Cell>().ElementAt(2);
					Cell runningTimeCell = r.Elements<Cell>().ElementAt(10);

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
							ttrt.Add(title, runningTime);

						}
					}
				}
			}
			Console.WriteLine("Opened main sheet");
			return ttrt;
		}
	}
}
