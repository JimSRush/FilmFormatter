using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
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
            parseFile(file);
           
        }

        private void parseFile(String fileName) 
        {
            Console.WriteLine("parsing file or something");
            using (SpreadsheetDocument myDoc = SpreadsheetDocument.Open(fileName, true)) 
            {
                WorkbookPart workbookPart = myDoc.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().Last();
                Console.WriteLine("Opened sheet");

 
                foreach (Row r in sheetData.Elements<Row>())
                {
                    foreach (Cell c in r.Elements<Cell>())
                    {
                        if ((c.DataType != null) && (c.DataType == CellValues.SharedString))
                        {
                            String text = workbookPart.SharedStringTablePart.SharedStringTable
                                .Elements<SharedStringItem>().ElementAt(
                                    Convert.ToInt32(c.CellValue.Text)).InnerText;
                            Console.WriteLine(text);
                        }
                        //Console.WriteLine(c.DataType);
                    }
                }
            }
        }
   } 
}
