/*
Released under the [MIT License](http://www.opensource.org/licenses/mit-license.php)

Copyright (c) 2013 Doug Thompson

Permission is hereby granted, free of charge, to any person obtaining a
copy of this software and associated documentation files (the
"Software"),to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

The above copyright notice and this permission notice shall be included
in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
*/

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using ExcelXml;
using System.Data.SqlClient;

namespace ExcelXmlTestHarness
{
	public partial class frmMain : Form
	{
		public frmMain()
		{
			InitializeComponent();
		}

		private void frmMain_Load(object sender, EventArgs e)
		{
			DataReaderSheet();
			ManualSheet();
		}

		private void DataReaderSheet()
		{
			int i = 0;
			bool HeaderGenerated = false;
			WorksheetRow row = null;
			WorksheetColumn col = null;

			Workbook wb = new Workbook();
			wb.Properties.Created = DateTime.Now;

			WorksheetStyle style = wb.Styles.Add("Default");
			style.Font.FontName = "Tahoma";
			style.Font.Size = 10;

			style = wb.Styles.Add("HeaderStyle");
			style.Font.FontName = "Tahoma";
			style.Font.Size = 10;
			style.Font.Bold = true;

			ExcelWorksheet sheet = wb.WorkSheets.Add("Staff Data");
			string ConnectionString = "server=server; database=database; UID=uid ;PWD=pwd";
			using (SqlConnection conn = new SqlConnection(ConnectionString))
			{
				conn.Open();
				using (SqlCommand cmd = new SqlCommand("select * from <table>", conn))
				{
					using (SqlDataReader rdr = cmd.ExecuteReader())
					{
						if (rdr.HasRows)
						{
							while (rdr.Read())
							{
								row = sheet.Table.Rows.Add();
								if (!HeaderGenerated)
								{
									for (i = 0; i < rdr.FieldCount - 1; i++)
									{
										col = sheet.Table.Columns.Add();
										col.AutoFitWidth = true;
										row.Cells.Add(rdr.GetName(i), "HeaderStyle");
									}
									HeaderGenerated = true;
								}
								else
								{
									for (i = 0; i < rdr.FieldCount - 1; i++)
									{
										row.Cells.Add(rdr[i].ToString());
									}
								}
							}
						}
						string output = "";
						output = wb.SaveToFile("MyTestFile.xml");
					}
				}
			}
		}

		private void ManualSheet()
		{
			Workbook wb = new Workbook();
			wb.Properties.Created = DateTime.Now;

			WorksheetStyle style = wb.Styles.Add("Default");
			style.Font.FontName = "Tahoma";
			style.Font.Size = 10;

			style = wb.Styles.Add("HeaderStyle1");
			style.Font.FontName = "Tahoma";
			style.Font.Size = 14;
			style.Font.Bold = true;
			style.Alignment.Horizontal = StyleHorizontalAlignment.Center;
			style.Font.Color = Color.White;
			style.Interior.Color = Color.CadetBlue;
			style.Interior.Pattern = StyleInteriorPattern.HorzStripe;
			style.Interior.PatternColor = Color.Orange;

			style = wb.Styles.Add("HeaderStyle2");
			style.Font.FontName = "Tahoma";
			style.Font.Size = 14;
			style.Font.Bold = true;
			style.Alignment.Horizontal = StyleHorizontalAlignment.Center;
			style.Font.Color = Color.Orange;
			style.Interior.Color = Color.Blue;
			style.Interior.Pattern = StyleInteriorPattern.Solid;
			style.Interior.PatternColor = Color.Black;

			style = wb.Styles.Add("Footer");
			style.Font.FontName = "Tahoma";
			style.Font.Size = 12;
			style.Font.Bold = true;
			style.Font.Italic = true;
			StyleBorder border = style.Borders.Add();
			border.LineStyle = CellBorderLineStyle.DashDotDot;
			border.Position = CellBorderPosition.Bottom;
			border.Weight = 3.0;
			border.Color = Color.Red;
			border = style.Borders.Add();
			border.LineStyle = CellBorderLineStyle.Dash;
			border.Position = CellBorderPosition.Top;

			style = wb.Styles.Add("Currency");
			style.NumberFormat = "$ #0.00";
			style.Alignment.WrapText = true;

			ExcelWorksheet sheet = wb.WorkSheets.Add("NewSheet");
			sheet.Table.Columns.Add(new WorksheetColumn(150));
			sheet.Table.Columns.Add(new WorksheetColumn(100));

			WorksheetRow row = sheet.Table.Rows.Add();
			row.Cells.Add("Header 1", "HeaderStyle1");
			row.Cells.Add("Header 2", "HeaderStyle2");
			WorksheetCell cell = row.Cells.Add("Header 3");
			cell.MergeAcross = 1;			// Merge two cells together
			cell.StyleName = "HeaderStyle";

			row = sheet.Table.Rows.Add();
			// Skip one row, and add some text
			row.Index = 3;
			row.Cells.Add("Data");
			row.Cells.Add("Data 1");
			row.Cells.Add("Data 2");
			row.Cells.Add("Data 3");

			// Generate 30 rows
			for (int i = 0; i < 30; i++)
			{
				row = sheet.Table.Rows.Add();
				row.Cells.Add("Row " + i.ToString());
				cell = row.Cells.Add(i.ToString(), CellDataType.Number);
				cell.StyleName = "Currency";
			}

			// Add a Hyperlink
			row = sheet.Table.Rows.Add();
			row.StyleName = "Footer";
			cell = row.Cells.Add();
			cell.Data.Value = "My Intranet";
			cell.HRef = "http://www.google.com/search?q=Excel+XML";
			// Add a Formula for the above 30 rows
			cell = row.Cells.Add();
			cell.Formula = "=SUM(R[-30]C:R[-1]C)";

			wb.Names.Add("TestNamedRange", "NewSheet!R1C6:R4C10");

			string output = "";
			output = wb.SaveToFile("MyTestFile2.xml");
		}
	}
}