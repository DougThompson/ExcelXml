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
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Text;

namespace ExcelXml
{
	public enum StyleUnderline { None, Single, Double, SingleAccounting, DoubleAccounting };
	public enum StyleHorizontalAlignment { CenterAcrossSelection, Fill, Left, Right, Justify, Distributed, Center, Automatic, JustifyDistributed };
	public enum StyleVerticalAlignment { Automatic, Top, Bottom, Center, Justify, Distributed, JustifyDistributed };
	public enum StyleInteriorPattern { None, Solid, Gray75, Gray50, Gray25, Gray125, Gray0625, HorzStripe, VertStripe, ReverseDiagStripe, DiagStripe, DiagCross, ThickDiagCross, ThinHorzStripe, ThinVertStripe, ThinReverseDiagStripe, ThinDiagStripe, ThinHorzCross, ThinDiagCross };
	public enum CellDataType { Number, DateTime, Boolean, String, Error };
	public enum CellFormat { General, GeneralNumber, GeneralDate, LongDate, MediumDate, ShortDate, LongTime, MediumTime, ShortTime, Currency, EuroCurrency, Fixed, Standard, Percent, Scientific, YesNo, TrueFalse, OnOff };
	public enum CellBorderPosition { Left, Top, Right, Bottom, DiagonalLeft, DiagonalRight };
	public enum CellBorderLineStyle { None, Continuous, Dash, Dot, DashDot, DashDotDot, SlantDashDot, Double };

	public class Workbook
	{
		private DocumentProperties _Properties = null;
		private ExcelWorkbookDetails _ExcelWorkbook = null;
		private SheetStyles _SheetStyles = null;
		private ExcelWorksheets _Worksheets = null;
		private NamedRanges _Names = null;

		private string[,] CellFormats = new string[,] { { "General", "General" }, { "GeneralNumber", "General Number" }, { "GeneralDate", "General Date" }, { "LongDate", "Long Date" }, { "MediumDate", "Medium Date" }, { "ShortDate", "Short Date" }, { "LongTime", "Long Time" }, { "MediumTime", "Medium Time" }, { "ShortTime", "Short Time" }, { "Currency", "Currency" }, { "EuroCurrency", "Euro Currency" }, { "Fixed", "Fixed" }, { "Standard", "Standard" }, { "Percent", "Percent" }, { "Scientific", "Scientific" }, { "YesNo", "Yes/No" }, { "TrueFalse", "True/False" }, { "OnOff", "On/Off" } };

		public Workbook()
		{
			_Properties = new DocumentProperties();
			_ExcelWorkbook = new ExcelWorkbookDetails();
			_SheetStyles = new SheetStyles();
			_Worksheets = new ExcelWorksheets();
			_Names = new NamedRanges();
		}

		public string convertToHexColor(Color c)
		{
			return ("#" + c.ToArgb().ToString("x").Substring(2)).ToUpper();
		}

		public string XmlEncode(string input)
		{
			input = input.Replace("\"", "&quot;");
			input = input.Replace("&", "&amp;");
			input = input.Replace("'", "&apos;");
			input = input.Replace("<", "&lt;");
			input = input.Replace(">", "&gt;");

			return input;
		}

		public string GetNumberFormat(string format)
		{
			string FormatLookup = format;

			for (int i = 0; i < 18; i++)
			{
				if (CellFormats[i, 0] == format)
				{
					FormatLookup = CellFormats[i, 1];
				}
			}

			return FormatLookup;
		}

		public string GetStyleID(string StyleName)
		{
			string StyleID = "";

			foreach (WorksheetStyle sht in _SheetStyles.Styles)
			{
				if (sht.Name == StyleName)
				{
					StyleID = sht.ID;
				}
			}

			return StyleID;
		}

		public string IsInNamedRange(ExcelWorksheet sheet, int Row, int Column)
		{
			string chunk = "";
			string refersTo = "";
			string min = "", max = "";
			int minRow = 0, minCol = 0;
			int maxRow = 0, maxCol = 0;

			foreach (NamedRange range in _Names.Collection)
			{
				refersTo = range.RefersTo.Substring(0, range.RefersTo.IndexOf("!"));
				if (refersTo.ToLower() == sheet.Name.ToLower())
				{
					chunk = range.RefersTo.Substring(range.RefersTo.IndexOf("!") + 1);
					min = chunk.Split(':')[0];
					max = chunk.Split(':')[1];

					minRow = int.Parse(min.Substring(1, min.IndexOf('C') - 1));
					minCol = int.Parse(min.Substring(min.IndexOf('C') + 1));
					maxRow = int.Parse(max.Substring(1, max.IndexOf('C') - 1));
					maxCol = int.Parse(max.Substring(max.IndexOf('C') + 1));

					if ((Row >= minRow && Row <= maxRow) && (Column >= minCol && Column <= maxCol))
					{
						return range.Name;
					}
				}
			}

			return "";
		}

		public int GetExpandedRowCount(ExcelWorksheet worksheet)
		{
			int count = 0;
			int index = 0;

			foreach (WorksheetRow row in worksheet.Table.Rows.Collection)
			{
				count++;
				if (row.Index > 0)
				{
					if (row.Index > index)
					{
						index = row.Index;
						count = row.Index;
					}
				}
			}

			if (index > count)
			{
				return index;
			}
			else
			{
				return count;
			}
		}

		public int GetExpandedColumnCount(ExcelWorksheet worksheet)
		{
			int index = 0;
			int count = 0;

			foreach (WorksheetRow row in worksheet.Table.Rows.Collection)
			{
				count = 0;
				foreach (WorksheetCell cell in row.Cells.Collection)
				{
					count++;
					if (cell.Index > 0)
					{
						if (cell.Index > index)
						{
							index = cell.Index;
						}
					}
				}
				if (count > index)
				{
					index = count;
				}
			}

			if (index > count)
			{
				return index;
			}
			else
			{
				return count;
			}
		}

		public StringBuilder GetHeader(StringBuilder sb)
		{
			sb.Append("<?xml version=\"1.0\"?>\n");
			sb.Append("<?mso-application progid=\"Excel.Sheet\"?>\n");
			sb.Append("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:html=\"http://www.w3.org/TR/REC-html40\">\n");

			return sb;
		}

		public StringBuilder GetDocumentProperties(StringBuilder sb)
		{
			sb.Append("<DocumentProperties xmlns=\"urn:schemas-microsoft-com:office:office\">\n");
			sb.Append(String.Format("\t<Author>{0}</Author>\n", XmlEncode(_Properties.Author)));
			sb.Append(String.Format("\t<LastAuthor>{0}</LastAuthor>\n", XmlEncode(_Properties.LastAuthor)));
			sb.Append(String.Format("\t<Created>{0}</Created>\n", _Properties.Created));
			sb.Append(String.Format("\t<Company>{0}</Company>\n", XmlEncode(_Properties.Company)));
			sb.Append(String.Format("\t<Version>{0}</Version>\n", XmlEncode(_Properties.Version)));
			sb.Append("</DocumentProperties>\n");

			return sb;
		}

		public StringBuilder GetExcelWorkbook(StringBuilder sb)
		{
			sb.Append("<ExcelWorkbook xmlns=\"urn:schemas-microsoft-com:office:excel\">\n");
			sb.Append(String.Format("\t<WindowHeight>{0}</WindowHeight>\n", _ExcelWorkbook.WindowHeight.ToString()));
			sb.Append(String.Format("\t<WindowWidth>{0}</WindowWidth>\n", _ExcelWorkbook.WindowWidth.ToString()));
			sb.Append(String.Format("\t<WindowTopX>{0}</WindowTopX>\n", _ExcelWorkbook.WindowTopX.ToString()));
			sb.Append(String.Format("\t<WindowTopY>{0}</WindowTopY>\n", _ExcelWorkbook.WindowTopY.ToString()));
			sb.Append(String.Format("\t<ProtectStructure>{0}</ProtectStructure>\n", _ExcelWorkbook.ProtectStructure.ToString()));
			sb.Append(String.Format("\t<ProtectWindows>{0}</ProtectWindows>\n", _ExcelWorkbook.ProtectWindows.ToString()));
			sb.Append("</ExcelWorkbook>\n");

			return sb;
		}

		public StringBuilder GetStyles(StringBuilder sb)
		{
			sb.Append("<Styles>\n");

			foreach (WorksheetStyle style in _SheetStyles.Styles)
			{
				sb.Append("\t<Style");
				if (style.ID != "")
				{
					sb.Append(String.Format(" ss:ID=\"{0}\"", style.ID));
				}
				if (style.Name != "")
				{
					sb.Append(String.Format(" ss:Name=\"{0}\"", XmlEncode(style.Name)));
				}
				sb.Append(">\n");

				sb.Append("\t\t<Alignment");
				sb.Append(String.Format(" ss:Horizontal=\"{0}\"", style.Alignment.Horizontal.ToString()));
				sb.Append(String.Format(" ss:Vertical=\"{0}\"", style.Alignment.Vertical.ToString()));
				if (style.Alignment.Rotate != 0)
				{
					if (style.Alignment.Rotate < -90)
					{
						style.Alignment.Rotate = -90;
					}

					if (style.Alignment.Rotate > 90)
					{
						style.Alignment.Rotate = 90;
					}

					sb.Append(String.Format(" ss:Rotate=\"{0}\"", style.Alignment.Rotate.ToString()));
				}
				if (style.Alignment.Indent > 0)
				{
					sb.Append(String.Format(" ss:Indent=\"{0}\"", style.Alignment.Indent.ToString()));
				}
				if (style.Alignment.WrapText == true)
				{
					sb.Append(" ss:WrapText=\"1\"");
				}
				if (style.Alignment.ShrinkToFit == true)
				{
					sb.Append(" ss:ShrinkToFit=\"1\"");
				}
				if (style.Alignment.VerticalText == true)
				{
					sb.Append(" ss:VerticalText=\"1\"");
				}

				sb.Append(" />\n");

				sb.Append("\t\t<Font");
				//sb.Append(" x:Family=\"Swiss\"");
				if (style.Font.Size > 0)
				{
					sb.Append(String.Format(" ss:Size=\"{0}\"", style.Font.Size.ToString()));
				}

				if (style.Font.FontName != "")
				{
					sb.Append(String.Format(" ss:FontName=\"{0}\"", style.Font.FontName));
				}

				if (style.Font.Color != Color.Black)
				{
					sb.Append(String.Format(" ss:Color=\"{0}\"", convertToHexColor(style.Font.Color)));
				}

				if (style.Font.Bold == true)
				{
					sb.Append(" ss:Bold=\"1\"");
				}

				if (style.Font.Italic == true)
				{
					sb.Append(" ss:Italic=\"1\"");
				}

				if (style.Font.Outline == true)
				{
					sb.Append(" ss:Outline=\"1\"");
				}

				if (style.Font.Shadow == true)
				{
					sb.Append(" ss:Shadow=\"1\"");
				}

				if (style.Font.StrikeThrough == true)
				{
					sb.Append(" ss:StrikeThrough=\"1\"");
				}

				if (style.Font.Underline != StyleUnderline.None)
				{
					sb.Append(" ss:Underline=\"" + style.Font.Underline.ToString());
				}

				sb.Append(" />\n");

				sb.Append("\t\t<Interior");
				//sb.Append(" x:Family=\"Swiss\"");
				if (style.Interior.Color != System.Drawing.Color.White)
				{
					sb.Append(String.Format(" ss:Color=\"{0}\"", convertToHexColor(style.Interior.Color)));
				}

				if (style.Interior.Pattern != StyleInteriorPattern.None)
				{
					sb.Append(String.Format(" ss:Pattern=\"{0}\"", style.Interior.Pattern));
					if (style.Interior.Pattern != StyleInteriorPattern.Solid)
					{
						sb.Append(String.Format(" ss:PatternColor=\"{0}\"", convertToHexColor(style.Interior.PatternColor)));
					}

				}
				sb.Append(" />\n");

				if ((style.NumberFormat != "General") && (style.NumberFormat != ""))
				{
					sb.Append(String.Format("\t\t<NumberFormat ss:Format=\"{0}\" />\n", GetNumberFormat(style.NumberFormat)));
				}


				sb.Append("\t\t<Borders>\n");
				foreach (StyleBorder border in style.Borders.Collection)
				{
					sb.Append("\t\t\t<Border");
					sb.Append(String.Format(" ss:Position=\"{0}\"", border.Position.ToString()));
					sb.Append(String.Format(" ss:LineStyle=\"{0}\"", border.LineStyle.ToString()));
					sb.Append(String.Format(" ss:Weight=\"{0}\"", border.Weight.ToString()));
					if (border.Color != Color.Black)
					{
						sb.Append(String.Format(" ss:Color=\"{0}\"", convertToHexColor(border.Color)));
					}

					sb.Append(" />\n");
				}
				sb.Append("\t\t</Borders>\n");

				sb.Append("\t\t<Protection />\n");
				sb.Append("\t</Style>\n");
			}
			sb.Append("</Styles>\n");

			return sb;
		}

		public StringBuilder GetNames(StringBuilder sb)
		{
			sb.Append("<Names>\n");

			foreach (NamedRange name in _Names.Collection)
			{
				sb.Append(String.Format("\t<NamedRange ss:Name=\"{0}\" ss:RefersTo=\"={1}\"/>\n", name.Name, name.RefersTo));
			}
			sb.Append("</Names>\n");

			return sb;
		}

		public StringBuilder GetWorksheets(StringBuilder sb)
		{
			int curRow = 0, curCol = 0;
			string namedRange = "";

			foreach (ExcelWorksheet worksheet in _Worksheets.Worksheets)
			{
				int exRows = GetExpandedRowCount(worksheet);
				int exCols = GetExpandedColumnCount(worksheet);

				sb.Append(String.Format("<Worksheet ss:Name=\"{0}\">\n", worksheet.Name));
				sb.Append(String.Format("\t<Table ss:ExpandedColumnCount=\"{0}\" ss:ExpandedRowCount=\"{1}\" x:FullColumns=\"1\" x:FullRows=\"1\">\n", exCols.ToString(), exRows.ToString()));

				foreach (WorksheetColumn col in worksheet.Table.Columns.Collection)
				{
					sb.Append("\t\t<Column");
					if (col.Index > 0)
					{
						sb.Append(String.Format(" ss:Index=\"{0}\"", col.Index.ToString()));
					}

					if (col.AutoFitWidth == true)
					{
						sb.Append(" ss:AutoFitWidth=\"1\"");
					}

					if (col.Hidden == true)
					{
						sb.Append(" ss:Hidden=\"1\"");
					}

					if (col.Width > 0)
					{
						sb.Append(String.Format(" ss:Width=\"{0}\"", col.Width.ToString()));
					}

					if (GetStyleID(col.StyleName) != "")
					{
						sb.Append(String.Format(" ss:StyleID=\"{0}\"", GetStyleID(col.StyleName)));
					}

					sb.Append(" />\n");
				}

				curRow = 0;
				foreach (WorksheetRow row in worksheet.Table.Rows.Collection)
				{
					sb.Append("\t\t<Row");
					if (row.Index > 0)
					{
						sb.Append(String.Format(" ss:Index=\"{0}\"", row.Index.ToString()));
						curRow = row.Index;
					}
					else
					{
						curRow++;
					}
					if (row.AutoFitHeight == true)
					{
						sb.Append(" ss:AutoFitHeight=\"1\"");
					}

					if (row.Hidden == true)
					{
						sb.Append(" ss:Hidden=\"1\"");
					}

					if (row.Height > 0)
					{
						sb.Append(String.Format(" ss:Height=\"{0}\"", row.Height.ToString()));
					}

					if (GetStyleID(row.StyleName) != "")
					{
						sb.Append(String.Format(" ss:StyleID=\"{0}\"", GetStyleID(row.StyleName)));
					}

					sb.Append(">\n");

					curCol = 0;
					foreach (WorksheetCell cell in row.Cells.Collection)
					{
						sb.Append("\t\t\t<Cell");
						if (cell.Index > 0)
						{
							sb.Append(String.Format(" ss:Index=\"{0}\"", cell.Index.ToString()));
							curCol = cell.Index;
						}
						else
						{
							curCol++;
						}
						if (cell.MergeAcross > 0)
						{
							sb.Append(String.Format(" ss:MergeAcross=\"{0}\"", cell.MergeAcross.ToString()));
						}

						if (cell.MergeDown > 0)
						{
							sb.Append(String.Format(" ss:MergeDown=\"{0}\"", cell.MergeDown.ToString()));
						}

						if (GetStyleID(cell.StyleName) != "")
						{
							sb.Append(String.Format(" ss:StyleID=\"{0}\"", GetStyleID(cell.StyleName)));
						}

						if (cell.Formula != "")
						{
							sb.Append(String.Format(" ss:Formula=\"{0}\"", cell.Formula));
						}

						if (cell.HRef != "")
						{
							sb.Append(String.Format(" ss:HRef=\"{0}\"", cell.HRef));
						}

						sb.Append("><Data");
						sb.Append(String.Format(" ss:Type=\"{0}\">", cell.Data.Type.ToString()));
						sb.Append(XmlEncode(cell.Data.Value));
						sb.Append("</Data>");
						namedRange = IsInNamedRange(worksheet, curRow, curCol);
						if (namedRange != "")
						{
							sb.Append(String.Format("<NamedCell ss:Name=\"{0}\"/>", namedRange));
						}

						sb.Append("</Cell>\n");
					}
					sb.Append("\t\t</Row>\n");
				}
				sb.Append("\t</Table>\n");
				sb.Append("\t<WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\">\n");
				sb.Append("\t\t<Print>\n");
				sb.Append("\t\t\t<ValidPrinterInfo/>\n");
				sb.Append("\t\t\t<HorizontalResolution>600</HorizontalResolution>\n");
				sb.Append("\t\t\t<VerticalResolution>600</VerticalResolution>\n");
				sb.Append("\t\t</Print>\n");
				sb.Append("\t\t<Selected/>\n");
				sb.Append("\t\t<ProtectObjects>False</ProtectObjects>\n");
				sb.Append("\t\t<ProtectScenarios>False</ProtectScenarios>\n");
				sb.Append("\t</WorksheetOptions>\n");
				sb.Append("</Worksheet>\n");
			}

			return sb;
		}

		public bool SetData(DataSet ds, string TabName)
		{
			bool result = false;

			int i = 0, j = 0;
			bool HeaderGenerated = false;
			WorksheetRow row = null;
			WorksheetColumn col = null;

			WorksheetStyle style = this.Styles.Add("Default");

			style = this.Styles.Add("_HeaderStyle");
			style.Font.Bold = true;

			style = this.Styles.Add("_DateTime");
			style.NumberFormat = @"[$-409]m/d/yy\ h:mm\ AM/PM;@";

			style = this.Styles.Add("_Integer");
			style.NumberFormat = "0";

			ExcelWorksheet sheet = this.WorkSheets.Add(TabName);

			DataTable dt = new DataTable();
			dt = ds.Tables[0];

			if (dt.Rows.Count > 0)
			{
				for (j = 0; j <= dt.Rows.Count - 1; j++)
				{
					row = sheet.Table.Rows.Add();
					if (!HeaderGenerated)
					{
						for (i = 0; i < dt.Columns.Count - 1; i++)
						{
							col = sheet.Table.Columns.Add();
							col.AutoFitWidth = true;
							col.Width = 99.75;

							switch (dt.Rows[j][i].GetType().Name)
							{
								case "DateTime":
									col.StyleName = "_DateTime";
									break;
								case "Int16":
									col.StyleName = "_Integer";
									break;
							}

							row.Cells.Add(dt.Rows[j][i].ToString(), "_HeaderStyle");
						}
						HeaderGenerated = true;
					}
					else
					{
						WorksheetCell cell = null;

						for (i = 0; i < dt.Columns.Count - 1; i++)
						{
							switch (dt.Rows[j][i].GetType().Name)
							{
								case "DateTime":
									cell = new WorksheetCell();
									if (dt.Rows[j][i] != System.DBNull.Value)
										cell = row.Cells.Add(((DateTime)dt.Rows[j][i]).Date.ToString("yyyy-MM-ddTHH:mm:ss.fff"), CellDataType.DateTime);
									else
										cell = row.Cells.Add("", CellDataType.String);
									break;
								case "Int16":
									cell = new WorksheetCell();
									if (dt.Rows[j][i] != System.DBNull.Value)
										cell = row.Cells.Add(dt.Rows[j][i].ToString(), CellDataType.Number);
									else
										cell = row.Cells.Add("", CellDataType.String);
									break;
								default:
									row.Cells.Add(dt.Rows[j][i].ToString());
									break;
							}
						}
					}
				}
				result = true;
			}
			else
			{
				result = false;
			}

			return result;
		}

		public bool SetData(SqlDataReader reader, string TabName)
		{
			bool result = false;

			int i = 0;
			bool HeaderGenerated = false;
			WorksheetRow row = null;
			WorksheetColumn col = null;

			WorksheetStyle style = this.Styles.Add("Default");

			style = this.Styles.Add("_HeaderStyle");
			style.Font.Bold = true;

			style = this.Styles.Add("_DateTime");
			style.NumberFormat = @"[$-409]m/d/yy\ h:mm\ AM/PM;@";

			style = this.Styles.Add("_Integer");
			style.NumberFormat = "0";

			ExcelWorksheet sheet = this.WorkSheets.Add(TabName);

			if (reader.HasRows)
			{
				while (reader.Read())
				{
					row = sheet.Table.Rows.Add();
					if (!HeaderGenerated)
					{
						for (i = 0; i < reader.FieldCount - 1; i++)
						{
							col = sheet.Table.Columns.Add();
							col.AutoFitWidth = true;
							col.Width = 99.75;

							switch (reader.GetFieldType(i).Name)
							{
								case "DateTime":
									col.StyleName = "_DateTime";
									break;
								case "Int16":
									col.StyleName = "_Integer";
									break;
							}

							row.Cells.Add(reader.GetName(i), "_HeaderStyle");
						}
						HeaderGenerated = true;
					}
					else
					{
						WorksheetCell cell = null;

						for (i = 0; i < reader.FieldCount - 1; i++)
						{
							switch (reader.GetFieldType(i).Name)
							{
								case "DateTime":
									cell = new WorksheetCell();
									if (reader[i] != System.DBNull.Value)
										cell = row.Cells.Add(reader.GetDateTime(i).ToString("yyyy-MM-ddTHH:mm:ss.fff"), CellDataType.DateTime);
									else
										cell = row.Cells.Add("", CellDataType.String);
									break;
								case "Int16":
									cell = new WorksheetCell();
									if (reader[i] != System.DBNull.Value)
										cell = row.Cells.Add(reader.GetInt16(i).ToString(), CellDataType.Number);
									else
										cell = row.Cells.Add("", CellDataType.String);
									break;
								default:
									row.Cells.Add(reader[i].ToString());
									break;
							}
						}
					}
				}
				result = true;
			}
			else
			{
				result = false;
			}

			return result;
		}

		public string Generate()
		{
			StringBuilder sb = new StringBuilder();

			sb = GetHeader(sb);
			sb = GetDocumentProperties(sb);
			sb = GetExcelWorkbook(sb);
			sb = GetStyles(sb);
			sb = GetNames(sb);
			sb = GetWorksheets(sb);
			sb.Append("</Workbook>");

			return sb.ToString();
		}

		public string Generate(System.Web.HttpResponse response, string FileName, bool OpenInBrowser)
		{
			StringBuilder sb = new StringBuilder();

			sb = GetHeader(sb);
			sb = GetDocumentProperties(sb);
			sb = GetExcelWorkbook(sb);
			sb = GetStyles(sb);
			sb = GetNames(sb);
			sb = GetWorksheets(sb);
			sb.Append("</Workbook>");

			response.Buffer = false;
			response.Clear();
			response.ContentType = "application/vnd.ms-excel";
			if (!OpenInBrowser)
				response.AddHeader("Content-Disposition", "attachment; filename=" + FileName);
			System.Text.ASCIIEncoding encoding = new System.Text.ASCIIEncoding();
			response.OutputStream.Write(encoding.GetBytes(sb.ToString()), 0, encoding.GetBytes(sb.ToString()).Length);
			response.Flush();
			response.End();

			return sb.ToString();
		}

		public string SaveToFile(string FileName)
		{
			string Excel = Generate();
			using (FileStream fs = new FileStream(FileName, FileMode.Create))
			{
				using (StreamWriter objWriter = new StreamWriter(fs))
				{
					objWriter.Write(Excel);
				}
			}
			return Excel;
		}

		public ExcelWorksheets WorkSheets
		{
			get
			{ return _Worksheets; }

			set
			{ _Worksheets = value; }
		}

		public DocumentProperties Properties
		{
			get
			{ return _Properties; }

			set
			{ _Properties = value; }
		}

		public ExcelWorkbookDetails ExcelWorkbook
		{
			get
			{ return _ExcelWorkbook; }

			set
			{ _ExcelWorkbook = value; }
		}

		public SheetStyles Styles
		{
			get
			{ return _SheetStyles; }

			set
			{ _SheetStyles = value; }
		}

		public NamedRanges Names
		{
			get
			{ return _Names; }
			set
			{ _Names = value; }
		}

		public class DocumentProperties
		{
			public string Author = "Excel XML Test Harness";
			public string LastAuthor = "Excel XML Test Harness";
			public string Company = "Your Company";
			public string Version = "11.8132";
			private string _Created = "";

			public DocumentProperties()
			{
			}

			public DateTime Created
			{
				set
				{ _Created = value.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ"); }

				get
				{ return DateTime.Parse(_Created); }
			}
		}

		public class ExcelWorkbookDetails
		{
			public int ActiveSheetIndex = 1;
			public int WindowTopX = 120;
			public int WindowTopY = 60;
			public int WindowHeight = 7995;
			public int WindowWidth = 12000;
			public bool ProtectStructure = false;
			public bool ProtectWindows = false;

			public ExcelWorkbookDetails()
			{
			}
		}

		public class NamedRanges
		{
			private List<NamedRange> _Names = new List<NamedRange>();

			public NamedRanges()
			{
			}

			public NamedRange Add()
			{
				NamedRange name = new NamedRange();
				_Names.Add(name);

				return name;
			}

			public NamedRange Add(string Name, string RefersTo)
			{
				NamedRange name = new NamedRange();
				name.Name = Name;
				name.RefersTo = RefersTo;

				_Names.Add(name);
				return name;
			}

			public List<NamedRange> Collection
			{
				get
				{ return _Names; }
				set
				{ _Names = value; }
			}

		}

		public class SheetStyles
		{
			private List<WorksheetStyle> _Styles = new List<WorksheetStyle>();

			public SheetStyles()
			{
			}

			private string GetNextID()
			{
				string sID = "";
				int id = 20;

				foreach (WorksheetStyle sht in _Styles)
				{
					if (sht.ID.StartsWith("s"))
					{
						sID = sht.ID;
						if (id <= int.Parse(sID.Substring(1)))
						{
							id = int.Parse(sID.Substring(1)) + 1;
						}
					}
				}
				sID = "s" + id.ToString();

				return sID;
			}

			public WorksheetStyle Add(string StyleName)
			{
				WorksheetStyle style = new WorksheetStyle();

				if (StyleName == "Default")
				{
					style.Name = "Normal";
					style.ID = "Default";
				}
				else
				{
					style.Name = StyleName;
					style.ID = GetNextID();
				}

				_Styles.Add(style);
				return style;
			}

			public List<WorksheetStyle> Styles
			{
				get
				{ return _Styles; }

				set
				{ _Styles = value; }
			}
		}

		public class ExcelWorksheets
		{
			private List<ExcelWorksheet> _Worksheets = new List<ExcelWorksheet>();

			public ExcelWorksheets()
			{
			}

			private string CheckWorksheetName(string TabName)
			{
				int j = 1;
				bool matchFound = true;

				while (matchFound)
				{
					matchFound = false;
					foreach (ExcelWorksheet sheet in _Worksheets)
					{
						if (TabName == sheet.Name)
						{
							matchFound = true;
							if (!TabName.EndsWith(")"))
							{
								TabName += String.Format(" ({0})", j.ToString());
							}
							else
							{
								TabName = TabName.Replace(String.Format("({0})", (j - 1).ToString()), String.Format("({0})", j.ToString()));
							}
							j++;
							break;
						}
					}
				}

				return TabName;
			}

			public List<ExcelWorksheet> Worksheets
			{
				get
				{ return _Worksheets; }
				set
				{ _Worksheets = value; }
			}

			public ExcelWorksheet Add()
			{
				string TabName = "Sheet 1";

				ExcelWorksheet worksheet = new ExcelWorksheet();
				TabName = CheckWorksheetName(TabName);
				worksheet.Name = TabName;
				_Worksheets.Add(worksheet);
				return worksheet;
			}

			public ExcelWorksheet Add(string TabName)
			{
				ExcelWorksheet worksheet = new ExcelWorksheet();
				TabName = CheckWorksheetName(TabName);
				worksheet.Name = TabName;
				_Worksheets.Add(worksheet);
				return worksheet;
			}
		}

	}

	public class WorksheetStyle
	{
		public string Name = "";
		public string ID = "";
		public FontType Font = new FontType();
		public StyleAlignment Alignment = new StyleAlignment();
		public StyleInterior Interior = new StyleInterior();
		public StyleBorders Borders = new StyleBorders();
		public string NumberFormat = "General";

		public WorksheetStyle()
		{
		}

		public class FontType
		{
			public string FontName = "";
			public System.Drawing.Color Color = Color.Black;
			public double Size = 0;
			public bool Bold = false;
			public bool Italic = false;
			public bool Outline = false;
			public bool Shadow = false;
			public bool StrikeThrough = false;
			public StyleUnderline Underline = StyleUnderline.None;
		}

		public class StyleAlignment
		{
			public StyleHorizontalAlignment Horizontal = StyleHorizontalAlignment.Automatic;
			public StyleVerticalAlignment Vertical = StyleVerticalAlignment.Automatic;
			public double Rotate = 0;
			public bool WrapText = false;
			public long Indent = 0;
			public bool ShrinkToFit = false;
			public bool VerticalText = false;
		}

		public class StyleInterior
		{
			public System.Drawing.Color Color = System.Drawing.Color.White;
			public System.Drawing.Color PatternColor = System.Drawing.Color.Black;
			public StyleInteriorPattern Pattern = StyleInteriorPattern.None;
		}

		public class StyleBorders
		{
			private List<StyleBorder> _Borders = new List<StyleBorder>();

			public StyleBorder Add()
			{
				StyleBorder border = new StyleBorder();
				_Borders.Add(border);

				return border;
			}

			public List<StyleBorder> Collection
			{
				get
				{ return _Borders; }

				set
				{ _Borders = value; }
			}

			public StyleBorder Add(CellBorderPosition Position, CellBorderLineStyle LineStyle, Color Color, double Weight)
			{
				StyleBorder border = new StyleBorder();
				border.Color = Color;
				border.LineStyle = LineStyle;
				border.Position = Position;
				border.Weight = Weight;
				_Borders.Add(border);

				return border;
			}
		}
	}

	public class StyleBorder
	{
		public CellBorderPosition Position;
		public System.Drawing.Color Color = System.Drawing.Color.Black;
		public CellBorderLineStyle LineStyle;
		public double Weight = 1;
	}

	public class ExcelWorksheet
	{
		public string Name = "";
		public SheetTable Table = null;

		public ExcelWorksheet()
		{
			Table = new SheetTable();
		}

		public class SheetTable
		{
			public SheetRows Rows = null;
			public SheetColumns Columns = null;

			public SheetTable()
			{
				Rows = new SheetRows();
				Columns = new SheetColumns();
			}

			public class SheetRows
			{
				private List<WorksheetRow> _Rows = new List<WorksheetRow>();

				public SheetRows()
				{
				}

				public List<WorksheetRow> Collection
				{
					get
					{ return _Rows; }
					set
					{ _Rows = value; }
				}

				public WorksheetRow Add()
				{
					WorksheetRow row = new WorksheetRow();
					_Rows.Add(row);
					return row;
				}

				public WorksheetRow Add(WorksheetRow Row)
				{
					_Rows.Add(Row);
					return Row;
				}

				public WorksheetRow Add(double Height)
				{
					WorksheetRow row = new WorksheetRow();
					row.Height = Height;
					_Rows.Add(row);
					return row;
				}
			}

			public class SheetColumns
			{
				private List<WorksheetColumn> _Columns = new List<WorksheetColumn>();

				public SheetColumns()
				{
				}

				public List<WorksheetColumn> Collection
				{
					get
					{ return _Columns; }
					set
					{ _Columns = value; }
				}

				public WorksheetColumn Add()
				{
					WorksheetColumn col = new WorksheetColumn();
					_Columns.Add(col);
					return col;
				}

				public WorksheetColumn Add(WorksheetColumn Column)
				{
					_Columns.Add(Column);
					return Column;
				}

				public WorksheetColumn Add(double Width)
				{
					WorksheetColumn col = new WorksheetColumn();
					col.Width = Width;
					_Columns.Add(col);
					return col;
				}
			}
		}
	}

	public class WorksheetRow
	{
		public bool AutoFitHeight = false;
		public double Height = 0;
		public bool Hidden = false;
		public int Index = 0;
		public int Span = 0;
		public string StyleName = "";
		public WorksheetCells Cells = null;

		public WorksheetRow()
		{
			Cells = new WorksheetCells();
		}

		public WorksheetRow(double RowHeight)
		{
			Height = RowHeight;
			Cells = new WorksheetCells();
		}

		public class WorksheetCells
		{
			private List<WorksheetCell> _Cells = new List<WorksheetCell>();

			public List<WorksheetCell> Collection
			{
				get
				{ return _Cells; }
				set
				{ _Cells = value; }
			}

			public WorksheetCell Add()
			{
				WorksheetCell cell = new WorksheetCell();
				_Cells.Add(cell);
				return cell;
			}

			public WorksheetCell Add(WorksheetCell Cell)
			{
				_Cells.Add(Cell);
				return Cell;
			}

			public WorksheetCell Add(string Value)
			{
				WorksheetCell cell = new WorksheetCell();
				cell.Data.Value = Value;
				_Cells.Add(cell);
				return cell;
			}

			public WorksheetCell Add(string Value, CellDataType DataType)
			{
				WorksheetCell cell = new WorksheetCell();
				cell.Data.Value = Value;
				cell.Data.Type = DataType;
				_Cells.Add(cell);
				return cell;
			}

			public WorksheetCell Add(string Value, string StyleName)
			{
				WorksheetCell cell = new WorksheetCell();
				cell.Data.Value = Value;
				cell.StyleName = StyleName;
				_Cells.Add(cell);
				return cell;
			}

			public WorksheetCell Add(string Value, string StyleName, CellDataType DataType)
			{
				WorksheetCell cell = new WorksheetCell();
				cell.Data.Value = Value;
				cell.StyleName = StyleName;
				cell.Data.Type = DataType;
				_Cells.Add(cell);
				return cell;
			}

			public WorksheetCell Add(string Value, string StyleName, string NamedCell)
			{
				WorksheetCell cell = new WorksheetCell();
				cell.Data.Value = Value;
				cell.StyleName = StyleName;
				cell.Data.NamedCell = NamedCell;
				_Cells.Add(cell);
				return cell;
			}
		}
	}

	public class WorksheetColumn
	{
		public bool AutoFitWidth = false;
		public bool Hidden = false;
		public int Index = 0;
		public int Span = 0;
		public string StyleName = "";
		public double Width = 0;

		public WorksheetColumn()
		{
		}

		public WorksheetColumn(double ColumnWidth)
		{
			Width = ColumnWidth;
		}
	}

	public class WorksheetCell
	{
		public string ArrayRange = "";
		public string Formula = "";
		public string HRef = "";
		public int Index = 0;
		public long MergeAcross = 0;
		public long MergeDown = 0;
		public string StyleName = "";
		public CellData Data = new CellData();

		public class CellData
		{
			public string Value = "";
			public string NamedCell = "";
			public CellDataType Type = CellDataType.String;
		}
	}

	public class NamedRange
	{
		public string Name = "";
		public string RefersTo = "";
	}
}
