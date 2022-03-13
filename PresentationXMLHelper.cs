// -----------------------------------------------------------------------
// <copyright file="PresentationHelper.cs" company="">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

namespace PresentationGenerator
{
	using System;
	using System.Collections.Generic;
	using System.Collections.Specialized;
	using System.Data;
	using System.Linq;
	using System.Text;
	using System.Text.RegularExpressions;
	using System.IO;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Drawing;
	using A = DocumentFormat.OpenXml.Drawing;
	using DocumentFormat.OpenXml.Packaging;
	using DocumentFormat.OpenXml.Presentation;
	using SheetML = DocumentFormat.OpenXml.Spreadsheet;

	/// <summary>
	/// Open Office XML Presentation Methods
	/// </summary>
	public class PresentationXMLHelper : IPresentation, IDisposable
	{
		const string RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

		const string STARTING_DELIMITER = "<%";
		const string ENDING_DELIMITER = "%>";
		/// <summary>
		/// Alpha Numeric with Underscore
		/// </summary>
		const string KEY_FORMAT = "^" + STARTING_DELIMITER + "[A-Za-z0-9_]*" + ENDING_DELIMITER + "$";
		/// <summary>
		/// Numeric
		/// </summary>
		const string POSITIONAL_KEY_FORMAT = "^" + STARTING_DELIMITER + "[0-9]*" + ENDING_DELIMITER + "$";

		PresentationDocument _currentDocument = null;
		PresentationPart _presentationPart = null;
		SlideId _currentSlideId = null;
		SlidePart _currentSlidePart = null;

		int _slideCount = 0;

		/// <summary>
		/// Returns the number of slides in the presentation
		/// </summary>
		public int SlideCount
		{
			get
			{
				return _slideCount;
			}
		}

		public PresentationXMLHelper(string fileName)
		{
			if (!File.Exists(fileName))
				throw new ArgumentException("File " + fileName + " does not exist.");

			// Open file for editing purpose
			_currentDocument = PresentationDocument.Open(fileName, true);
			_presentationPart = _currentDocument.PresentationPart;
			_slideCount = _presentationPart.SlideParts.Count();
		}

		/// <summary>
		/// Set the slide to work on. The other methods will work on this slide.
		/// </summary>
		/// <param name="slideIndex">Zero Based Index of Slide to set as current slide</param>
		public void SetCurrentSlide(int slideIndex)
		{
			_currentSlideId = null;
			_currentSlidePart = null;

			if (slideIndex < 0 || slideIndex >= _slideCount)
				throw new ArgumentOutOfRangeException("slideIndex");

			_currentSlideId = _presentationPart.Presentation.SlideIdList.ChildElements[slideIndex] as SlideId;
			// Get the slide part for the specified slide.
			_currentSlidePart = _presentationPart.GetPartById(_currentSlideId.RelationshipId) as SlidePart;
		}

		/// <summary>
		/// Method to delete a slide
		/// http://msdn.microsoft.com/en-us/library/office/cc850840.aspx
		/// </summary>
		public void DeleteSlide()
		{
			// Remove Slide from Slide List
			_presentationPart.Presentation.SlideIdList.RemoveChild(_currentSlideId);

			// In our normal case this CustomShowList would be empty
			if (_presentationPart.Presentation.CustomShowList != null)
			{
				foreach (var customShow in _presentationPart.Presentation.CustomShowList.Elements<CustomShow>())
				{
					if (customShow.SlideList != null)
					{
						// Declare a link list of slide list entries.
						LinkedList<SlideListEntry> slideListEntries = new LinkedList<SlideListEntry>();
						foreach (SlideListEntry slideListEntry in customShow.SlideList.Elements())
						{
							// Find the slide reference to remove from the custom show.
							if (slideListEntry.Id != null && slideListEntry.Id == _currentSlideId.RelationshipId)
								slideListEntries.AddLast(slideListEntry);
						}

						// Remove all references to the slide from the custom show.
						foreach (SlideListEntry slideListEntry in slideListEntries)
							customShow.SlideList.RemoveChild(slideListEntry);
					}
				}
			}
			_currentDocument.PresentationPart.Presentation.Save();

			// Remove the slide part.
			_presentationPart.DeletePart(_currentSlidePart);
			_currentSlidePart = null;
			_slideCount--;
		}

		/// <summary>
		/// Inserts data into the first table in the current slide
		/// Assumptions: Table Rows are available in the slide already
		/// Initial Row in the Slide Table is considered as Header Row.
		/// The text if any in the cell will be overwritten
		/// </summary>
		/// <param name="tableData">DataTable containing the data to be inserted</param>
		public void InsertDataIntoTable(DataTable tableData)
		{
			Table tbl = _currentSlidePart.Slide.Descendants<Table>().First();
			int tableRowCount = tbl.Descendants<TableRow>().Count();
			if (tableRowCount <= 1)
				throw new Exception("Table does not have any rows to insert data.");
			int tableColumnCount = tbl.Descendants<TableCell>().Count() / tableRowCount;
			int dataRowCount = tableData.Rows.Count;
			int dataColumnCount = tableData.Columns.Count;
			if (tableRowCount <= dataRowCount || tableColumnCount < dataColumnCount)
				throw new Exception("Table does not have enough rows/cells to insert data.");

			for (int dataRowIndex = 0; dataRowIndex < dataRowCount; dataRowIndex++)
			{
				TableRow currentTableRow = tbl.Descendants<A.TableRow>().ElementAt(dataRowIndex + 1);
				for (int dataColumnIndex = 0; dataColumnIndex < dataColumnCount; dataColumnIndex++)
				{
					object val = tableData.Rows[dataRowIndex][dataColumnIndex];
					string dataValue = val != null ? val.ToString() : String.Empty;

					if (String.IsNullOrEmpty(dataValue))
						continue;
					TableCell currentTableCell = currentTableRow.Descendants<TableCell>().ElementAt(dataColumnIndex);
					Paragraph p = currentTableCell.Descendants<Paragraph>().FirstOrDefault();
					if (p == null)
					{
						p = new Paragraph();
						currentTableCell.TextBody.Append(p);
					}
					else
					{
						EndParagraphRunProperties endProp = p.Elements<EndParagraphRunProperties>().FirstOrDefault();
						if (endProp != null)
							endProp.Remove();
					}


					Run r = p.Descendants<Run>().FirstOrDefault();
					if (r == null)
					{
						r = new Run();
						p.Append(r);
					}

					A.Text txt = r.Descendants<A.Text>().FirstOrDefault();
					if (txt == null)
					{
						txt = new A.Text();
						r.Append(txt);
					}

					txt.Text = dataValue;
				}
			}
		}

		/// <summary>
		/// Replaces the keys in the slide.
		/// Keys will be surrounded with <% and %>
		/// No errors are raised when the key is not available in the slide
		/// </summary>
		/// <param name="nameValuePairs">Collection of Key and Values</param>
		public void SubstituteText(NameValueCollection nameValuePairs)
		{
			Regex searchTerm = new Regex(KEY_FORMAT);
			IEnumerable<A.Text> textElements = _currentSlidePart.Slide.Descendants<A.Text>();
			var matchingTextElements = from textElement in textElements
									   let matches = searchTerm.Matches(textElement.Text.Trim())
									   where matches.Count > 0
									   select textElement;

			foreach (A.Text textElement in matchingTextElements)
			{
				string key = GetKey(textElement.Text);
				if (nameValuePairs.AllKeys.Contains(key))
					textElement.Text = textElement.Text.Replace(STARTING_DELIMITER + key + ENDING_DELIMITER, nameValuePairs[key]);
			}
		}

		/// <summary>
		/// Replaces the positional keys in the slide with values provided
		/// e.g. <%0%>, <%1%>, <%2%> and etc
		/// No erros are raised when the positional parameter is not present in the slide
		/// </summary>
		/// <param name="values">Value Collection</param>
		public void SubstituteText(string[] values)
		{
			Regex searchTerm = new Regex(POSITIONAL_KEY_FORMAT);
			IEnumerable<A.Text> textElements = _currentSlidePart.Slide.Descendants<A.Text>();
			var matchingTextElements = from textElement in textElements
									   let matches = searchTerm.Matches(textElement.Text.Trim())
									   where matches.Count > 0
									   select textElement;

			foreach (A.Text textElement in matchingTextElements)
			{
				string key = GetKey(textElement.Text);
				int keyIndex = 0;
				if (Int32.TryParse(key, out keyIndex))
					textElement.Text = textElement.Text.Replace(STARTING_DELIMITER + key + ENDING_DELIMITER, values[keyIndex]);
			}
		}

		private string GetKey(string keyText)
		{
			string key = keyText;
			int startingIndex = key.IndexOf("<%");
			if (startingIndex >= 0)
			{
				startingIndex += 2;
				int endingIndex = key.IndexOf("%>", startingIndex);
				if (endingIndex > startingIndex)
					key = key.Substring(startingIndex, endingIndex - startingIndex);
			}
			return key;
		}

		/// <summary>
		/// Method to replace the chart in powerpoint slide with chart from excel workbook
		/// Opens Excel Workbook using Open XML API's
		/// The Width, Height and position parameters of chart are maintained in common slide data.
		/// </summary>
		/// <param name="workbookFilePath">Excel Full File Path</param>
		/// <param name="worksheetName">Name of the worksheet</param>
		public void ReplaceChart(string workbookFilePath, string worksheetName)
		{
			var currentSlideCharts = _currentSlidePart.GetPartsOfType<ChartPart>();
			if (currentSlideCharts.Count() == 0)
				throw new Exception("Unable to get chart object in current slide.");

			using (SpreadsheetDocument chartSpreadsheet = SpreadsheetDocument.Open(workbookFilePath, false))
			{
				WorkbookPart workbookPart = chartSpreadsheet.WorkbookPart;
				var chartWorkSheets = from workSheet in workbookPart.Workbook.Sheets
										where workSheet.GetAttribute("name", String.Empty).Value.Equals(worksheetName, StringComparison.OrdinalIgnoreCase)
										select workSheet;

				if (chartWorkSheets == null || chartWorkSheets.Count() == 0)
					throw new Exception("Unable to get worksheet " + worksheetName + " in the workbook.");

				SheetML.Sheet chartWorkSheet = chartWorkSheets.ElementAt(0) as SheetML.Sheet;
				string relationshipId = chartWorkSheet.GetAttribute("id", RELATIONSHIP_NAMESPACE).Value;
				WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(relationshipId);
				ChartPart chartPart = (ChartPart)worksheetPart.DrawingsPart.GetPartById(relationshipId);

				if (chartPart == null)
					throw new Exception("Unable to get chart from worksheet");

				currentSlideCharts.ElementAt(0).FeedData(chartPart.GetStream());
			}
		}

		public void Dispose()
		{
			if (_currentDocument != null)
				_currentDocument.Dispose();
		}
	}
}
