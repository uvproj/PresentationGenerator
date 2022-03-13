// -----------------------------------------------------------------------
// <copyright file="IPresentation.cs" company="CASAHL Technology" />
// -----------------------------------------------------------------------

namespace PresentationGenerator
{
	using System;
	using System.Collections.Generic;
	using System.Collections.Specialized;
	using System.Data;
	using System.Linq;
	using System.Text;
	
	/// <summary>
	/// TODO: Update summary.
	/// </summary>
	public interface IPresentation
	{
		int SlideCount { get; }

		void SetCurrentSlide(int slideIndex);

		void DeleteSlide();

		void InsertDataIntoTable(DataTable tableData);

		void SubstituteText(NameValueCollection nameValuePairs);

		void SubstituteText(string[] values);

		void ReplaceChart(string workbookFilePath, string worksheetName);
	}
}
