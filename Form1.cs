using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace PresentationGenerator
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
		}

		private void Form1_FormClosing(object sender, FormClosingEventArgs e)
		{
			Properties.Settings.Default.Save();
		}

		private void btnBrowseTemplate_Click(object sender, EventArgs e)
		{
			if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
			{
				txtTemplate.Text = openFileDialog1.FileName;
			}
		}

		private void btnSelectOutputFolder_Click(object sender, EventArgs e)
		{
			if (folderBrowserDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
			{
				txtOutputFolder.Text = folderBrowserDialog1.SelectedPath;
			}
		}

		private void ReportError(string message)
		{
			MessageBox.Show(message, "Error");
		}
		private void btnGenerateReport_Click(object sender, EventArgs e)
		{
			if (!File.Exists(txtTemplate.Text))
			{
				ReportError("Template file does not exist");
				return;
			}

			string reportFileName = String.Format("{0}\\Report_{1}.pptx", txtOutputFolder.Text, DateTime.Now.ToString("yyyy-MM-ddThhmmss"));
			File.Copy(txtTemplate.Text, reportFileName);

			using(PresentationXMLHelper helper = new PresentationXMLHelper(reportFileName))
			{
				helper.SetCurrentSlide(0);
				NameValueCollection nameValues = new NameValueCollection();
				nameValues.Add("INVENTORYDATE", DateTime.Now.ToShortDateString());
				helper.SubstituteText(nameValues);

				helper.SetCurrentSlide(2);
				helper.SubstituteText(new string[] { "2", "2", "2", "6", "0", "0", "20", "493", "9"});

				DataTable scoresTable = new DataTable();
				scoresTable.Columns.Add("Key", Type.GetType("System.String"));
				scoresTable.Columns.Add("Value", Type.GetType("System.String"));
				scoresTable.Rows.Add("1", "23");
				scoresTable.Rows.Add("2", "1");
				scoresTable.Rows.Add("3", "1");
				scoresTable.Rows.Add("4", "0");
				scoresTable.Rows.Add("5", "0");

				helper.SetCurrentSlide(3);
				helper.InsertDataIntoTable(scoresTable);
				helper.ReplaceChart(@"D:\Workbench\ShortThings\PresentationGenerator\PresentationGenerator\Template\Charts.xlsx", "Application Complexity");
				
				DataTable usersTable = new DataTable();
				usersTable.Columns.Add("SlNo", Type.GetType("System.Int32"));
				usersTable.Columns.Add("UserName", Type.GetType("System.String"));
				usersTable.Columns.Add("LastDataModified", Type.GetType("System.DateTime"));
				usersTable.Columns.Add("Created", Type.GetType("System.Int32"));
				usersTable.Columns.Add("Edited", Type.GetType("System.Int32"));

				usersTable.Rows.Add(1, "System Account", DateTime.Now.Subtract(new TimeSpan(2, 0 ,0, 0)), 4612, 4589);
				usersTable.Rows.Add(2, "Glen Baguley", DateTime.Now, 2202, 2152);

				helper.SetCurrentSlide(8);
				helper.InsertDataIntoTable(usersTable);
				helper.SetCurrentSlide(5);
				helper.DeleteSlide();
			}

			linkLabel1.Text = reportFileName;
		}

		private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
		{
			System.Diagnostics.Process.Start(linkLabel1.Text);
		}
	}
}
