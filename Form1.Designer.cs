namespace PresentationGenerator
{
	partial class Form1
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.btnBrowseTemplate = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.btnSelectOutputFolder = new System.Windows.Forms.Button();
			this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
			this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.btnGenerateReport = new System.Windows.Forms.Button();
			this.txtOutputFolder = new System.Windows.Forms.TextBox();
			this.txtTemplate = new System.Windows.Forms.TextBox();
			this.linkLabel1 = new System.Windows.Forms.LinkLabel();
			this.label3 = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// btnBrowseTemplate
			// 
			this.btnBrowseTemplate.Location = new System.Drawing.Point(379, 12);
			this.btnBrowseTemplate.Name = "btnBrowseTemplate";
			this.btnBrowseTemplate.Size = new System.Drawing.Size(38, 23);
			this.btnBrowseTemplate.TabIndex = 1;
			this.btnBrowseTemplate.Text = "...";
			this.btnBrowseTemplate.UseVisualStyleBackColor = true;
			this.btnBrowseTemplate.Click += new System.EventHandler(this.btnBrowseTemplate_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(-2, 15);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(51, 13);
			this.label1.TabIndex = 2;
			this.label1.Text = "Template";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(-2, 42);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(71, 13);
			this.label2.TabIndex = 4;
			this.label2.Text = "Output Folder";
			// 
			// btnSelectOutputFolder
			// 
			this.btnSelectOutputFolder.Location = new System.Drawing.Point(379, 36);
			this.btnSelectOutputFolder.Name = "btnSelectOutputFolder";
			this.btnSelectOutputFolder.Size = new System.Drawing.Size(38, 23);
			this.btnSelectOutputFolder.TabIndex = 5;
			this.btnSelectOutputFolder.Text = "...";
			this.btnSelectOutputFolder.UseVisualStyleBackColor = true;
			this.btnSelectOutputFolder.Click += new System.EventHandler(this.btnSelectOutputFolder_Click);
			// 
			// openFileDialog1
			// 
			this.openFileDialog1.FileName = "openFileDialog1";
			this.openFileDialog1.Filter = "Presenation Files|*.pptx";
			// 
			// btnGenerateReport
			// 
			this.btnGenerateReport.Location = new System.Drawing.Point(156, 64);
			this.btnGenerateReport.Name = "btnGenerateReport";
			this.btnGenerateReport.Size = new System.Drawing.Size(111, 23);
			this.btnGenerateReport.TabIndex = 6;
			this.btnGenerateReport.Text = "Generate Report";
			this.btnGenerateReport.UseVisualStyleBackColor = true;
			this.btnGenerateReport.Click += new System.EventHandler(this.btnGenerateReport_Click);
			// 
			// txtOutputFolder
			// 
			this.txtOutputFolder.DataBindings.Add(new System.Windows.Forms.Binding("Text", global::PresentationGenerator.Properties.Settings.Default, "OutputFolderPath", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
			this.txtOutputFolder.Location = new System.Drawing.Point(75, 38);
			this.txtOutputFolder.Name = "txtOutputFolder";
			this.txtOutputFolder.Size = new System.Drawing.Size(298, 20);
			this.txtOutputFolder.TabIndex = 3;
			this.txtOutputFolder.Text = global::PresentationGenerator.Properties.Settings.Default.OutputFolderPath;
			// 
			// txtTemplate
			// 
			this.txtTemplate.DataBindings.Add(new System.Windows.Forms.Binding("Text", global::PresentationGenerator.Properties.Settings.Default, "TemplatePath", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
			this.txtTemplate.Location = new System.Drawing.Point(75, 12);
			this.txtTemplate.Name = "txtTemplate";
			this.txtTemplate.Size = new System.Drawing.Size(298, 20);
			this.txtTemplate.TabIndex = 0;
			this.txtTemplate.Text = global::PresentationGenerator.Properties.Settings.Default.TemplatePath;
			// 
			// linkLabel1
			// 
			this.linkLabel1.AutoSize = true;
			this.linkLabel1.Location = new System.Drawing.Point(129, 103);
			this.linkLabel1.Name = "linkLabel1";
			this.linkLabel1.Size = new System.Drawing.Size(55, 13);
			this.linkLabel1.TabIndex = 7;
			this.linkLabel1.TabStop = true;
			this.linkLabel1.Text = "linkLabel1";
			this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(13, 102);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(110, 13);
			this.label3.TabIndex = 8;
			this.label3.Text = "Report Generated In: ";
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(443, 125);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.linkLabel1);
			this.Controls.Add(this.btnGenerateReport);
			this.Controls.Add(this.btnSelectOutputFolder);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtOutputFolder);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btnBrowseTemplate);
			this.Controls.Add(this.txtTemplate);
			this.Name = "Form1";
			this.Text = "Form1";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.TextBox txtTemplate;
		private System.Windows.Forms.Button btnBrowseTemplate;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox txtOutputFolder;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button btnSelectOutputFolder;
		private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
		private System.Windows.Forms.Button btnGenerateReport;
		private System.Windows.Forms.LinkLabel linkLabel1;
		private System.Windows.Forms.Label label3;
	}
}

