namespace ODSExport
{
	partial class frmMain
	{
		/// <summary>
		///  Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		///  Clean up any resources being used.
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
		///  Required method for Designer support - do not modify
		///  the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			label1 = new Label();
			txtFileName = new TextBox();
			btnOpenFileDLG = new Button();
			dgvSheet = new DataGridView();
			drpSheetNames = new ComboBox();
			label2 = new Label();
			oFDLG = new OpenFileDialog();
			lblStatus = new Label();
			chkFirstRowColumnNames = new CheckBox();
			btnExport = new Button();
			((System.ComponentModel.ISupportInitialize)dgvSheet).BeginInit();
			SuspendLayout();
			// 
			// label1
			// 
			label1.AutoSize = true;
			label1.Location = new Point(12, 9);
			label1.Name = "label1";
			label1.Size = new Size(189, 15);
			label1.TabIndex = 0;
			label1.Text = "Select Only Office Spreadsheet File";
			// 
			// txtFileName
			// 
			txtFileName.Location = new Point(12, 27);
			txtFileName.Name = "txtFileName";
			txtFileName.ReadOnly = true;
			txtFileName.Size = new Size(726, 23);
			txtFileName.TabIndex = 1;
			// 
			// btnOpenFileDLG
			// 
			btnOpenFileDLG.Location = new Point(744, 27);
			btnOpenFileDLG.Name = "btnOpenFileDLG";
			btnOpenFileDLG.Size = new Size(44, 23);
			btnOpenFileDLG.TabIndex = 2;
			btnOpenFileDLG.Text = "...";
			btnOpenFileDLG.UseVisualStyleBackColor = true;
			btnOpenFileDLG.Click += btnOpenFileDLG_Click;
			// 
			// dgvSheet
			// 
			dgvSheet.AllowUserToAddRows = false;
			dgvSheet.AllowUserToDeleteRows = false;
			dgvSheet.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
			dgvSheet.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			dgvSheet.Location = new Point(12, 86);
			dgvSheet.MultiSelect = false;
			dgvSheet.Name = "dgvSheet";
			dgvSheet.ReadOnly = true;
			dgvSheet.RowTemplate.Height = 25;
			dgvSheet.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
			dgvSheet.Size = new Size(776, 352);
			dgvSheet.TabIndex = 3;
			// 
			// drpSheetNames
			// 
			drpSheetNames.DropDownStyle = ComboBoxStyle.DropDownList;
			drpSheetNames.FormattingEnabled = true;
			drpSheetNames.Location = new Point(68, 58);
			drpSheetNames.Name = "drpSheetNames";
			drpSheetNames.Size = new Size(202, 23);
			drpSheetNames.TabIndex = 4;
			drpSheetNames.SelectedIndexChanged += drpSheetNames_SelectedIndexChanged;
			// 
			// label2
			// 
			label2.AutoSize = true;
			label2.Location = new Point(12, 61);
			label2.Name = "label2";
			label2.Size = new Size(41, 15);
			label2.TabIndex = 5;
			label2.Text = "Sheets";
			// 
			// oFDLG
			// 
			oFDLG.DefaultExt = "*.ods | Open Document Spreadsheet";
			oFDLG.FileName = "*.ods";
			// 
			// lblStatus
			// 
			lblStatus.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
			lblStatus.ForeColor = Color.Blue;
			lblStatus.Location = new Point(465, 6);
			lblStatus.Name = "lblStatus";
			lblStatus.Size = new Size(323, 18);
			lblStatus.TabIndex = 6;
			// 
			// chkFirstRowColumnNames
			// 
			chkFirstRowColumnNames.AutoSize = true;
			chkFirstRowColumnNames.Location = new Point(281, 62);
			chkFirstRowColumnNames.Name = "chkFirstRowColumnNames";
			chkFirstRowColumnNames.Size = new Size(201, 19);
			chkFirstRowColumnNames.TabIndex = 7;
			chkFirstRowColumnNames.Text = "First row contains column names";
			chkFirstRowColumnNames.UseVisualStyleBackColor = true;
			chkFirstRowColumnNames.CheckedChanged += chkFirstRowColumnNames_CheckedChanged;
			// 
			// btnExport
			// 
			btnExport.Location = new Point(695, 62);
			btnExport.Name = "btnExport";
			btnExport.Size = new Size(93, 23);
			btnExport.TabIndex = 8;
			btnExport.Text = "Export to CVS";
			btnExport.UseVisualStyleBackColor = true;
			btnExport.Click += btnExport_Click;
			// 
			// frmMain
			// 
			AutoScaleDimensions = new SizeF(7F, 15F);
			AutoScaleMode = AutoScaleMode.Font;
			ClientSize = new Size(800, 450);
			Controls.Add(btnExport);
			Controls.Add(chkFirstRowColumnNames);
			Controls.Add(lblStatus);
			Controls.Add(label2);
			Controls.Add(drpSheetNames);
			Controls.Add(dgvSheet);
			Controls.Add(btnOpenFileDLG);
			Controls.Add(txtFileName);
			Controls.Add(label1);
			Name = "frmMain";
			StartPosition = FormStartPosition.CenterScreen;
			Text = "ODS Sheet Exporter";
			((System.ComponentModel.ISupportInitialize)dgvSheet).EndInit();
			ResumeLayout(false);
			PerformLayout();
		}

		#endregion

		private Label label1;
		private TextBox txtFileName;
		private Button btnOpenFileDLG;
		private DataGridView dgvSheet;
		private ComboBox drpSheetNames;
		private Label label2;
		private OpenFileDialog oFDLG;
		private Label lblStatus;
		private CheckBox chkFirstRowColumnNames;
		private Button btnExport;
	}
}