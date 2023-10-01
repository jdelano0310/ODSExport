using Microsoft.VisualBasic;
using System.Data;
using System.Security;
using System.Text;
using System.Windows.Forms;
using Zaretto.ODS;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace ODSExport
{
    public partial class frmMain : Form
    {
        private System.Data.DataSet? spreadsheetData;
        public frmMain()
        {
            InitializeComponent();
        }

        private void UpdateStatus(string statusMSG)
        {
            lblStatus.Text = statusMSG;
            Application.DoEvents();

        }

        private void btnOpenFileDLG_Click(object sender, EventArgs e)
        {
            if (oFDLG.ShowDialog() == DialogResult.OK)
            {
                txtFileName.Text = oFDLG.FileName;
                Application.DoEvents();

                UpdateStatus("Reading file");
                ReadODSFile(oFDLG.FileName);

                try
                {
                    txtFileName.BackColor = SystemColors.Window;
                }
                catch (Exception ex)
                {
                    txtFileName.BackColor = Color.Red;
                    MessageBox.Show($"Error opening file: {ex.Message}\n\n" +
                    $"Details:\n\n{ex.StackTrace}");
                }
                UpdateStatus("");
            }
        }

        private void ReadODSFile(string selectedFileName)
        {

            var odsReaderWriter = new ODSReaderWriter();
            spreadsheetData = odsReaderWriter.ReadOdsFile(selectedFileName);

            dgvSheet.Rows.Clear();
            drpSheetNames.Items.Clear();

            // get the sheets that are available in this file
            if (spreadsheetData != null)
            {
                drpSheetNames.Items.Add("Select a sheet to export");
                foreach (DataTable table in spreadsheetData.Tables)
                {
                    drpSheetNames.Items.Add(table.TableName);
                }
                drpSheetNames.SelectedIndex = 0;
            }
            else
            {
                txtFileName.BackColor = Color.Red;
                MessageBox.Show("Unable to find any spreadsheet objects in this file.");
            }
        }

        private void drpSheetNames_SelectedIndexChanged(object sender, EventArgs e)
        {
            // after the user selects a sheet - load it into the grid
            if (drpSheetNames.SelectedIndex < 1) { return; }

            UpdateStatus("Loading sheet");

            foreach (DataTable table in spreadsheetData.Tables)
            {
                if (table.TableName == drpSheetNames.Text)
                {
                    if (!chkFirstRowColumnNames.Checked)
                    {
                        // the data doesn't include the column names in the first row
                        // the current table can be added to teh grid using generic 
                        // column headers
                        dgvSheet.Columns.Clear();
                        dgvSheet.Rows.Clear();

                        dgvSheet.DataSource = table;

                    }
                    else
                    {
                        // the user has indicated that the data contains a head row
                        dgvSheet.DataSource = null;
                        dgvSheet.Columns.Clear();
                        dgvSheet.Rows.Clear();

                        var rownum = 1;
                        foreach (var row in table.AsEnumerable())
                        {
                            if (rownum == 1)
                            {
                                // this is the header row, add these as columns
                                for (int i = 0; table.Columns.Count > i; i++)
                                {
                                    DataGridViewColumn newCol = new DataGridViewColumn(); // add a column to the grid
                                    DataGridViewCell cell = new DataGridViewTextBoxCell(); //Specify which type of cell in this column
                                    newCol.CellTemplate = cell;

                                    newCol.HeaderText = row.ItemArray[i].ToString();
                                    newCol.Name = row.ItemArray[i].ToString();
                                    newCol.Visible = true;
                                    newCol.Width = 100;

                                    dgvSheet.Columns.Add(newCol);
                                }
                            }
                            else
                            {
                                // row of data
                                var dgvRow = new DataGridViewRow();
                                var newRowIndex = dgvSheet.Rows.Add();

                                for (int i = 0; table.Columns.Count > i; i++)
                                {
                                    dgvSheet.Rows[newRowIndex].Cells[i].Value = row.ItemArray[i].ToString();
                                }

                            }
                            rownum++;
                        }
                    }
                    UpdateStatus("");
                    return;
                }

            }
        }

        private void chkFirstRowColumnNames_CheckedChanged(object sender, EventArgs e)
        {
            drpSheetNames_SelectedIndexChanged(sender, e);
        }

        private char whichDelimiter()
        {
            // check for which file delimiter should be used when creating the text file
            if (cboDelimiter.SelectedIndex < 1) return (char)32;  // nothing is selected

            char requestedDelimiter = (char)32;
            switch (cboDelimiter.Text)
            {
                case "Comma":
                    requestedDelimiter = Convert.ToChar(",");
                    break;
                case "Tab":
                    requestedDelimiter = (char)9;
                    break;
                case "Semicolon":
                    requestedDelimiter = Convert.ToChar(";");
                    break;
                case "Colon":
                    requestedDelimiter = Convert.ToChar(":");
                    break;
                case "Pipe":
                    requestedDelimiter = Convert.ToChar("|");
                    break;
                case "Forward Slash":
                    requestedDelimiter = Convert.ToChar("/");
                    break;
                case "Back Slash":
                    requestedDelimiter = Convert.ToChar("\\");
                    break;
            }

            return requestedDelimiter;
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            char selectedDelimiter = whichDelimiter();

            if (selectedDelimiter == (char)32)
            {
                // no delimiter was selected, the export process can not continue
                MessageBox.Show("A file delimiter must be selected to export the data");
                return;
            };

            if (dgvSheet.Rows.Count > 0)
            {
                // there are rows displayed in the grid the user has selected the data they wish to be exported
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "TXT (*.txt)|*.txt";
                sfd.FileName = "Output.txt";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    // a file was selected, if it exists try to deleted it  
                    if (File.Exists(sfd.FileName))
                    {
                        try
                        {
                            File.Delete(sfd.FileName);
                        }
                        catch (IOException ex)
                        {
                            MessageBox.Show("It wasn't possible to write the data to the disk." + ex.Message);
                            return;
                        }
                    }

                    try
                    {
                        int columnCount = dgvSheet.Columns.Count;
                        string columnNames = "";
                        string[] outputStr = new string[dgvSheet.Rows.Count + 1];

                        if (chkFirstRowColumnNames.Checked)
                        {
                            // include the column names in the export
                            for (int i = 0; i < columnCount; i++)
                            {
                                columnNames += dgvSheet.Columns[i].HeaderText.ToString() + selectedDelimiter;
                            }
                            outputStr[0] = columnNames.Substring(0, columnNames.Length - 1);
                        }

                        for (int i = 1; (i - 1) < dgvSheet.Rows.Count; i++)
                        {
                            for (int j = 0; j < columnCount; j++)
                            {
                                if (selectedDelimiter == Convert.ToChar(","))
                                {
                                    // If the selected delimiter is a comma, check to see if there is a comma embedded
                                    // in the string that is being exported to ensure that it is not
                                    // used to split the data
                                    if (dgvSheet.Rows[i - 1].Cells[j].Value.ToString().Contains(','))
                                    {
                                        // the value contains a , so it needs to be surrounded by quotes
                                        outputStr[i] += $"\"{dgvSheet.Rows[i - 1].Cells[j].Value.ToString()}\",";
                                    }
                                    else
                                    {
                                        outputStr[i] += dgvSheet.Rows[i - 1].Cells[j].Value.ToString() + selectedDelimiter;
                                    }
                                }
                                else
                                {
                                    outputStr[i] += dgvSheet.Rows[i - 1].Cells[j].Value.ToString() + selectedDelimiter;
                                }

                            }

                            // remove the trailing delimiter from the line
                            outputStr[i] = outputStr[i].Substring(0, outputStr[i].Length - 1);
                        }

                        File.WriteAllLines(sfd.FileName, outputStr, Encoding.UTF8);
                        MessageBox.Show("Data Exported Successfully !!!", "Info");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error :" + ex.Message);
                    }

                }
            }
            else
            {
                MessageBox.Show("No Record To Export !!!", "Info");
            }
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            // show the select delimiter message in the dropdown
            cboDelimiter.SelectedIndex = 0;
        }
    }
}