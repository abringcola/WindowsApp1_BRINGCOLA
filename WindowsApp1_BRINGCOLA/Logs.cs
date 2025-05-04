using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsApp1_BRINGCOLA
{
    public partial class Logs : Form
    {
        private Mylogs myLogs; // Create an instance of Mylogs
        public Logs()
        {
            InitializeComponent();
            myLogs = new Mylogs(); // Initialize Mylogs
            LoadLogs(); // Load logs into DataGridView
        }
        private void LoadLogs()
        {
            DataTable logsData = myLogs.LoadLogs(); // Get logs from Mylogs
            dataGridView3.DataSource = logsData; // Bind data to DataGridView
        }

        private void btnupdate_Click(object sender, EventArgs e)
        {

            // Check if a row is selected in the logs DataGridView
            if (dataGridView3.SelectedRows.Count > 0)
            {
                var selectedRow = dataGridView3.SelectedRows[0];

                // Prompt for new log value
                string newLogValue = Prompt.ShowDialog("Enter new log value:", "Update Log");

                if (!string.IsNullOrEmpty(newLogValue))
                {
                    try
                    {
                        // Update the log value directly in the DataGridView
                        selectedRow.Cells[1].Value = newLogValue; // Assuming the log message is in the second column

                        // Load the Excel workbook to save changes
                        using (Workbook book = new Workbook())
                        {
                            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Downloads\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGA\Book1(1).xlsx");
                            Worksheet sheet = book.Worksheets[1]; // Assume logs are in the second sheet

                            // Clear existing data in the sheet (optional)
                            sheet.Clear(); // Clear existing data if needed, or skip this step

                            // Write updated DataGridView content to Excel
                            for (int i = 0; i < dataGridView3.Rows.Count; i++)
                            {
                                for (int j = 0; j < dataGridView3.Columns.Count; j++)
                                {
                                    // Explicitly casting to string
                                    sheet.Range[i + 2, j + 1].Value = dataGridView3.Rows[i].Cells[j].Value?.ToString(); // Adjust for headers
                                }
                            }

                            // Save the changes to the Excel file
                            book.SaveToFile(@"C:\Users\ACT-STUDENT\Downloads\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGA\Book1(1).xlsx", ExcelVersion.Version2016);
                        }

                        // Refresh the DataGridView
                        LoadLogs(); // Refresh logs display
                        MessageBox.Show("Log updated successfully.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error updating log: {ex.Message}");
                    }
                }
                else
                {
                    MessageBox.Show("Log value cannot be empty.");
                }
            }
            else
            {
                MessageBox.Show("Please select a log to update.");
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            // Check if a row is selected in the logs DataGridView
            if (dataGridView3.SelectedRows.Count > 0)
            {
                var selectedRow = dataGridView3.SelectedRows[0];

                // Confirm deletion
                if (MessageBox.Show("Are you sure you want to delete this log?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    // Get the index of the selected row
                    int rowIndex = selectedRow.Index; // Get the index of the selected row

                    // Load the Excel workbook
                    using (Workbook book = new Workbook())
                    {
                        book.LoadFromFile(@"C:\Users\ACT-STUDENT\Downloads\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGA\Book1(1).xlsx");
                        Worksheet sheet = book.Worksheets[1]; // Assume logs are in the second sheet

                        // Delete the corresponding row in the Excel sheet
                        // Note: dataGridView3 index starts from 0, while Excel row index starts from 1 (and has headers)
                        int excelRowIndex = rowIndex + 2; // Adjusting for the header row (assuming headers are in the first row)
                        sheet.DeleteRow(excelRowIndex); // Delete the row from the sheet

                        // Save the changes to the Excel file
                        book.SaveToFile(@"C:\Users\ACT-STUDENT\Downloads\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGA\Book1(1).xlsx", ExcelVersion.Version2016);
                    }

                    // Remove the row from DataGridView
                    dataGridView3.Rows.Remove(selectedRow);

                    // Show success message
                    MessageBox.Show("Log deleted successfully!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                MessageBox.Show("Please select a log to delete.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close(); // Close the form
        }
    }
    // Helper class for prompting user input
    public static class Prompt
    {
        public static string ShowDialog(string text, string caption)
        {
            Form prompt = new Form()
            {
                Width = 400,
                Height = 200,
                Text = caption,
                StartPosition = FormStartPosition.CenterScreen
            };
            Label textLabel = new Label() { Left = 50, Top = 20, Text = text };
            TextBox textBox = new TextBox() { Left = 50, Top = 50, Width = 300 };
            Button confirmation = new Button() { Text = "Ok", Left = 250, Width = 100, Top = 100, DialogResult = DialogResult.OK };

            confirmation.Click += (sender, e) => { prompt.Close(); };

            prompt.Controls.Add(textLabel);
            prompt.Controls.Add(textBox);
            prompt.Controls.Add(confirmation);
            prompt.AcceptButton = confirmation;

            return prompt.ShowDialog() == DialogResult.OK ? textBox.Text : "";
        }
    }

}
