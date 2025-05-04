using Microsoft.Extensions.Logging;
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
using static Guna.UI2.Native.WinApi;

namespace WindowsApp1_BRINGCOLA
{
    public partial class Inactive_Students: Form
    {
        private readonly ILogger<Signin> _logger;
        public Inactive_Students()
        {
            InitializeComponent();
            LoadInactiveStudents();
        }

        public void AddInactiveUser(object[] userDetails)
        {
            // Add the user details to the DataGridView for inactive students
            dataGridView2.Rows.Add(userDetails);
        }

        public void LoadExcelFile()
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Downloads\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGA\Book1(1).xlsx");
            Worksheet sheet = book.Worksheets[0];
            DataTable dt = sheet.ExportDataTable();
            dataGridView2.DataSource = dt;
        }
        public void LoadInactiveStudents()
        {
            string filePath = @"C:\Users\ACT-STUDENT\Downloads\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGA\Book1(1).xlsx"; // Update this with the actual path
            Workbook book = new Workbook();
            book.LoadFromFile(filePath);
            Worksheet sh = book.Worksheets[0];

            DataTable dt = sh.ExportDataTable();
            DataTable inactiveStudents = dt.Clone();

            foreach (DataRow row in dt.Rows)
            {
                if (row["Status"] != DBNull.Value && Convert.ToInt32(row["Status"]) == 0)
                {
                    // Check for duplicates based on Username
                    if (!inactiveStudents.AsEnumerable().Any(r => r.Field<string>("Username") == row.Field<string>("Username")))
                    {
                        inactiveStudents.ImportRow(row);
                    }
                }
            }

            // Set the DataSource of the DataGridView to the DataTable
            dataGridView2.DataSource = inactiveStudents;
        }







        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void btnupdate_Click(object sender, EventArgs e)
        {
            // Perform your update logic here
            // After updating, call SaveUserData from DbStatus
            // Perform your update logic here
            // After updating, call SaveUserData from DbStatus
            DbStatus dbStatus = new DbStatus(_logger);
            dbStatus.SaveUserData(new DataGridView(), dataGridView2); // Pass a new instance for active students and the inactive students grid
        }

        private void btnsearch_Click(object sender, EventArgs e)
        {
            string searchValue = txtName.Text.Trim();
            dataGridView2.ClearSelection();

            try
            {
                bool found = false;

                foreach (DataGridViewRow r in dataGridView2.Rows)
                {
                    if (r.Cells[1].Value != null && r.Cells[1].Value.ToString().IndexOf(searchValue, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        r.Selected = true;
                        found = true;
                        break; // Exit loop on first match
                    }
                }

                if (!found) // If no match was found, inform the user
                {
                    MessageBox.Show("No matching records found.", "Search Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                // Handle any exceptions that may occur
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            // Check if a row is selected
            if (dataGridView2.CurrentRow != null)
            {
                // Confirm deletion
                if (MessageBox.Show("Are you sure you want to make this record active?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    // Get the current row index in DataGridView
                    int rowIndex = dataGridView2.CurrentCell.RowIndex;

                    // Load the Excel workbook
                    using (Workbook book = new Workbook())
                    {
                        book.LoadFromFile(@"C:\Users\ACT-STUDENT\Downloads\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGA\Book1(1).xlsx");
                        Worksheet sheet = book.Worksheets[0];

                        // Get the username of the student to update the status
                        string usernameToUpdate = dataGridView2.Rows[rowIndex].Cells["Username"].Value.ToString();

                        // Find the row index in the Excel sheet based on the username
                        for (int i = 2; i <= sheet.Rows.Length; i++) // Assuming headers are in the first row
                        {
                            if (sheet.Range[i, 10].Value == usernameToUpdate) // Assuming username is in the 10th column
                            {
                                // Update the status to 0 (inactive) in the 12th column
                                sheet.Range[i, 12].Value = 1.ToString(); // Assuming status is now in the 12th column

                                // Get user details to move to inactive students
                                var userDetails = new List<object>();
                                for (int col = 0; col < dataGridView2.Columns.Count; col++)
                                {
                                    userDetails.Add(dataGridView2.Rows[rowIndex].Cells[col].Value); // Use dataGridView1 here
                                }

                                // Remove the row from the DataGridView
                                dataGridView2.Rows.RemoveAt(rowIndex);

                                // Move the user details to the Inactive_Students DataGridView
                                Form2 f2 = Application.OpenForms.OfType<Form2>().FirstOrDefault();
                                if (f2 != null)
                                {
                                    f2.AddActiveUser(userDetails.ToArray());
                                }

                                break; // Exit loop after finding and updating the row
                            }
                        }

                        // Save the changes to the Excel file
                        book.SaveToFile(@"C:\Users\ACT-STUDENT\Downloads\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGA\Book1(1).xlsx", ExcelVersion.Version2016);

                        // Update the DbStatus counts
                        DbStatus dbStatus = new DbStatus(_logger);
                        dbStatus.UpdateCounts(-1, 1); // Decrease active count by 1, increase inactive count by 1

                        MessageBox.Show("Record made active successfully!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        // Refresh the DataGridView to show current active students
                        LoadInactiveStudents();
                    }
                }
            }
            else
            {
                MessageBox.Show("Please select a row to make active.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void txtName_TextChanged(object sender, EventArgs e)
        {

        }

        private void lblSearch_Click(object sender, EventArgs e)
        {
            txtName.Focus();
        }

        private void txtName_Enter(object sender, EventArgs e)
        {
            if (lblSearch.Visible)
            {
                lblSearch.Visible = false;
            }
        }

        private void txtName_Leave(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtName.Text))
            {
                lblSearch.Visible = true;
            }
        }
    }
}
