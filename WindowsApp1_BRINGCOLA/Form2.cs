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
using System.Web.Security;
using System.Windows.Forms;

namespace WindowsApp1_BRINGCOLA
{
    public partial class Form2 : Form
    {

        string gender;
        string hobbies;

        private readonly ILogger<Signin> _logger;

        public void LoadExcelFile()
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Downloads\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGA\Book1(1).xlsx");
            Worksheet sheet = book.Worksheets[0];
            DataTable dt = sheet.ExportDataTable();
            dataGridView1.DataSource = dt;
        }



        public Form2(ILogger<Signin> logger)
        {
            InitializeComponent();
            LoadExcelFile();
            _logger = logger;
            LoadActiveStudents();

        }

        private void LoadActiveStudents()
        {
            using (Workbook book = new Workbook())
            {
                book.LoadFromFile(@"C:\Users\ACT-STUDENT\Downloads\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGA\Book1(1).xlsx");
                Worksheet sh = book.Worksheets[0];

                DataTable dt = sh.ExportDataTable();
                DataTable activeStudents = dt.Clone(); // Create an empty DataTable with the same structure

                foreach (DataRow row in dt.Rows)
                {
                    if (row["Status"] != DBNull.Value && Convert.ToInt32(row["Status"]) == 1) // Check for active status
                    {
                        // Check for duplicates based on Username
                        if (!activeStudents.AsEnumerable().Any(r => r.Field<string>("Username") == row.Field<string>("Username")))
                        {
                            activeStudents.ImportRow(row); // Import the row if it's not a duplicate
                        }
                    }
                }

                // Set the DataSource of the DataGridView to the DataTable
                dataGridView1.DataSource = activeStudents;
            }
        }











        public void AddAllUsers()
        {
            dataGridView1.Rows.Clear(); // Clear existing rows

            foreach (var user in Form1.RegisteredUsers) // Access the registered users
            {
                // Check if the user is already in the DataGridView
                if (!dataGridView1.Rows.Cast<DataGridViewRow>().Any(r => r.Cells["Username"].Value?.ToString() == user.Username))
                {
                    // Add user information to the DataGridView
                    dataGridView1.Rows.Add(user.Name, user.Gender, user.Hobbies, user.Address, user.Email, user.Birthday, user.Age, user.Saying, user.Username);
                }
            }

            MessageBox.Show("All user data loaded successfully!", "Info", MessageBoxButtons.OK);
        }

        public void AddActiveUser(object[] userDetails)
        {
            // Add the user details to the DataGridView for inactive students
            dataGridView1.Rows.Add(userDetails);
        }
      
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
           
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {


            // Check if a row is selected
            if (dataGridView1.CurrentRow != null)
            {
                // Confirm deletion
                if (MessageBox.Show("Are you sure you want to make this record inactive?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    // Get the current row index in DataGridView
                    int rowIndex = dataGridView1.CurrentCell.RowIndex;

                    // Load the Excel workbook
                    using (Workbook book = new Workbook())
                    {
                        book.LoadFromFile(@"C:\Users\ACT-STUDENT\Downloads\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGA\Book1(1).xlsx");
                        Worksheet sheet = book.Worksheets[0];

                        // Get the username of the student to update the status
                        string usernameToUpdate = dataGridView1.Rows[rowIndex].Cells["Username"].Value.ToString();

                        // Find the row index in the Excel sheet based on the username
                        for (int i = 2; i <= sheet.Rows.Length; i++) // Assuming headers are in the first row
                        {
                            if (sheet.Range[i, 10].Value == usernameToUpdate) // Assuming username is in the 10th column
                            {
                                // Update the status to 0 (inactive) in the 12th column
                                sheet.Range[i, 12].Value = 0.ToString(); // Assuming status is now in the 12th column

                                // Get user details to move to inactive students
                                var userDetails = new List<object>();
                                for (int col = 0; col < dataGridView1.Columns.Count; col++)
                                {
                                    userDetails.Add(dataGridView1.Rows[rowIndex].Cells[col].Value); // Use dataGridView1 here
                                }

                                // Remove the row from the DataGridView
                                dataGridView1.Rows.RemoveAt(rowIndex);

                                // Move the user details to the Inactive_Students DataGridView
                                Inactive_Students inactiveStudentsForm = Application.OpenForms.OfType<Inactive_Students>().FirstOrDefault();
                                if (inactiveStudentsForm != null)
                                {
                                    inactiveStudentsForm.AddInactiveUser(userDetails.ToArray());
                                }

                                break; // Exit loop after finding and updating the row
                            }
                        }

                        // Save the changes to the Excel file
                        book.SaveToFile(@"C:\Users\ACT-STUDENT\Downloads\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGA\Book1(1).xlsx", ExcelVersion.Version2016);

                        // Update the DbStatus counts
                        DbStatus dbStatus = new DbStatus(_logger);
                        dbStatus.UpdateCounts(-1, 1); // Decrease active count by 1, increase inactive count by 1

                        MessageBox.Show("Record made inactive successfully!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        // Refresh the DataGridView to show current active students
                        LoadActiveStudents();
                    }
                }
            }
            else
            {
                MessageBox.Show("Please select a row to make inactive.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }


            //foreach (DataGridViewRow item in dataGridView1.SelectedRows)
            //{
            //    dataGridView1.Rows.RemoveAt(item.Index);
            //}

            //if (MessageBox.Show("Are you sure you want to delete?", "Warning", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning) == DialogResult.Yes)
            //{
            //    foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            //    {

            //        //dataGridView1.Rows.Remove(row);

            //    }





            //}
            //sheet.DeleteRows(row);


        }



        private void btnClose_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnupdate_Click(object sender, EventArgs e)
        {
            //Workbook book = new Workbook();
            //book.LoadFromFile(@"D:\Downloads\Book1(1).xlsx");
            //Worksheet sheet = book.Worksheets[0];

            //Form1 f1 =(Form1)Application.OpenForms["Form1"];
            //int r = dataGridView1.CurrentCell.RowIndex;
            //f1.Range = dataGridView1.Rows[r].Cells[1].Value.ToString();
            //f1.Range = dataGridView1.Rows[r].Cells[8].Value.ToString();
            //f1.Range = dataGridView1.Rows[r].Cells[6].Value.ToString();
            //f1.Range = dataGridView1.Rows[r].Cells[4].Value.ToString();
            //f1.Range = dataGridView1.Rows[r].Cells[5].Value.ToString();
            //f1.Range = dataGridView1.Rows[r].Cells[7].Value.ToString();


            //// Get the gender directly from the selected row
            //string selectedGender = dataGridView1.Rows[r].Cells[2].Value.ToString();
            //f1.radMale.Checked = selectedGender == "Male";
            //f1.radFemale.Checked = selectedGender == "Female";

            //// Get hobbies directly from the selected row
            //string selectedhobbies = dataGridView1.Rows[r].Cells[3].Value.ToString();

            //string[]hobbiesArray = selectedhobbies.Split(',');

            //// Clear previous selections
            //f1.chkXB.Checked = false;
            //f1.chkBB.Checked = false;
            //f1.chkBMT.Checked = false;

            //// Set the checked state based on hobbies
            //foreach (string hobby in hobbiesArray)
            //{
            //    switch (hobby.Trim())
            //    {
            //        case "VolleyBall":
            //            f1.chkXB.Checked = true;
            //            break;
            //        case "BasketBall":
            //            f1.chkBB.Checked = true;
            //            break;
            //        case "Badminton":
            //            f1.chkBMT.Checked = true;
            //            break;
            //    }
            //}

            // Perform your update logic here
            // After updating, call SaveUserData from DbStatus

            // Perform your update logic here
            // After updating, call SaveUserData from DbStatus
            DbStatus dbStatus = new DbStatus(_logger);
            dbStatus.SaveUserData(dataGridView1, new DataGridView()); // Pass the active students grid and a new instance for inactive students



        }

        private void dataGridView1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            btnupdate_Click(sender, e); // Reuse the existing update logic







        }

        private void btnsearch_Click(object sender, EventArgs e)
        {
            string searchValue = txtName.Text.Trim();
            dataGridView1.ClearSelection();

            try
            {
                bool found = false;

                foreach (DataGridViewRow r in dataGridView1.Rows)
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

        private void Form2_Load(object sender, EventArgs e)
        {
           
        }

        private void btnSignin_Click(object sender, EventArgs e)
        {
            Signin si = new Signin(_logger);
            si.Show();
            this.Hide();
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void txtName_TextChanged(object sender, EventArgs e)
        {

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

        private void lblSearch_Click(object sender, EventArgs e)
        {
            txtName.Focus();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btnAddNewS_Click(object sender, EventArgs e)
        {
            // Get the owner of this form, which should be the Dashboard
            Dashboard dashboard = (Dashboard)this.Owner;

            if (dashboard != null)
            {
                Form1 form1 = new Form1(_logger); // Create an instance of Form1
                dashboard.OpenChildForm(form1, null); // Open Form1 in the Dashboard
            }
        }
    }
}
