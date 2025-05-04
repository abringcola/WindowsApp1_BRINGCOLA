using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Extensions.Logging;
using Spire.Xls;

namespace WindowsApp1_BRINGCOLA
{
    public partial class DbStatus : Form
    {
        private readonly ILogger<Signin> _logger;

        public DbStatus(ILogger<Signin> logger)
        {
            InitializeComponent();
            _logger = logger ;
            LoadDataFromExcel();

        }

        private void LoadDataFromExcel()
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Downloads\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGA\Book1(1).xlsx"); // Replace with your actual file path
            Worksheet sheet = book.Worksheets[0];
            DataTable dt = sheet.ExportDataTable(); // Convert sheet data to DataTable

            HashSet<string> addedStudents = new HashSet<string>(); // Track added student numbers


            int rowCount = dt.Rows.Count;

            int activeStudents = 0, inactiveStudents = 0;
            int maleCount = 0, femaleCount = 0;
            int redCount = 0, blueCount = 0, yellowCount = 0;
            int basketballCount = 0, volleyballCount = 0, badmintonCount = 0, soccerCount = 0;
            int bsitCount = 0, bscsCount = 0, bscpeCount = 0, bsnCount = 0, bstmCount = 0, bshmCount = 0, actCount = 0;

            foreach (DataRow row in dt.Rows)
            {
                // Read values from the DataRow
                string status = row["Status"].ToString();
                string gender = row["Gender"].ToString();
                string color = row["Fav. Color"].ToString();
                string hobby = row["Hobbies"].ToString();
                string course = row["Courses"].ToString();




                if (status == "1") // Active
                {
                    activeStudents++;
                }
                else if (status == "0") // Inactive
                {
                    inactiveStudents++;
                }

                // Increment gender counts
                if (!string.IsNullOrEmpty(gender))
                {
                    if (gender.ToLower() == "male") maleCount++;
                    else if (gender.ToLower() == "female") femaleCount++;
                }

                // Increment favorite color counts
                if (!string.IsNullOrEmpty(color))
                {
                    if (color.ToLower() == "red") redCount++;
                    else if (color.ToLower() == "blue") blueCount++;
                    else if (color.ToLower() == "yellow") yellowCount++;
                }

                // Increment hobby counts
                if (!string.IsNullOrEmpty(hobby))
                {
                    if (hobby.ToLower() == "basketball") basketballCount++;
                    else if (hobby.ToLower() == "volleyball") volleyballCount++;
                    else if (hobby.ToLower() == "badminton") badmintonCount++;
                    else if (hobby.ToLower() == "soccerball") soccerCount++;
                }

                // Increment course counts
                if (!string.IsNullOrEmpty(course))
                {
                    switch (course.ToLower())
                    {
                        case "bsit": bsitCount++; break;
                        case "bscs": bscsCount++; break;
                        case "bscpe": bscpeCount++; break;
                        case "bsn": bsnCount++; break;
                        case "bstm": bstmCount++; break;
                        case "bshm": bshmCount++; break;
                        case "act": actCount++; break;
                    }
                }
            }

            // Update the labels with string values
            lblAS.Text = activeStudents.ToString(); // Convert int to string
            lblIAS.Text = inactiveStudents.ToString(); // Display inactive student count in label2
            label5.Text = maleCount.ToString();
            lblFemale.Text = femaleCount.ToString();
            label13.Text = basketballCount.ToString();
            label14.Text = volleyballCount.ToString();
            label15.Text = badmintonCount.ToString();
            label37.Text = soccerCount.ToString();
            label22.Text = bsitCount.ToString();
            label25.Text = bscsCount.ToString();
            label26.Text = bscpeCount.ToString();
            label31.Text = bsnCount.ToString();
            label28.Text = bstmCount.ToString();
            label29.Text = bshmCount.ToString();
            label30.Text = actCount.ToString();
            label40.Text = redCount.ToString();
            label38.Text = blueCount.ToString();
            label39.Text = yellowCount.ToString();




        }

        private bool CheckForDuplicates(Worksheet sheet, string username)
        {
            for (int rowIndex = 1; rowIndex <= sheet.Rows.Length; rowIndex++)
            {
                if (sheet.Range[rowIndex, 10].Value?.ToString() == username) // Assuming Username is in the 10th column
                {
                    return true; // Duplicate found
                }
            }
            return false; // No duplicates
        }


        public void SaveUserData(DataGridView activeStudentsGrid, DataGridView inactiveStudentsGrid)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Downloads\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGA\Book1(1).xlsx");
            Worksheet sheet = book.Worksheets[0];

            // Find the last row with data to append new users
            int lastRow = FindLastRow(sheet);

            // Save active students
            for (int rowIndex = 0; rowIndex < activeStudentsGrid.Rows.Count; rowIndex++)
            {
                DataGridViewRow row = activeStudentsGrid.Rows[rowIndex];
                string username = row.Cells["Username"].Value?.ToString();
                if (!CheckForDuplicates(sheet, username))
                {
                    for (int colIndex = 0; colIndex < row.Cells.Count; colIndex++)
                    {
                        sheet.Range[lastRow + rowIndex + 1, colIndex + 1].Value = row.Cells[colIndex].Value?.ToString();
                    }
                }
            }

            // Save inactive students
            for (int rowIndex = 0; rowIndex < inactiveStudentsGrid.Rows.Count; rowIndex++)
            {
                DataGridViewRow row = inactiveStudentsGrid.Rows[rowIndex];
                string username = row.Cells["Username"].Value?.ToString();
                if (!CheckForDuplicates(sheet, username))
                {
                    for (int colIndex = 0; colIndex < row.Cells.Count; colIndex++)
                    {
                        sheet.Range[lastRow + activeStudentsGrid.Rows.Count + rowIndex + 1, colIndex + 1].Value = row.Cells[colIndex].Value?.ToString();
                    }
                }
            }

           

            // Save the workbook
            book.SaveToFile(@"C:\Users\ACT-STUDENT\Downloads\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGA\Book1(1).xlsx", ExcelVersion.Version2016);
          
        }

        public void UpdateCounts(int activeChange, int inactiveChange)
        {
            int currentActive = int.Parse(lblAS.Text);
            int currentInactive = int.Parse(lblIAS.Text);

            // Update the counts based on the changes
            lblAS.Text = (currentActive + activeChange).ToString();
            lblIAS.Text = (currentInactive + inactiveChange).ToString();
        }


        private int FindLastRow(Worksheet sheet)
        {
            int lastRow = 0;
            for (int rowIndex = 1; rowIndex <= sheet.Rows.Length; rowIndex++) // Start from row 1
            {
                if (!string.IsNullOrEmpty(sheet.Range[rowIndex, 1].Value.ToString())) // Check the first column for non-empty values
                {
                    lastRow = rowIndex; // Update lastRow if the cell is non-empty
                }
            }
            return lastRow; // Return the last row with data
        }

        private void DbStatus_Load(object sender, EventArgs e)
        {

        }

        private void label32_Click(object sender, EventArgs e)
        {

        }

        private void guna2ShadowPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void lblAS_Click(object sender, EventArgs e)
        {

        }

        private void guna2ShadowPanel5_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label19_Click(object sender, EventArgs e)
        {

        }

        private void label30_Click(object sender, EventArgs e)
        {

        }

        private void label31_Click(object sender, EventArgs e)
        {

        }

        private void label20_Click(object sender, EventArgs e)
        {

        }

        private void guna2CustomGradientPanel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}

