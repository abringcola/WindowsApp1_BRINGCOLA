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

namespace WindowsApp1_BRINGCOLA
{
    public partial class Signin : Form
    {
        private readonly ILogger<Signin> _logger;
        private string imagePath;
        private string label1FullText;
        private string label2FullText;
        private int label1Index = 0;
        private int label2Index = 0;
        private object lblName { get; set; }
        public Signin(ILogger<Signin> logger)
        {
            InitializeComponent();
            _logger = logger;
            
              // Set lblName to the provided username
            _logger.LogInformation("Signin form initialized.");
        }
        public void SetImagePath(string path)
        {
            imagePath = path;
        }

        private void btnShowPass_Click(object sender, EventArgs e)
        {
            if (txtPassword.PasswordChar == '*')
            {
                btnHidePass.BringToFront();
                txtPassword.PasswordChar = '\0';

            }
        }

        private void btnHidePass_Click(object sender, EventArgs e)
        {
            if (txtPassword.PasswordChar == '\0')
            {
                btnShowPass.BringToFront();
                txtPassword.PasswordChar = '*';
            }
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void lblpassword_Click(object sender, EventArgs e)
        {
            txtPassword.Focus();
        }

        private void txtPassword_TextChanged(object sender, EventArgs e)
        {

        }

        private void lblusername_Click(object sender, EventArgs e)
        {
            txtUsername.Focus();
        }

        private void btnsignup_Click(object sender, EventArgs e)
        {
           
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void txtUsername_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            //string filePath = @"C:\Users\ACT-STUDENT\Downloads\Book1(1).xlsx"; // Change this to your file path
            //_logger.LogInformation("User clicked Submit.");
            //Workbook book = new Workbook();
            //book.LoadFromFile(@"C:\Users\ACT-STUDENT\Downloads\Book1(1).xlsx");
            //Worksheet sheet = book.Worksheets[0];
            //int row = sheet.Rows.Length + 1;
            //DataTable dt = sheet.ExportDataTable(); // Convert sheet data to DataTable
            //bool log = false;
            //bool userFound = false;

            //foreach (DataRow row1 in dt.Rows)
            //{
            //    if (row1["Username"].ToString() == txtUsername.Text && row1["Password"].ToString() == txtPassword.Text)
            //    {
            //        userFound = true;
            //        int status = Convert.ToInt32(row1["Status"]);

            //        if (status == 1)
            //        {
            //            Form2 activeForm = new Form2(_logger);
                       
            //        }
            //        else
            //        {
            //            Inactive_Students inactiveForm = new Inactive_Students();
                        
            //        }

            //        this.Hide(); // Hide SignIn form after login
            //        break;
            //    }
            //}

            //if (!userFound)
            //{
            //    MessageBox.Show("Invalid username or password!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}

            //for (int i = 2; i <= row; i++)
            //{
            //    if (sheet.Range[i, 10].Value == txtUsername.Text && sheet.Range[i, 11].Value == txtPassword.Text)
            //    {
            //        //Success
            //        log = true;
            //        _logger.LogInformation($"User {txtUsername.Text} successfully logged in.");
            //        break;

            //    }
                
            //}

            //if (log)
            //{
            //    MessageBox.Show("Successful", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    _logger.LogInformation("Navigating to Dashboard.");


            //    Dashboard db = new Dashboard(_logger);
            //    this.Hide();
            //    db.Show();


            //}
            //else
            //{
            //    MessageBox.Show("Invalid Username & Password", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    _logger.LogWarning("Login failed for user: " + txtUsername.Text);
            //}


            //string username = txtUsername.Text;
            //string password = txtPassword.Text;

            //// Check if the user exists
            //var user = Form1.RegisteredUsers.FirstOrDefault(u => u.Username == username && u.Password == password);

            //if (user.Username != null)
            //{
            //    MessageBox.Show("Sign in successful!", "Welcome", MessageBoxButtons.OK, MessageBoxIcon.Information);

            //    // Pass all user data to Form2
            //    Form2 f2 = new Form2();
            //    f2.AddAllUsers(); // Method to populate the DataGridView with all users
            //    f2.Show();
            //    this.Hide(); // Hide the Signin form
            //}
            //else
            //{
            //    MessageBox.Show("Invalid username or password.", "Sign In Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void lblText_Click(object sender, EventArgs e)
        {

        }

        private void txtUsername_Enter(object sender, EventArgs e)
        {
            if (lblusername.Visible)
            {
                lblusername.Visible = false;
            }
        }

        private void txtUsername_Leave(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtUsername.Text))
            {
                lblusername.Visible = true;
            }
        }

        private void txtPassword_Enter(object sender, EventArgs e)
        {
            if (lblpassword.Visible)
            {
                lblpassword.Visible = false;
            }
        }

        private void txtPassword_Leave(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtPassword.Text))
            {
                lblpassword.Visible = true;
            }
        }

        private void guna2SignUp_Click(object sender, EventArgs e)
        {
            Form1 f1 = new Form1(_logger);
            f1.Show();
            this.Hide();
        }

        private void guna2Login_Click(object sender, EventArgs e)
        {
            // Show loading indicator
            guna2Loading.Visible = true;
            timer2.Start();

            Task.Run(() =>
            {
                string filePath = @"C:\Users\ACT-STUDENT\Downloads\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGA\Book1(1).xlsx"; // File path to the Excel file
                _logger.LogInformation("User clicked Submit.");
                Workbook book = new Workbook();
                book.LoadFromFile(filePath);
                Worksheet sheet = book.Worksheets[0];
                DataTable dt = sheet.ExportDataTable(); // Convert sheet data to DataTable
                bool log = false;
                string fullName = "Unknown"; // Default value for full name
                string imagePath = ""; // Variable to hold the image path
                int status = 0; // Default status
                bool userFound = false; // Track if user was found
                string gender = "Unknown"; // Default gender

                // Check if the DataTable has the ImagePath column
                if (!dt.Columns.Contains("ImagePath"))
                {
                    Invoke(new Action(() =>
                    {
                        MessageBox.Show("The ImagePath column does not exist in the Excel file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        guna2Loading.Visible = false; // Hide loading
                    }));
                    return;
                }

                // Check for matching credentials in DataTable
                foreach (DataRow row1 in dt.Rows)
                {
                    if (row1["Username"].ToString() == txtUsername.Text && row1["Password"].ToString() == txtPassword.Text)
                    {
                        userFound = true;
                        status = Convert.ToInt32(row1["Status"]);
                        fullName = row1["Name"].ToString(); // Get the full name
                        gender = row1["Gender"].ToString(); // Get the gender
                        imagePath = row1["ImagePath"].ToString(); // Retrieve the image path

                        // Handle user status
                        if (status == 1) // Check if active
                        {
                            log = true; // Successful login
                        }
                        break; // Exit loop once found
                    }
                }

                if (!userFound)
                {
                    Invoke(new Action(() =>
                    {
                        MessageBox.Show("Invalid username or password!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        _logger.LogWarning("Login failed for user: " + txtUsername.Text);
                        guna2Loading.Visible = false; // Hide loading
                    }));
                    return; // Exit if user not found
                }

                // Handle login result
                if (log)
                {
                    Invoke(new Action(() =>
                    {
                        MessageBox.Show($"Successful! Welcome, {fullName}", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        _logger.LogInformation("Navigating to Dashboard.");

                        // Pass the full name, gender, and image path to the Dashboard
                        Dashboard db = new Dashboard(_logger, fullName, gender, imagePath);
                        db.lblDate.Text = DateTime.Now.ToShortDateString(); // Set the current date
                        db.lblName.Text = fullName; // Set the full name
                        db.guna2CirclePictureBox1.Image = Image.FromFile(imagePath); // Set the image in the dashboard
                        this.Hide();
                        db.Show();
                        guna2Loading.Visible = false; // Hide loading
                    }));
                }
                else
                {
                    Invoke(new Action(() =>
                    {
                        MessageBox.Show("Invalid Username & Password", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        _logger.LogWarning("Login failed for user: " + txtUsername.Text);
                        guna2Loading.Visible = false; // Hide loading
                    }));
                }
            });


            //string username = txtUsername.Text;
            //string password = txtPassword.Text;

            //// Check if the user exists
            //var user = Form1.RegisteredUsers.FirstOrDefault(u => u.Username == username && u.Password == password);

            //if (user.Username != null)
            //{
            //    MessageBox.Show("Sign in successful!", "Welcome", MessageBoxButtons.OK, MessageBoxIcon.Information);

            //    // Pass all user data to Form2
            //    Form2 f2 = new Form2();
            //    f2.AddAllUsers(); // Method to populate the DataGridView with all users
            //    f2.Show();
            //    this.Hide(); // Hide the Signin form
            //}
            //else
            //{
            //    MessageBox.Show("Invalid username or password.", "Sign In Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void Signin_Load(object sender, EventArgs e)
        {
            // Hide the loading gif initially
            guna2Loading.Visible = false;
            label1FullText = label1.Text;
            label2FullText = label2.Text; // make sure you have label2 on your form

            label1.Text = "";
            label2.Text = "";

            label1Index = 0;
            label2Index = 0;

            timer1.Start();

        }

        private void guna2GradientPanel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (label1Index < label1FullText.Length)
            {
                label1.Text += label1FullText[label1Index];
                label1Index++;
            }
            else if (label2Index < label2FullText.Length)
            {
                label2.Text += label2FullText[label2Index];
                label2Index++;
            }
            else
            {
                timer1.Stop(); // Stop the timer when both are done
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {

        }
    }
}
