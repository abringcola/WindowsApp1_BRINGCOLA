using Microsoft.Extensions.Logging;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsApp1_BRINGCOLA
{
   
    public partial class Form1 : Form
    {
        // Declare the RegisteredUsers list here
        public static List<(string Username, string Password, string Name, string Gender, string Hobbies, string Address, string Email, string Age, string Birthday, string Saying, string FavColor, string Course, int status)> RegisteredUsers = new List<(string, string, string, string, string, string, string, string, string, string, string, string, int)>();

        private readonly ILogger<Signin> _logger;

        string[] student = new string[100];
        int i = 1;

        public string Range { get; internal set; }

        public Form1(ILogger<Signin> logger)
        {
            InitializeComponent();
            _logger = logger;

        }
       // Form2 f2 = new Form2(_logger);

        private void txtS_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnDisplay_Click(object sender, EventArgs e)
        {
            string data = "";
            string name = "";
            string gender = "";
            string hobbies = "";
            string birthday = "";
            string saying = "";
            string address = " ";
            string email = "";
            int age;
            data = data  + txtName.Text + "";
            student[1] = txtName.Text;

            if (string.IsNullOrEmpty(txtName.Text))
            {
                name += txtName.Text;
            }
           if(radMale.Checked == true)
            {
                gender += radMale.Text;

            }

            else
            {
                gender += radFemale.Text;
            }
           
            if (chkXB.Checked == true )
            {
                hobbies += chkXB.Text;
            }
            if (chkBB.Checked == true)
            {
                hobbies += chkBB.Text;

            }
            if (chkBMT.Checked == true)
            {
                hobbies += chkBMT.Text;
            }
            if (string.IsNullOrEmpty(txtS.Text))
            {
                saying += txtS.Text;
                
            }
            if (string.IsNullOrEmpty(DTPBirth.Text))
            {
                birthday += DTPBirth.Text;

            }
            if (string.IsNullOrEmpty(txtAddress.Text))
            {
                address += txtAddress.Text;
            }
            if (string.IsNullOrEmpty(txtemail.Text))
            {
                email += txtemail.Text;
            }

            if (int.TryParse(txtage.Text, out age) && age > 0)
            {
               

            }
            else
            {
                MessageBox.Show("Please enter a valid positive age.");
                return; // Exit if age is invalid

            }

            student[i] = data;
          
            
            
            data = txtName.Text + "";
            

            

            MessageBox.Show("Submission completed successfuly!! We appreciate your contribution.", "Submit", MessageBoxButtons.OK,MessageBoxIcon.Information);

            
            Update();
            


          
            
            
        }

        private void cboFC_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void chkBMT_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void chkBB_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void chkXB_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radFemale_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radMale_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void txtName_TextChanged(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string val = "";
            for(int x = 1 ; x < 5; x++)
            {
                val += "[" + x + "] = " + student[x] + "\n";
            }
            MessageBox.Show(val);

            
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
           // f2.Show();
           
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            txtName.Clear();
            txtS.Clear();
            txtAddress.Clear();
            txtemail.Clear();
            txtage.Clear();
            
            radMale.Checked = false;
            radFemale.Checked = false;
            chkXB.Checked = false;
            chkBB.Checked = false;
            chkBMT.Checked = false;

        }

        private void btnCreate_Click(object sender, EventArgs e)
        {


            string imagePath = txtImagePath.Text; // Get the image path
            string username = txtusername.Text;
            string password = txtPassword1.Text;
            string name = txtName.Text;
            string gender = radMale.Checked ? "Male" : "Female";
            string hobbies = "";
            if (chkXB.Checked) hobbies += chkXB.Text + " ";
            if (chkBB.Checked) hobbies += chkBB.Text + " ";
            if (chkBMT.Checked) hobbies += chkBMT.Text + " ";
            if (chkSb.Checked) hobbies += chkSb.Text + " ";

            string address = txtAddress.Text;
            string email = txtemail.Text;
            string age = txtage.Text;
            string birthday = DTPBirth.Text;
            string saying = txtS.Text;
            string favColor = guna2ComboBox2.SelectedItem.ToString(); // Assuming you have a ComboBox for favorite color
            string course = guna2ComboBox1.SelectedItem.ToString(); // Assuming you have a ComboBox for courses

            // int status = (guna2ComboBox3.SelectedItem != null && guna2ComboBox3.SelectedItem.ToString() == "1") ? 1 : 0;

            int status = 1;

            lblMessage.Text = checkEmpty();

            // Check if the username is already taken
            if (RegisteredUsers.Any(user => user.Username == username))
            {
                RegisteredUsers.Add((username, password, name, gender, hobbies, address, email, age, birthday, saying, favColor, course, status));
                // Save the new user to Excel
                SaveUserToExcel(name, gender, hobbies, birthday, address, email, age, saying, username, password, status, favColor, course);
                

                MessageBox.Show("Username already exists. Please choose another one.");
                return;


            }

           

            // Save user information to Excel
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Downloads\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGA\Book1(1).xlsx");
            Worksheet sheet = book.Worksheets[0];

            int row = sheet.Rows.Length + 1; // Get the next empty row
                                             // Populate the row with user data
            sheet.Range[row, 1].Value = (row - 1).ToString(); // Assuming the first column is an ID
            sheet.Range[row, 2].Value = name;                // Name
            sheet.Range[row, 3].Value = gender;              // Gender
            sheet.Range[row, 4].Value = hobbies;             // Hobbies
            sheet.Range[row, 8].Value = birthday;            // Birthday
            sheet.Range[row, 5].Value = address;             // Address
            sheet.Range[row, 6].Value = email;               // Email
            sheet.Range[row, 7].Value = age;                 // Age
            sheet.Range[row, 9].Value = saying;              // Saying
            sheet.Range[row, 10].Value = username;           // Username
            sheet.Range[row, 11].Value = password;           // Password
            sheet.Range[row, 12].Value = status.ToString();  // Status (1 for active, 0 for inactive)
            sheet.Range[row, 13].Value = favColor;           // Favorite Color
            sheet.Range[row, 14].Value = course;             // Courses

            // Save the profile picture to a designated folder
            string destinationPath = @"C:\Users\ACT-STUDENT\Downloads\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGA\ProfilePictures\";
            string newImagePath = System.IO.Path.Combine(destinationPath, username + System.IO.Path.GetExtension(imagePath));

            // Ensure the directory exists
            System.IO.Directory.CreateDirectory(destinationPath);
            System.IO.File.Copy(imagePath, newImagePath, true); // Copy the image to the new location

            // Save the new image path in the Excel file
            sheet.Range[row, 15].Value = newImagePath; // Assuming column 15 is for the image path

            book.SaveToFile(@"C:\Users\ACT-STUDENT\Downloads\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGA\Book1(1).xlsx", ExcelVersion.Version2016);

           
            // Show a message confirming account creation
            MessageBox.Show("Account created successfully! You can now sign in.", "Account Created", MessageBoxButtons.OK, MessageBoxIcon.Information);

            // Navigate to Signin
            Signin signInForm = new Signin(_logger);
            signInForm.SetImagePath(imagePath); // Pass the image path to the Signin form
            signInForm.Show();
            this.Hide(); // Hide current form


        }

        private void SaveUserToExcel(string name, string gender, string hobbies, string birthday, string address, string email, string age, string saying, string username, string password, int status, string favColor, string course)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Downloads\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGA\Book1(1).xlsx");
            Worksheet sheet = book.Worksheets[0];

            int row = sheet.Rows.Length + 1; // Get the next empty row (1-based index)

            // Populate the row with user data
            sheet.Range[row, 1].Value = (row - 1).ToString(); // Assuming the first column is an ID
            sheet.Range[row, 2].Value = name;                // Name
            sheet.Range[row, 3].Value = gender;              // Gender
            sheet.Range[row, 4].Value = hobbies;             // Hobbies
            sheet.Range[row, 8].Value = birthday;            // Birthday
            sheet.Range[row, 5].Value = address;             // Address
            sheet.Range[row, 6].Value = email;               // Email
            sheet.Range[row, 7].Value = age;                 // Age
            sheet.Range[row, 9].Value = saying;              // Saying
            sheet.Range[row, 10].Value = username;           // Username
            sheet.Range[row, 11].Value = password;           // Password
            sheet.Range[row, 12].Value = status.ToString();  // Status (1 for active, 0 for inactive)
            sheet.Range[row, 13].Value = favColor;           // Favorite Color
            sheet.Range[row, 14].Value = course;             // Courses

            book.SaveToFile(@"C:\Users\ACT-STUDENT\Downloads\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGA\Book1(1).xlsx", ExcelVersion.Version2016);
        }

        






        private void DTPBirth_ValueChanged(object sender, EventArgs e)
        {
            String[] User = DTPBirth.Text.Split(',');

            txtage.Text = (2025 - Convert.ToInt32(User[2]) ).ToString();
            

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.gif;*.bmp|All Files|*.*";
                openFileDialog.Title = "Select an Image";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtImagePath.Text = openFileDialog.FileName;
                }
            }
        }

        public string checkEmpty()
        {
            string errors = "";

            foreach(Control c in Controls)
            {
                if(c is TextBox)
                {
                    if(c.Text == "")
                    {
                        errors += c.Name + "";
                    }
                }
            }
            return errors;
        }
    }
}
