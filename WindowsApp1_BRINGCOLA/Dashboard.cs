using System;
using Microsoft.Extensions.Logging;
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
    public partial class Dashboard: Form
    {
        private readonly ILogger<Signin> _logger; // Change to ILogger<Signin>
        private Button currentButton;
        private Random random;
        private int tempIndex;
        private Form activeForm;

        public Dashboard(ILogger<Signin> logger, string fullName, string gender, string imagePath)
        {
            InitializeComponent();
            _logger = logger;
            lblName.Text = fullName; // Set the full name in the label
            _logger.LogInformation("Dashboard initialized.");

            // Set the selected image if the path is not empty
            if (!string.IsNullOrEmpty(imagePath))
            {
                guna2CirclePictureBox1.Image = Image.FromFile(imagePath); // Set the selected image
            }
            // Load DbStatus on initialization if needed
            OpenChildForm(new WindowsApp1_BRINGCOLA.DbStatus(_logger), null);
        }

        //private void SetUserGenderImage(string gender)
        //{
        //    if (gender.Equals("Male", StringComparison.OrdinalIgnoreCase))
        //    {
        //        guna2CirclePictureBox1.Image = Properties.Resources.Male_User; // Assuming the resource name is 'Male_User'
        //    }
        //    else if (gender.Equals("Female", StringComparison.OrdinalIgnoreCase))
        //    {
        //        guna2CirclePictureBox1.Image = Properties.Resources.Female_User; // Assuming the resource name is 'Female_User'
        //    }
        //    else
        //    {
        //        guna2CirclePictureBox1.Image = null; // Default image or set a placeholder
        //    }
        //}







        public void OpenChildForm(Form childForm, object btnSender)
        {

            if (activeForm != null)
                activeForm.Close();
            activeForm = childForm;
            childForm.TopLevel = false;
            childForm.FormBorderStyle = FormBorderStyle.None;
            childForm.Dock = DockStyle.Fill;
            this.panelDesktopPane.Controls.Add(childForm);
            this.panelDesktopPane.Tag = childForm;
            childForm.BringToFront();
            childForm.Show();
            lblDashboard.Text = childForm.Text;
        }


        private void Dashboard_Load(object sender, EventArgs e)
        {

        }

        private void llblDashB_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            OpenChildForm(new WindowsApp1_BRINGCOLA.DbStatus(_logger), sender);
        }

        private void guna2btnIAS_Click(object sender, EventArgs e)
        {
            OpenChildForm(new WindowsApp1_BRINGCOLA.Inactive_Students(), sender);
        }

        private void guna2Logs_Click(object sender, EventArgs e)
        {
            OpenChildForm(new WindowsApp1_BRINGCOLA.Logs(), sender);
        }

        private void guna2btnAS_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2(_logger);
            form2.Owner = this; // Set the owner to Dashboard
            OpenChildForm(form2, sender); // Open Form2 in Dashboard
        }

        private void guna2BtnLO_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show("Are you sure you want to Logout?", "Warning", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                Mylogs myLogs = new Mylogs();
                myLogs.insertLogs(lblName.Text, "User logged out."); // Log the logout action

                // Proceed with logout
                Signin si = new Signin(_logger);
                this.Hide();
                si.Show();
                _logger.LogInformation("User logged out: " + lblName.Text);
            }
            else
            {
                return;
            }
        }



        
    }
}
