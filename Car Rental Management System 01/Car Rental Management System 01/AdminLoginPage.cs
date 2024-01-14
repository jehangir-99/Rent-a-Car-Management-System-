using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Car_Rental_Management_System_01
{
    public partial class AdminLoginPage : Form
    {
        public AdminLoginPage()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string username = textBox1.Text;
            string password = textBox2.Text;

            if (string.IsNullOrEmpty(username) || string.IsNullOrEmpty(password))
            {
                MessageBox.Show("Please enter both USERNAME & PASSWORD!");
            }
            else if (username == "abc" && password == "123")
            {
                MessageBox.Show("LOGIN SUCCESSFUL");
                Form1 form1 = new Form1();
                form1.Close();
                StartPage startPage = new StartPage();
                startPage.Show();
                this.Close();
            }
            else
            {
                MessageBox.Show("Invalid Username or Password...");
            }
        }
    }
}
