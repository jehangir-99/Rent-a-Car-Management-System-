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
    public partial class StartPage : Form
    {
        public StartPage()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
            CarRegistration carRegistration = new CarRegistration();
            carRegistration.Show(); 
        }

        //logout button
        private void button5_Click(object sender, EventArgs e)
        {
            this.Close();
            AdminLoginPage adminLoginPage = new AdminLoginPage();   
            adminLoginPage.Show();  
        }

        //customer button
        private void button2_Click(object sender, EventArgs e)
        {
            CustomerPage customerPage = new CustomerPage();
            customerPage.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            RentalPage rental = new RentalPage();
            rental.Show();  
        }
    }
}
