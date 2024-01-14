using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Car_Rental_Management_System_01
{
    public partial class CustomerPage : Form
    {
        private Stack<Customer> customerStack = new Stack<Customer>();

        public class Customer
        {
            public string CustomerId { get; set; }
            public string CustomerName { get; set; }
            public string Address { get; set; }
            public string MobileNo { get; set; }

            public Customer(string customerId, string customerName, string address, string mobileNo)
            {
                CustomerId = customerId;
                CustomerName = customerName;
                Address = address;
                MobileNo = mobileNo;
            }
            public class CustomerManager
            {
                private Stack<Customer> customerStack = new Stack<Customer>();
            }
        }
        public CustomerPage()
        {
            InitializeComponent();
            // Initialize the DataGridView
            dataGridView1.ColumnCount = 4;
            dataGridView1.Columns[0].HeaderText = "Customer ID";
            dataGridView1.Columns[1].HeaderText = "Customer Name";
            dataGridView1.Columns[2].HeaderText = "Address";
            dataGridView1.Columns[3].HeaderText = "Mobile Number";
        }

        //CANCEL BUTTON:-
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
            AdminLoginPage adminLoginPage = new AdminLoginPage();
            adminLoginPage.Close();
        }

        // ADD BUTTON:-
        private void button1_Click(object sender, EventArgs e)
        {
            // Get customer information from textboxes
            string customerId = textBox1.Text;
            string customerName = textBox2.Text;
            string address = textBox3.Text;
            string mobileNo = textBox4.Text;

            // Validate input (you might want to add more validation)
            if (string.IsNullOrEmpty(customerId) || string.IsNullOrEmpty(customerName))
            {
                MessageBox.Show("Customer ID and Name are required.");
                return;
            }

            Customer newCustomer = new Customer(customerId, customerName, address, mobileNo);
            customerStack.Push(newCustomer);
            UpdateDataGridView();
            ClearInputFields();
        }

        private void CustomerPage_Load(object sender, EventArgs e)
        {

        }

        private void UpdateDataGridView()
        { 
            // Clear existing rows
            dataGridView1.Rows.Clear();
            
            // Add customers from stack to DataGridView
            foreach (var customer in customerStack.ToArray())
            {
                AddCustomerToDataGridView(customer);
            }
        }

        private void AddCustomerToDataGridView(Customer customer)
        {
            // Add a new row to DataGridView
            dataGridView1.Rows.Add(
                customer.CustomerId,
                customer.CustomerName,
                customer.Address,
                customer.MobileNo);
        }

        // DELETE BUTTON:-
        private void button5_Click(object sender, EventArgs e)
        {
            // Remove customer from stack
            Customer removedCustomer = RemoveCustomerFromStack();

            // Update DataGridView
            UpdateDataGridView();

            // Display a message about the removed customer
            if (removedCustomer != null)
            {
                MessageBox.Show($"Removed customer from stack: {removedCustomer.CustomerName}");
            }
        }

        private Customer RemoveCustomerFromStack()
        {
            if (customerStack.Count > 0)
            {
                return customerStack.Pop();
            }
            else
            {
                Console.WriteLine("Stack is empty");
                return null;
            }
        }
        private void ClearInputFields()
        {
            // Clear input textboxes
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
        }

        //SEARCH BUTTON:-
        private void button6_Click(object sender, EventArgs e)
        {
            string customerIdSearch = txtCustomerSearch.Text;

            // Iterate through each row in the DataGridView
            for (int rowIndex = 0; rowIndex < dataGridView1.Rows.Count; rowIndex++)
            {
                DataGridViewRow row = dataGridView1.Rows[rowIndex];

                // Check if the customer ID in the current row matches the search term
                if (row.Cells[0].Value != null &&
                    row.Cells[0].Value.ToString() == customerIdSearch)
                {
                    row.Selected = true;
                    MessageBox.Show("Customer Found");
                }
                else
                {
                    row.Selected = false;
                }
            }
        }
        
        // ADDING TO EXCEL
        private void button7_Click(object sender, EventArgs e)
        {
            SortDataGridViewByName();
            ExportToExcel("E:\\PAF KIET\\Rental Management 2\\Customer.xlsx", dataGridView1);
        }
        private void SortDataGridViewByName()
        {
            // Assuming the column index of Customer Name is 0, adjust it accordingly
            int columnIndex = 0;

            int rowCount = dataGridView1.Rows.Count;

            for (int i = 0; i < rowCount - 1; i++)
            {
                for (int j = 0; j < rowCount - i - 1; j++)
                {
                    string name1 = dataGridView1.Rows[j].Cells[columnIndex].Value?.ToString() ?? string.Empty;
                    string name2 = dataGridView1.Rows[j + 1].Cells[columnIndex].Value?.ToString() ?? string.Empty;

                    if (string.Compare(name1, name2) < 0)
                    {
                        // Swap the rows if name1 is less than name2
                        for (int k = 0; k < dataGridView1.Columns.Count; k++)
                        {
                            object temp = dataGridView1.Rows[j].Cells[k].Value;
                            dataGridView1.Rows[j].Cells[k].Value = dataGridView1.Rows[j + 1].Cells[k].Value;
                            dataGridView1.Rows[j + 1].Cells[k].Value = temp;
                        }
                    }
                }
            }
        }
        private void ExportToExcel(string filePath, DataGridView dataGridView1)
        {
            try
            {
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook;

                // Try to open the existing workbook
                try
                {
                    workbook = excelApp.Workbooks.Open(filePath);
                }
                catch
                {
                    // If the workbook doesn't exist, create a new one
                    workbook = excelApp.Workbooks.Add();
                }

                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                // Find the last used row in the worksheet
                int lastRow = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

                // Find the next available row for the new data
                int newRow = lastRow + 1;

                // Check if headers are already present in the worksheet
                if (lastRow == 0 || HeadersMatch(dataGridView1, worksheet, lastRow))
                {
                    // Export headers with formatting only if headers are not present
                    for (int i = 1; i <= dataGridView1.Columns.Count; i++)
                    {
                        worksheet.Cells[newRow, i] = dataGridView1.Columns[i - 1].HeaderText;
                        worksheet.Cells[newRow, i].Font.Bold = true;
                        worksheet.Cells[newRow, i].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                    }

                    newRow++; // Move to the next row for data
                }

                // Export data with formatting
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        string cellValue = dataGridView1.Rows[i].Cells[j].Value?.ToString() ?? string.Empty;

                        // Remove extra spaces from the cell value
                        cellValue = cellValue.Replace(" ", "");

                        worksheet.Cells[newRow, j + 1] = cellValue;

                        // Add formatting to cells (adjust as needed)
                        worksheet.Cells[newRow, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    }
                    newRow++;
                }

                // Auto-fit columns for better appearance
                worksheet.Columns.AutoFit();

                // Save the Excel file
                workbook.SaveAs(filePath);
                workbook.Close();
                excelApp.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                MessageBox.Show("Data exported to Excel successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error exporting data to Excel: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool HeadersMatch(DataGridView dataGridView, Excel.Worksheet worksheet, int lastRow)
        {
            for (int i = 1; i <= dataGridView.Columns.Count; i++)
            {
                var columnHeader = dataGridView.Columns[i - 1].HeaderText;
                var cellValue = worksheet.Cells[lastRow, i].Value?.ToString() ?? string.Empty;

                if (columnHeader != cellValue)
                {
                    return false;
                }
            }
            return true;
        }


        // Read From Excel
        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                // Specify the path to your Excel file
                string filePath = "YOUR PATH";

                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                // Clear existing data in the DataGridView
                dataGridView1.Rows.Clear();

                // Read data from Excel and populate the DataGridView
                int lastRow = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                for (int i = 2; i <= lastRow; i++) // Assuming data starts from row 2 (row 1 contains headers)
                {
                    string[] rowData = new string[dataGridView1.Columns.Count];

                    for (int j = 1; j <= dataGridView1.Columns.Count; j++)
                    {
                        rowData[j - 1] = worksheet.Cells[i, j].Value?.ToString() ?? string.Empty;
                    }

                    dataGridView1.Rows.Add(rowData);
                }

                // Close Excel application
                workbook.Close();
                excelApp.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                MessageBox.Show("Data read from Excel successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error reading data from Excel: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
