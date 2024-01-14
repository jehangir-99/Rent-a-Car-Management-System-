using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using Excel = Microsoft.Office.Interop.Excel;


namespace Car_Rental_Management_System_01
{
    public partial class CarRegistration : Form
    {
        private const int MaxCars = 100; // Maximum number of cars
        private Car[] cars = new Car[MaxCars];
        private int carCount = 0;
        public CarRegistration()
        {
            InitializeComponent();

            // Initialize the DataGridView
            dataGridView1.ColumnCount = 4;
            dataGridView1.Columns[0].HeaderText = "CarRegNo";
            dataGridView1.Columns[1].HeaderText = "Make";
            dataGridView1.Columns[2].HeaderText = "Model";
            dataGridView1.Columns[3].HeaderText = "Available";
        }

        private void CarRegistration_Load(object sender, EventArgs e)
        {

        }
        
        //ADD BUTTON:-
        private void button1_Click(object sender, EventArgs e)
        {
            // Validate input
            if (string.IsNullOrEmpty(txtregno.Text) || string.IsNullOrEmpty(txtmake.Text) || string.IsNullOrEmpty(txtmodel.Text))
            {
                MessageBox.Show("Please fill in all fields before adding a car.");
                return;
            }
            else if (carCount < MaxCars)
            {
                // Retrieve input from textboxes
                string CarRegNo = txtregno.Text;
                string Make = txtmake.Text;
                string Model = txtmodel.Text;
                bool Available = bool.TryParse(txtavl.Text, out var availability) ? availability : false;

                // Check if the registration number is unique
                if (IsRegistrationNumberUnique(CarRegNo))
                {
                    // Create a new car
                    Car newCar = new Car(CarRegNo, Make, Model, Available);

                    // Add the car to the array
                    cars[carCount] = newCar;
                    carCount++;

                    // Display success message
                    MessageBox.Show("Car added successfully!");

                    // Clear input fields
                    ClearInputFields();

                    // Refresh the DataGridView
                    RefreshCarGrid();
                }
                else
                {
                    MessageBox.Show("Registration number already exists. Please enter a unique registration number.");
                }
            }
            else
            {
                MessageBox.Show("Maximum number of cars reached. Cannot add more cars.");
            }
        }

        private bool IsRegistrationNumberUnique(string CarRegNo)
        {
            // Check if the registration number already exists in the array
            for (int i = 0; i < carCount; i++)
            {
                if (cars[i].CarRegNo == CarRegNo)
                {
                    return false;
                }
            }
            return true;
        }
        private void RefreshCarGrid()
        {
            // Clear and update the DataGridView
            dataGridView1.Rows.Clear();
            for (int i = 0; i < carCount; i++)
            {
                dataGridView1.Rows.Add(cars[i].CarRegNo, cars[i].Make, cars[i].Model, cars[i].Available ? "Yes" : "No");
            }
        }
        private void ClearInputFields()
        {
            // Clear input fields
            txtregno.Clear();
            txtmake.Clear();
            txtmodel.Clear();
                    
        }

        // DELETE BUTTON:-
        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                // Get the selected index
                int selectedIndex = dataGridView1.SelectedRows[0].Index;

                // Shift the elements to remove the selected car from the array
                for (int i = selectedIndex; i < carCount - 1; i++)
                {
                    cars[i] = cars[i + 1];
                }

                // Decrement the car count
                carCount--;

                // Refresh the DataGridView
                RefreshDataGridView2();

                // Clear input fields or perform any additional cleanup
                ClearFields();
            }
        }
        private void RefreshDataGridView2()
        {
            // Clear existing rows in DataGridView
            dataGridView1.Rows.Clear();

            // Add cars to DataGridView
            for (int i = 0; i < carCount; i++)
            {
                dataGridView1.Rows.Add(cars[i].CarRegNo, cars[i].Make, cars[i].Model, cars[i].Available ? "Yes" : "No");
            }
        }

            // EDIT BUTTON:-
        private void button5_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                // Get the selected car
                int selectedCar = dataGridView1.SelectedRows[0].Index;

                // Update the selected car with the new values
                cars[selectedCar].CarRegNo = txtregno.Text;
                cars[selectedCar].Make = txtmake.Text;
                cars[selectedCar].Model = txtmodel.Text;
                string selectedAvailability = txtavl.SelectedItem?.ToString();
                if (string.Equals(selectedAvailability, "YES", StringComparison.OrdinalIgnoreCase))
                {
                    cars[selectedCar].Available = true;
                }
                else if (string.Equals(selectedAvailability, "NO", StringComparison.OrdinalIgnoreCase))
                {
                    cars[selectedCar].Available = false;
                }

                // Refresh the DataGridView
                RefreshDataGridView3();
                ClearFields();
            }
        }
        private void RefreshDataGridView3()
        {
            // Clear existing rows in DataGridView
            dataGridView1.Rows.Clear();

            // Add cars to DataGridView
            for (int i = 0; i < carCount; i++)
            {
                dataGridView1.Rows.Add(
                    cars[i].CarRegNo,
                    cars[i].Make,
                    cars[i].Model,
                    cars[i].Available ? "Yes" : "No"
                );
            }
        }

        //CLEAR BUTTON:-
        private void button3_Click(object sender, EventArgs e)
        {
            ClearFields();
        }

        //CANCEL BUTTON:-
        private void button4_Click(object sender, EventArgs e)
        {
            ClearFields();
            this.Close();
            StartPage startPage = new StartPage();
            startPage.Show();
            AdminLoginPage adminLoginPage = new AdminLoginPage();
            adminLoginPage.Close();
        }

        private void RefreshDataGridView()
        {
            // Bind the list of cars to the DataGridView
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = cars;
        }

        private void ClearFields()
        {
            // Clear input fields
            txtregno.Clear();
            txtmake.Clear();
            txtmodel.Clear();
            txtavl.SelectedIndex = -1;
        }
        public class Car
        {
            public string CarRegNo { get; set; }
            public string Make { get; set; }
            public string Model { get; set; }
            public bool Available { get; set; }

            public Car(string registrationNumber, string make, string model, bool available)
            {
                CarRegNo = registrationNumber;
                Make = make;
                Model = model;
                Available = available;
            }
        }

        //SEARCH BUTTON:-
        private void button6_Click(object sender, EventArgs e)
        {
            string CarRegNo = txtCarSearch.Text;

            // Iterate through each row in the DataGridView
            for (int rowIndex = 0; rowIndex < dataGridView1.Rows.Count; rowIndex++)
            {
                DataGridViewRow row = dataGridView1.Rows[rowIndex];

                // Check if the car registration number in the current row matches the search term
                if (row.Cells[0].Value != null &&
                    row.Cells[0].Value.ToString() == CarRegNo)
                {
                    // Highlight the entire row
                    row.Selected = true;
                    MessageBox.Show("Car Found");
                }
                else
                {
                    row.Selected = false;
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            ExportToExcel("E:\\PAF KIET\\Rental Management 2\\Car.xlsx", dataGridView1);
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
                string filePath = "E:\\PAF KIET\\Rental Management 2\\Car.xlsx";

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
