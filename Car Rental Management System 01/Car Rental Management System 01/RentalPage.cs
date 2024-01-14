using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Car_Rental_Management_System_01
{
    public partial class RentalPage : Form
    {
        private RentalGraph rentalGraph;
        public RentalPage()
        {
            InitializeComponent();
            // Initialize the DataGridView
            rentalGraph = new RentalGraph();
            dataGridView1.ColumnCount = 6;
            dataGridView1.Columns[0].HeaderText = "Customer ID";
            dataGridView1.Columns[1].HeaderText = "Car Reg No.";
            dataGridView1.Columns[2].HeaderText = "Customer Name";
            dataGridView1.Columns[3].HeaderText = "Rent Fee";
            dataGridView1.Columns[4].HeaderText = "Date";
            dataGridView1.Columns[5].HeaderText = "Due Date";
        }
        public class RentalGraph
        {
            private Dictionary<string, Vertex> vertices;

            public RentalGraph()
            {
                vertices = new Dictionary<string, Vertex>();
            }

            public void AddCustomer(string customerId, string customerName)
            {
                if (!vertices.ContainsKey(customerId))
                {
                    vertices.Add(customerId, new Vertex(customerName));
                }
            }
            public void AddCar(string carRegNo, string carMakeModel)
            {
                if (!vertices.ContainsKey(carRegNo))
                {
                    vertices.Add(carRegNo, new Vertex(carMakeModel));
                }
            }

            public void RentCar(string customerId, string carRegNo)
            {
                if (vertices.ContainsKey(customerId) && vertices.ContainsKey(carRegNo))
                {
                    vertices[customerId].AddEdge(vertices[carRegNo]);
                }
                else
                {
                    Console.WriteLine("Invalid customer or car ID");
                }
            }
            public void DisplayGraph()
            {
                foreach (var vertex in vertices.Values)
                {
                    Console.WriteLine($"Vertex: {vertex.ID}, Name: {vertex.Name}");
                    Console.Write("Edges: ");
                    foreach (var neighbor in vertex.Neighbors)
                    {
                        Console.Write($"{neighbor.ID} ");
                    }
                    Console.WriteLine();
                    Console.WriteLine("----------");
                }
            }
        }
        public class Vertex
        {
            public string ID { get; }
            public string Name { get; }
            public List<Vertex> Neighbors { get; }

            public Vertex(string name)
            {
                ID = Guid.NewGuid().ToString();
                Name = name;
                Neighbors = new List<Vertex>();
            }

            public void AddEdge(Vertex neighbor)
            {
                Neighbors.Add(neighbor);
            }
        }

        private void Rental_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Validate input (you can add more validation as needed)
            if (string.IsNullOrEmpty(txtCustomerId.Text) || string.IsNullOrEmpty(txtCarRegNo.Text) || string.IsNullOrEmpty(txtCustomerName.Text)
                || string.IsNullOrEmpty(txtRentFee.Text))
            {
                MessageBox.Show("Please fill in all fields before adding a rental.");
                return;
            }

            // Create a new row with the input data
            string[] rowValues = 
            {
                txtCustomerId.Text,
                txtCarRegNo.Text,
                txtCustomerName.Text,
                txtRentFee.Text,
                dateTimePicker1.Text,
                dateTimePicker2.Text,
            };

            rentalGraph.AddCustomer(txtCustomerId.Text, txtCustomerName.Text);
            rentalGraph.AddCar(txtCarRegNo.Text, "Car Make and Model");
            rentalGraph.RentCar(txtCustomerId.Text, txtCarRegNo.Text);

            dataGridView1.Rows.Add(rowValues);

            // Optionally, you can clear the input fields after adding a rental
            ClearInputField();
        }

        private void ClearInputField()
        {
            txtCustomerId.Clear();
            txtCarRegNo.Clear();
            txtCustomerName.Clear();
            txtRentFee.Clear();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // Connecting to Excel
            ExportToExcel("E:\\PAF KIET\\Rental Management 2\\Rental.xlsx", dataGridView1);
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
        private void button4_Click(object sender, EventArgs e)
        {
                try
                {
                    // Specify the path to your Excel file
                    string filePath = "E:\\PAF KIET\\Rental Management 2\\Rental.xlsx";

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
