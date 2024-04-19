using System;
using System.Data;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using ExcelDataReader;
using SpreadSheet = Bytescout.Spreadsheet;


namespace POSDOS2XION
{
    public partial class PD2X : Form
    {
        private string connectionString = ""; // Your SQL Server connection string
        public PD2X()
        {
            InitializeComponent();
        }

        private void PD2X_Load(object sender, EventArgs e)
        {

        }


        private void FetchDatabases()
        {
            string serverName = textBox1.Text; // SQL Server name
            string userName = textBox2.Text; // SQL Server username
            string password = textBox3.Text; // SQL Server password

            // Build connection string
            connectionString = $"Data Source={serverName};Initial Catalog=master;User ID={userName};Password={password}";

            try
            {
                // Connect to SQL Server
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    // Query to fetch databases
                    string query = "SELECT name FROM sys.databases WHERE name NOT IN ('master', 'tempdb', 'model', 'msdb')";

                    // Execute the query
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            // Clear existing items in combobox
                            comboBox1.Items.Clear();

                            // Read the result set and populate combobox
                            while (reader.Read())
                            {
                                comboBox1.Items.Add(reader["name"].ToString());
                            }
                            if(comboBox1.Items.Count > 0)
                            {
                            comboBox1.SelectedIndex = 0;

                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle any exceptions
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FetchDatabases();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Show open file dialog to select Excel file
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog.Title = "Select an Excel File";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;

                // Read data from Excel file
                ReadExcelFile(filePath);
                DataTable dataTable = new DataTable();

                if (dataTable != null)
                {
                    // Insert data into SQL Server database
                    //InsertDataIntoDatabase(dataTable);
                }
                else
                {
                    MessageBox.Show("Error reading Excel file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        private DataTable ReadExcel(string filePath)
        {
            try
            {
                // Initialize DataTable to hold the Excel data
                DataTable dataTable = new DataTable();

                // Open the Excel file stream
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    // Create ExcelDataReader to read data from the Excel stream
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        // Create a new DataSet to hold the Excel data
                        DataSet dataSet = new DataSet();

                        // Read data from the Excel file into the DataSet
                        while (reader.Read())
                        {
                            Console.Write(reader.ToString());
                        }
                        

                        // Get the first DataTable (assuming there's only one sheet)
                        dataTable = dataSet.Tables[0];

                        // Map Excel columns to database columns based on provided columnMappings
                        /*foreach (var mapping in columnMappings)
                        {
                            // Check if the mapping contains an Excel column header
                            if (dataTable.Columns.Contains(mapping.Key))
                            {
                                // Rename the Excel column to match the database column name
                                dataTable.Columns[mapping.Key].ColumnName = mapping.Value;
                            }
                        }*/
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error reading Excel file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        private void ReadExcelFile(string filePath)
        {
            SpreadSheet.Spreadsheet doc = new SpreadSheet.Spreadsheet();
            doc.LoadFromFile(filePath);
            SpreadSheet.Worksheet ws = doc.Worksheets.ByName("Sheet1");
            var text = "";
            for(int i = 0; i < 3; i++)
            {
                for(int j = 0; j< 3; j++)
                {
                    text += ws.Cell(i, j) + ",";
                }
                text += "\n";

            }
            label8.Text = text;

        }


        private void InsertDataIntoDatabase(DataTable dataTable)
        {
            try
            {
                // Connect to SQL Server
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    // Use SqlBulkCopy to efficiently insert data into database
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection))
                    {
                        bulkCopy.DestinationTableName = "YourTableName"; // Specify your table name
                        bulkCopy.WriteToServer(dataTable);
                    }
                }

                MessageBox.Show("Data inserted into database successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error inserting data into database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    }
}
