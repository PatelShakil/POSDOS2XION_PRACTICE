using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

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
    }
}
