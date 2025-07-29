using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;

namespace fcrud
{
    internal class DBHelper
    {
        private SqlConnection con;

        // Constructor to initialize the connection
        public DBHelper(string connectionString)
        {
            con = new SqlConnection(connectionString);
        }

        // Method to open the connection safely
        public void OpenConnection()
        {
            try
            {
                if (con.State == ConnectionState.Open)
                {
                    con.Close(); // Close if already open
                }
                con.Open();
            }
            catch (Exception ex)
            {
                throw new Exception($"Error while opening the connection: {ex.Message}");
            }
        }

        // Method to close the connection safely
        public void CloseConnection()
        {
            try
            {
                if (con.State == ConnectionState.Open)
                {
                    con.Close();
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error while closing the connection: {ex.Message}");
            }
        }

        // Method to retrieve the active connection
        public SqlConnection GetConnection()
        {
            return con;
        }


    }
}
