using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;
//DGV Library:https://www.codeproject.com/?cat=1
using DGVPrinterHelper;
//For Excel
using Excel = Microsoft.Office.Interop.Excel;
using Azure.Identity;
//Iron Library : (Codeproject.com, 2024)
using IronXL;
using System.Text.RegularExpressions;
/*Dependants and Namespaces . 

The code imports a number of namespaces needed for different features: 

essential features including UI elements (System.Windows.Forms),
asynchronous operations (System.Threading.Tasks), 
and database access (System.Data.SqlClient). 
Microsoft DGVPrinterHelper is one of the external libraries used to print DataGridView data.office.For Excel export, use Interop.Excel; for more complex Excel functions, use IronXL.
NewtonSoft.Json.JSON answers from Google APIs are parsed using Linq.

*/



namespace fcrud
{
    /*
   Class Initialisation and Definition 
   The primary Windows Form application is represented by the Form1 class.
  It sets up elements such as the DataGridView, the database connection (dbHelper), and the search box's placeholder text.
  Additionally, it includes private information such as the Google API key, which ought to be transferred to a safe place. 
    */
    public partial class Form1 : Form
    {

        private DBHelper dbHelper;
        private DataGridView dataGridView;
        // Google cloud Credentials key.


       


        private string apiKey = "AIzaSyB3PFpDOm2ZK6pui_lojMrcyIOqR8MZB2k";

        public Form1()
        {
            InitializeComponent();




            // Attach the CellClick event handler to the DataGridView
            DataGridView.CellClick += DataGridView_CellContentClick;
            txtStreet.TextChanged += TxtStreet_TextChanged;
            lstSuggestions.DoubleClick += LstSuggestions_DoubleClick;

            // Attach validation events
            txtFirstName.TextChanged += txtFirstName_TextChanged;
            txtLastName.TextChanged += txtLastName_TextChanged;


            // Initialize DBHelper with the connection string
            
            //dbHelper = new DBHelper("Data Source=Niraj; Initial Catalog=Kui; Integrated Security=True");// Database for Niraj punware.
            dbHelper = new DBHelper("Data Source=Shiv_Shakti\\SQLEXPRESS; Initial Catalog=ku; Integrated Security=True");//Database for Karan chawla

            /* Data can changed by the replacing dbhelper.*/



            // combo box.
            LoadProductOptions();
        }



        //=============================================================================================================================//
        // Real-time Validation for First Name
        private void txtFirstName_TextChanged(object sender, EventArgs e)
        {
            string firstName = txtFirstName.Text;

            if (!Regex.IsMatch(firstName, @"^[a-zA-Z\s]*$"))
            {
                MessageBox.Show("First Name cannot contain numbers or special characters.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                // Remove invalid characters
                txtFirstName.Text = new string(firstName.Where(char.IsLetterOrDigit).ToArray());
                txtFirstName.SelectionStart = txtFirstName.Text.Length; // Keep cursor at the end
            }
        }

        // Real-time Validation for Last Name
        private void txtLastName_TextChanged(object sender, EventArgs e)
        {
            string lastName = txtLastName.Text;

            if (!Regex.IsMatch(lastName, @"^[a-zA-Z\s]*$"))
            {
                MessageBox.Show("Last Name cannot contain numbers or special characters.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                // Remove invalid characters
                txtLastName.Text = new string(lastName.Where(char.IsLetterOrDigit).ToArray());
                txtLastName.SelectionStart = txtLastName.Text.Length; // Keep cursor at the end
            }
        }

        // Real-time Validation for Phone Number
        private void txtNumber_TextChanged(object sender, EventArgs e)
        {
            string phoneNumber = txtNumber.Text;

            if (!Regex.IsMatch(phoneNumber, @"^\d*$")) // Only digits allowed
            {
                MessageBox.Show("Phone Number can only contain numeric digits.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                // Remove invalid characters
                txtNumber.Text = new string(phoneNumber.Where(char.IsDigit).ToArray());
                txtNumber.SelectionStart = txtNumber.Text.Length; // Keep cursor at the end
            }
        }


        //=============================================================================================================================//

        // API HANDLING START FROM HERE: Using the Google cloud.

        //=============================================================================================================================//

        // Text Changed in Street TextBox
        /*
         * Integration of the Google Places API 
         * This section makes it possible to use the Google Places and Geocoding APIs to retrieve address recommendations and related country data: 
         * TxtStreet_TextChanged: Retrieves address recommendations when the text in the street field changes. 
         * GetAddressSuggestions: To obtain address recommendations, send an HTTP request to the Google Places API. 
         * UpdateSuggestionsList: Adds the retrieved suggestions to a list box (lstSuggestions). 
         * LstSuggestions_DoubleClick: Picks an address from the list and uses the Geocoding API to retrieve its nation. 
         */
        // Text Changed in Street TextBox
        private async void TxtStreet_TextChanged(object sender, EventArgs e)
        {
            string input = txtStreet.Text;
            if (!string.IsNullOrWhiteSpace(input))
            {
                var suggestions = await GetAddressSuggestions(input);
                UpdateSuggestionsList(suggestions);
            }
        }

        // Fetch Address Suggestions from Google Places API
        private async Task<string[]> GetAddressSuggestions(string input)
        {
            try
            {
                string apiUrl = $"https://maps.googleapis.com/maps/api/place/autocomplete/json?input={Uri.EscapeDataString(input)}&key={apiKey}";

                using (HttpClient client = new HttpClient())
                {
                    HttpResponseMessage response = await client.GetAsync(apiUrl);
                    if (response.IsSuccessStatusCode)
                    {
                        string json = await response.Content.ReadAsStringAsync();
                        JObject data = JObject.Parse(json);

                        var predictions = data["predictions"];
                        return predictions?.Select(p => p["description"].ToString()).ToArray() ?? Array.Empty<string>();
                    }
                    else
                    {
                        MessageBox.Show($"Error fetching suggestions: {response.StatusCode}");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
            return Array.Empty<string>();
        }

        // Update Suggestions ListBox
        private void UpdateSuggestionsList(string[] suggestions)
        {
            lstSuggestions.Items.Clear();
            if (suggestions.Length > 0)
            {
                lstSuggestions.Items.AddRange(suggestions);
                lstSuggestions.Visible = true;
            }
            else
            {
                lstSuggestions.Visible = false;
            }
        }

        // Suggestion Selected from ListBox
        private async void LstSuggestions_DoubleClick(object sender, EventArgs e)
        {
            if (lstSuggestions.SelectedItem != null)
            {
                string selectedAddress = lstSuggestions.SelectedItem.ToString();
                txtStreet.Text = selectedAddress;
                lstSuggestions.Visible = false;

                // Fetch the country for the selected address
                string country = await GetCountryFromAddress(selectedAddress);
                if (!string.IsNullOrWhiteSpace(country))
                {
                    txtCountry.Text = country;
                }
                else
                {
                    MessageBox.Show("Could not determine the country for the selected address.");
                }
            }
        }

        // Fetch Country from Google Geocode API
        private async Task<string> GetCountryFromAddress(string address)
        {
            try
            {
                string encodedAddress = Uri.EscapeDataString(address);
                string apiUrl = $"https://maps.googleapis.com/maps/api/geocode/json?address={encodedAddress}&key={apiKey}";

                using (HttpClient client = new HttpClient())
                {
                    HttpResponseMessage response = await client.GetAsync(apiUrl);
                    if (response.IsSuccessStatusCode)
                    {
                        string json = await response.Content.ReadAsStringAsync();
                        JObject data = JObject.Parse(json);

                        var results = data["results"];
                        if (results != null && results.Count() > 0)
                        {
                            var addressComponents = results[0]["address_components"];
                            foreach (var component in addressComponents)
                            {
                                var types = component["types"];
                                if (types != null && types.Any(type => type.ToString() == "country"))
                                {
                                    return component["long_name"].ToString();
                                }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show($"Error fetching country: {response.StatusCode}");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
            return string.Empty;
        }





        //=============================================================================================================================//

        // API HANDLING End HERE: Using the Google cloud. 

        //=============================================================================================================================//









        // Method to load product options into the ChecklistBox.

        private void LoadProductOptions()
        {
            try
            {
                dbHelper.OpenConnection();

                // Clear existing items in CheckedListBox
                clbProduct.Items.Clear();

                // Query to fetch product names from the database
                SqlCommand cmd = new SqlCommand("SELECT DISTINCT ProductName FROM Product", dbHelper.GetConnection());
                SqlDataReader reader = cmd.ExecuteReader();

                // Add each product name to the CheckedListBox
                while (reader.Read())
                {
                    clbProduct.Items.Add(reader["ProductName"].ToString());
                }

                reader.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading product options: {ex.Message}");
            }
            finally
            {
                dbHelper.CloseConnection();
            }
        }



        //=============================================================================================================================//
        // Creating the class for the display data for the Grid View.
        public void disp_data()
        {

            try
            {
                /*Open the connection
                 dbHelper = new DBHelper("Data Source=Shiv_Shakti\\SQLEXPRESS; Initial Catalog=karan; Integrated Security=True");]
                 open connection also check the database connection open and close.
                 dbHelper.OpenConnection(); */



                // Fetch data from the database
                SqlCommand cmd = new SqlCommand(
                  "SELECT c.ClientID, c.FirstName, c.LastName, c.Email, c.PhoneNumber, c.Postcode, c.Street, c.Country, " +
                  "STRING_AGG(p.ProductName, ', ') AS Products " +//Impt explain
                  "FROM Client c " +
                  "LEFT JOIN ClientProduct cp ON c.ClientID = cp.ClientID " +
                  "LEFT JOIN Product p ON cp.ProductID = p.ProductID " +
                  "GROUP BY c.ClientID, c.FirstName, c.LastName, c.Email, c.PhoneNumber, c.Postcode, c.Street, c.Country " +
                  "ORDER BY c.FirstName ASC",
                  dbHelper.GetConnection());

                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);

                // Bind the data to the DataGridView
                DataGridView.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
            finally
            {
                dbHelper.CloseConnection();
            }


        }

        //=============================================================================================================================//

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MessageBox.Show("AUTO INCREMENT");


        }

        //=============================================================================================================================//
        // Save button.
        /*
        * The "Save" button 

       The method btnSave_Click_1: 
       confirms that all necessary fields have been filled in. 
       To avoid SQL injection, a parameterised SQL query is used to insert a new client record into the database. 
       clears the input fields and refreshes the DataGridView to display the just inserted record. 
        */

        private void btnSave_Click_1(object sender, EventArgs e)
        {

            // Validate First Name input
            if (!System.Text.RegularExpressions.Regex.IsMatch(txtFirstName.Text, @"^[a-zA-Z\s]+$"))
            {
                MessageBox.Show("First Name can only contain letters and spaces.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Validate Last Name input
            if (!System.Text.RegularExpressions.Regex.IsMatch(txtLastName.Text, @"^[a-zA-Z\s]+$"))
            {
                MessageBox.Show("Last Name can only contain letters and spaces.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Validate Phone Number
            if (!Regex.IsMatch(txtNumber.Text, @"^\d+$"))
            {
                MessageBox.Show("Phone Number can only contain numeric digits.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }



            if (string.IsNullOrWhiteSpace(txtFirstName.Text) ||
         string.IsNullOrWhiteSpace(txtLastName.Text) ||
         string.IsNullOrWhiteSpace(txtEmail.Text) ||
         string.IsNullOrWhiteSpace(txtNumber.Text) ||
         string.IsNullOrWhiteSpace(txtPostCode.Text) ||
         string.IsNullOrWhiteSpace(txtStreet.Text) ||
         string.IsNullOrWhiteSpace(txtCountry.Text) ||
         clbProduct.CheckedItems.Count == 0)
            {
                MessageBox.Show("All fields are required, and at least one product must be selected.");
                return;
            }

            try
            {
                dbHelper.OpenConnection();

                // Insert into Client table and get the inserted ClientID
                SqlCommand cmdClient = new SqlCommand(
                    "INSERT INTO Client (FirstName, LastName, Email, PhoneNumber, Postcode, Street, Country) " +
                    "VALUES (@FirstName, @LastName, @Email, @PhoneNumber, @Postcode, @Street, @Country); " +
                    "SELECT SCOPE_IDENTITY();", dbHelper.GetConnection());

                cmdClient.Parameters.AddWithValue("@FirstName", txtFirstName.Text);
                cmdClient.Parameters.AddWithValue("@LastName", txtLastName.Text);
                cmdClient.Parameters.AddWithValue("@Email", txtEmail.Text);
                cmdClient.Parameters.AddWithValue("@PhoneNumber", txtNumber.Text);
                cmdClient.Parameters.AddWithValue("@Postcode", txtPostCode.Text);
                cmdClient.Parameters.AddWithValue("@Street", txtStreet.Text);
                cmdClient.Parameters.AddWithValue("@Country", txtCountry.Text);

                int clientId = Convert.ToInt32(cmdClient.ExecuteScalar()); // Get the inserted ClientID

                // Insert selected products into ClientProduct table
                foreach (var item in clbProduct.CheckedItems)
                {
                    SqlCommand cmdProduct = new SqlCommand(
                        "INSERT INTO ClientProduct (ClientID, ProductID) " +
                        "VALUES (@ClientID, (SELECT ProductID FROM Product WHERE ProductName = @ProductName))",
                        dbHelper.GetConnection());  //Impt explain
                    cmdProduct.Parameters.AddWithValue("@ClientID", clientId);
                    cmdProduct.Parameters.AddWithValue("@ProductName", item.ToString());
                    cmdProduct.ExecuteNonQuery();
                }

                MessageBox.Show("Record inserted successfully!");

                // Clear input fields
                txtClientID.Clear();
                txtFirstName.Clear();
                txtLastName.Clear();
                txtEmail.Clear();
                txtNumber.Clear();
                txtPostCode.Clear();
                txtStreet.Clear();
                txtCountry.Clear();
                foreach (int i in clbProduct.CheckedIndices)
                {
                    clbProduct.SetItemChecked(i, false);
                }

                // Refresh DataGridView
                disp_data();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
            finally
            {
                dbHelper.CloseConnection();
            }
        }


        //=============================================================================================================================//




        //Delete Button.
        //Delete Button.
        /*
         *  Delete Button 
        The btnDelete_Click method: 
        Validates that a ClientID is provided. 
        Prompts the user for confirmation before deleting. 
        Deletes the record from the database if it exists, or notifies the user if no match is found. 
        Refreshes the DataGridView and clears the form fields.
        */
        private void btnDelete_Click(object sender, EventArgs e)
        {
            // Validate that all required fields are filled
            if (string.IsNullOrWhiteSpace(txtClientID.Text))

            {
                MessageBox.Show("Please provide a Client ID to delete a record.");
                return;
            }


            // Show a confirmation dialog before deleting
            DialogResult dialogResult = MessageBox.Show(
                "Are you sure you want to delete this record?",
                "Confirm Delete",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            // Check if the user selected "Yes"
            if (dialogResult == DialogResult.Yes)
            {
                try
                {
                    //open the data connection.
                    dbHelper.OpenConnection();


                    // Prepare the SQL INSERT query (excluding ClientID since it's auto-incremented)
                    SqlCommand cmd = new SqlCommand(
                        "Delete from Client where ClientId = @ClientId", dbHelper.GetConnection());

                    // Add parameters to prevent SQL injection
                    cmd.Parameters.AddWithValue("@ClientId", txtClientID.Text);

                    // Execute the query
                    int rowsAffected = cmd.ExecuteNonQuery();

                    // Check if a record was deleted
                    if (rowsAffected > 0)
                    {
                        MessageBox.Show("Record deleted successfully!");
                    }
                    else
                    {
                        MessageBox.Show("No record found with the given Client ID.");
                    }



                    // Clear the input fields after successful insertion
                    txtClientID.Clear();
                    txtFirstName.Clear();
                    txtLastName.Clear();
                    txtEmail.Clear();
                    txtNumber.Clear();
                    txtPostCode.Clear();
                    txtStreet.Clear();
                    txtCountry.Clear();
                    foreach (int i in clbProduct.CheckedIndices)
                    {
                        clbProduct.SetItemChecked(i, false);
                    }
                    //calling the display button for the grid view.
                    disp_data();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error: {ex.Message}");
                }
                finally
                {
                    dbHelper.CloseConnection();
                }

            }
            else
            {
                // If user selects "No", do nothing
                MessageBox.Show("Delete operation canceled.", "Canceled", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }


        }




        //=============================================================================================================================//
        private void Form1_Load(object sender, EventArgs e)
        {
            disp_data();
        }



        // Update button
        /*
        * Update Button 

      The btnUpdate_Click method: 
      Validates that all fields are filled. 
      Updates the database record corresponding to the provided ClientID. 
      Refreshes the DataGridView and clears the form fields after a successful update. 
       */
        private void btnUpdate_Click(object sender, EventArgs e)
        {
            // Validate that all required fields are filled
            if (string.IsNullOrWhiteSpace(txtFirstName.Text) ||
        string.IsNullOrWhiteSpace(txtLastName.Text) ||
        string.IsNullOrWhiteSpace(txtEmail.Text) ||
        string.IsNullOrWhiteSpace(txtNumber.Text) ||
        string.IsNullOrWhiteSpace(txtPostCode.Text) ||
        string.IsNullOrWhiteSpace(txtStreet.Text) ||
        string.IsNullOrWhiteSpace(txtCountry.Text) ||
        clbProduct.CheckedItems.Count == 0)
            {
                MessageBox.Show("All fields are required.");
                return;
            }

            try
            {
                dbHelper.OpenConnection();

                // Update Client table
                SqlCommand cmd = new SqlCommand(
                    "UPDATE Client " +
                    "SET FirstName = @FirstName, LastName = @LastName, Email = @Email, " +
                    "PhoneNumber = @PhoneNumber, Postcode = @Postcode, Street = @Street, " +
                    "Country = @Country " +
                    "WHERE ClientID = @ClientID", dbHelper.GetConnection());

                cmd.Parameters.AddWithValue("@ClientID", txtClientID.Text);
                cmd.Parameters.AddWithValue("@FirstName", txtFirstName.Text);
                cmd.Parameters.AddWithValue("@LastName", txtLastName.Text);
                cmd.Parameters.AddWithValue("@Email", txtEmail.Text);
                cmd.Parameters.AddWithValue("@PhoneNumber", txtNumber.Text);
                cmd.Parameters.AddWithValue("@Postcode", txtPostCode.Text);
                cmd.Parameters.AddWithValue("@Street", txtStreet.Text);
                cmd.Parameters.AddWithValue("@Country", txtCountry.Text);
                cmd.ExecuteNonQuery();

                // Delete existing products for the ClientID
                SqlCommand deleteCmd = new SqlCommand(
                    "DELETE FROM ClientProduct WHERE ClientID = @ClientID", dbHelper.GetConnection());
                deleteCmd.Parameters.AddWithValue("@ClientID", txtClientID.Text);
                deleteCmd.ExecuteNonQuery();

                // Insert new products for the ClientID
                foreach (var item in clbProduct.CheckedItems)
                {
                    SqlCommand insertCmd = new SqlCommand(
                        "INSERT INTO ClientProduct (ClientID, ProductID) " +
                        "VALUES (@ClientID, (SELECT ProductID FROM Product WHERE ProductName = @ProductName))",
                        dbHelper.GetConnection());
                    insertCmd.Parameters.AddWithValue("@ClientID", txtClientID.Text);
                    insertCmd.Parameters.AddWithValue("@ProductName", item.ToString());
                    insertCmd.ExecuteNonQuery();
                }

                MessageBox.Show("Record updated successfully!");
                disp_data(); // Refresh the grid view
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
            finally
            {
                dbHelper.CloseConnection();
            }





        }


        //=============================================================================================================================//



        // Data grid view [That code help to select the data from the grid view.]
        /*The Constructor

       constructor sets up events for user interactions and initialises form components.
       It links the text change events for the search and street fields with the DataGridView click events.
       Additionally, it inserts product selections into the dropdown (ComboBox) and initialises the database connection string.
       Product Loading into ComboBox 

       Different product names are retrieved from the database and added to the ComboBox using the LoadProductOptions function.
       This enables users to manage client records by choosing from predefined product alternatives.
       */
       private void DataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
       {
           try
           {
               if (e.RowIndex >= 0)
               {
                   DataGridViewRow row = DataGridView.Rows[e.RowIndex];

                   // Populate text fields, Helping to selelcting the previous record
                   txtClientID.Text = row.Cells["ClientID"].Value.ToString();
                   txtFirstName.Text = row.Cells["FirstName"].Value.ToString();
                   txtLastName.Text = row.Cells["LastName"].Value.ToString();
                   txtEmail.Text = row.Cells["Email"].Value.ToString();
                   txtNumber.Text = row.Cells["PhoneNumber"].Value.ToString();
                   txtPostCode.Text = row.Cells["Postcode"].Value.ToString();
                   txtStreet.Text = row.Cells["Street"].Value.ToString();
                   txtCountry.Text = row.Cells["Country"].Value.ToString();

                   // Clear CheckedListBox
                   foreach (int i in clbProduct.CheckedIndices)
                   {
                       clbProduct.SetItemChecked(i, false);
                   }

                    // Check products in CheckedListBox
                    // Check products in CheckedListBox
                    //All client records are retrieved from the database and shown in the DataGridView using the disp_data method.
                    //This acts as the main screen where client data is managed.
                    string[] products = row.Cells["Products"].Value.ToString().Split(new[] { ", " }, StringSplitOptions.None);
                   foreach (string product in products)
                   {
                       int index = clbProduct.Items.IndexOf(product);
                       if (index >= 0)
                       {
                           clbProduct.SetItemChecked(index, true);
                       }
                   }
               }
           }
           catch (Exception ex)
           {
               MessageBox.Show($"Error: {ex.Message}");
           }
       }




        //=============================================================================================================================//

        //Clear Button
        /*
          * verifies the search query. 
          a LIKE operator to search the database for entries that match in several fields (ClientID, FirstName, LastName, etc.). 
          Notifies the user if no results are discovered, or displays matched results in the DataGridView. 
         The search box now has placeholder capability thanks to the InitializePlaceholder method.
         It enhances usability by altering the text colour according on whether the user is typing or leaving the box unfilled. 
          */
        private void btnClear_Click(object sender, EventArgs e)
       {
           // Clear the input fields after successful insertion
           txtClientID.Clear();
           txtFirstName.Clear();
           txtLastName.Clear();
           txtEmail.Clear();
           txtNumber.Clear();
           txtPostCode.Clear();
           txtStreet.Clear();
           txtCountry.Clear();

           foreach (int i in clbProduct.CheckedIndices)
           {
               clbProduct.SetItemChecked(i, false);
           }

       }



       //=============================================================================================================================//

       //Search 

       private void btnSearch_Click(object sender, EventArgs e)
       {

           try
           {
               if (string.IsNullOrWhiteSpace(txtSearch.Text))
               {
                   MessageBox.Show("Please enter a search term.");
                   return;
               }

               string searchTerm = txtSearch.Text.Trim();

               // Open database connection
               dbHelper.OpenConnection();

               // Adjusted SQL query to include JOIN for product information
               SqlCommand cmd = new SqlCommand(
                   "SELECT c.ClientID, c.FirstName, c.LastName, c.Email, c.PhoneNumber, c.Postcode, c.Street, c.Country, " +
                   "STRING_AGG(p.ProductName, ', ') AS Products " +
                   "FROM Client c " +
                   "LEFT JOIN ClientProduct cp ON c.ClientID = cp.ClientID " +
                   "LEFT JOIN Product p ON cp.ProductID = p.ProductID " +
                   "WHERE CAST(c.ClientID AS NVARCHAR) LIKE @Search " +
                   "OR c.FirstName LIKE @Search " +
                   "OR c.LastName LIKE @Search " +
                   "OR c.Email LIKE @Search " +
                   "OR c.PhoneNumber LIKE @Search " +
                   "OR c.Postcode LIKE @Search " +
                   "OR c.Street LIKE @Search " +
                   "OR c.Country LIKE @Search " +
                   "OR p.ProductName LIKE @Search " +
                   "GROUP BY c.ClientID, c.FirstName, c.LastName, c.Email, c.PhoneNumber, c.Postcode, c.Street, c.Country",
                   dbHelper.GetConnection()
               );

               cmd.Parameters.AddWithValue("@Search", "%" + searchTerm + "%");

               // Execute query and bind to DataGridView
               SqlDataAdapter da = new SqlDataAdapter(cmd);
               DataTable dt = new DataTable();
               da.Fill(dt);

               DataGridView.DataSource = dt;

               // Update the total records label
               lblTotal.Text = $"Total Records: {dt.Rows.Count}";

               if (dt.Rows.Count == 0)
               {
                   MessageBox.Show("No records found.");
               }
               else
               {
                   MessageBox.Show($"{dt.Rows.Count} record(s) found.");
               }
           }
           catch (Exception ex)
           {
               MessageBox.Show($"Error: {ex.Message}");
           }
           finally
           {
               dbHelper.CloseConnection();
           }



       }
        // Reset Button.
        /*
         * All input fields are cleared by the btnReset_Click method, which also reloads the entire dataset into the DataGridView.
         * The label for the total records is likewise reset. 
         */
        private void btnReset_Click(object sender, EventArgs e)
       {
           txtClientID.Clear();
           txtFirstName.Clear();
           txtLastName.Clear();
           txtEmail.Clear();
           txtNumber.Clear();
           txtPostCode.Clear();
           txtStreet.Clear();
           txtCountry.Clear();
           foreach (int i in clbProduct.CheckedIndices)
           {
               clbProduct.SetItemChecked(i, false);
           }
           txtSearch.Clear();
           //Total label box empty after reset button.
           lblTotal.Text = string.Empty;

           disp_data();
       }



        //=============================================================================================================================//

        //Print and Preview.
        /*
         * btnPrint_Click: Prints the contents of the DataGridView using the DGVPrinterHelper library. 
           btnPreview_Click: Uses DGVPrinterHelper to show a print preview of the DataGridView data. 
           Users can create and examine physical reports of client data with the use of these functionalities. 
         */
        private void btnPrint_Click(object sender, EventArgs e)
       {
           try
           {


               foreach (DataGridViewColumn column in DataGridView.Columns)
               {
                   column.Visible = true; // Make each column visible
               }
               DGVPrinter printer = new DGVPrinter
               {
                   Title = "Client Data Report", // Report title
                   SubTitle = $"Printed on {DateTime.Now:f}", //  date and time
                   SubTitleFormatFlags = StringFormatFlags.LineLimit | StringFormatFlags.NoClip,
                   PageNumbers = true,
                   PageNumberInHeader = false,
                   PorportionalColumns = true,
                   HeaderCellAlignment = StringAlignment.Near,
                   Footer = "Your Company Name Here", // Footer text
                   FooterSpacing = 15
               };

               // Print the DataGridView
               printer.PrintDataGridView(DataGridView);
           }
           catch (Exception ex)
           {
               MessageBox.Show($"Error during printing: {ex.Message}");
           }
       }

       private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
       {








       }
        /*
         * After pressing the print button you will got option to choose landscaper or potrait, you need to choose landscape so you can print.
         */

        private void btnPreview_Click(object sender, EventArgs e)
       {

           try
           {



               // Initialize DGVPrinter
               DGVPrinter printer = new DGVPrinter
               {

                   Title = "Client Data Report", // Report title
                   SubTitle = $"Printed on {DateTime.Now:f}", //  date and time
                   SubTitleFormatFlags = StringFormatFlags.LineLimit | StringFormatFlags.NoClip,
                   PageNumbers = true,
                   PageNumberInHeader = false,
                   PorportionalColumns = true,
                   HeaderCellAlignment = StringAlignment.Near,
                   Footer = "Your Company Name Here", // Footer text
                   FooterSpacing = 15
               };

               // Use DGVPrinter's built-in preview
               printer.PrintPreviewDataGridView(DataGridView);
           }
           catch (Exception ex)
           {
               MessageBox.Show($"Error during print preview: {ex.Message}");
           }






       }



        /* Print code Reference:
        //Codeproject.com. (2024). CodeProject. [online] Available at: https://www.codeproject.com/?cat=1 [Accessed 5 Dec. 2024].
        //‌

        //From<https://www.mybib.com/tools/harvard-referencing-generator> 

        //=============================================================================================================================//

        //Import to Excel*/
        
        /*
         * The method btnToExcel_Click: 
        uses the IronXL library to export DataGridView data to an Excel file. 
        opens the file using the built-in Excel application after saving it locally. 
        My database ruin of the Ironxl trial but you can connect on your DB and see the result.
         */
        private void btnToExcel_Click(object sender, EventArgs e)
        {

            try
            {
                if (DataGridView.Rows.Count == 0)
                {
                    MessageBox.Show("No data available to export.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                WorkBook workbook = new WorkBook();
                WorkSheet worksheet = workbook.CreateWorkSheet("ExportedData");


                for (int i = 0; i < DataGridView.Rows.Count; i++)
                {
                    for (int j = 0; j < DataGridView.Columns.Count; j++)
                    {
                        string cellAddress = ConvertToCellAddress(i, j);
                        worksheet[cellAddress].Value = DataGridView.Rows[i].Cells[j].Value != null
                            ? DataGridView.Rows[i].Cells[j].Value.ToString()
                            : string.Empty;
                    }
                }

                // Define the file path
                string filePath = "DataGridViewExport.xlsx";

                // Save the workbook
                workbook.SaveAs(filePath);

                // Notify the user
                MessageBox.Show("Data exported successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Open the exported Excel file
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true // Ensure the file opens with the default associated application
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show("An exception occurred: " + ex.Message);
            }
        }

        // Helper method to convert row and column indices to Excel cell addresses
        /*
         * DataGridView row and column indices are converted into Excel cell addresses
         * using the ConvertToCellAddress function. Data in the output Excel file is appropriately structured using this. 
        */

        private string ConvertToCellAddress(int row, int column)
        {
            // Columns in Excel are labeled as A, B, C, ..., Z, AA, AB, ..., etc.
            // The following code converts a column index to this format.
            string columnLabel = "";
            while (column >= 0)
            {
                columnLabel = (char)('A' + column % 26) + columnLabel;
                column = column / 26 - 1;
            }
            // Rows in Excel are labeled as 1, 2, 3, ..., n
            // Adding 1 because Excel is 1-based and our loop is 0-based.
            string rowLabel = (row + 1).ToString();
            return columnLabel + rowLabel;
        }





        // Creating the export Cs file.


        private void ExportDataGridViewToCSFile(string filePath)
        {
            try
            {
                // Validate if the DataGridView contains rows
                if (DataGridView.Rows.Count == 0)
                {
                    MessageBox.Show("No data available to export.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Start building the C# class string
                StringBuilder csFileContent = new StringBuilder();
                csFileContent.AppendLine("using System;");
                csFileContent.AppendLine("using System.Collections.Generic;");
                csFileContent.AppendLine();
                csFileContent.AppendLine("// Auto-generated class based on DataGridView content");
                csFileContent.AppendLine("namespace ExportedData");
                csFileContent.AppendLine("{");
                csFileContent.AppendLine("    public class ExportedDataRow");
                csFileContent.AppendLine("    {");

                // Add properties to the class based on the DataGridView columns
                foreach (DataGridViewColumn column in DataGridView.Columns)
                {
                    string propertyName = column.Name.Replace(" ", ""); // Remove spaces for valid property names
                    string propertyType = "string"; // Assume all fields as strings; customize as needed
                    csFileContent.AppendLine($"        public {propertyType} {propertyName} {{ get; set; }}");
                }

                csFileContent.AppendLine("    }");
                csFileContent.AppendLine();
                csFileContent.AppendLine("    public class ExportedDataSet");
                csFileContent.AppendLine("    {");
                csFileContent.AppendLine("        public List<ExportedDataRow> Rows { get; set; } = new List<ExportedDataRow>();");
                csFileContent.AppendLine("    }");
                csFileContent.AppendLine("}");

                // Add the data rows to the file as an example
                csFileContent.AppendLine();
                csFileContent.AppendLine("// Example usage:");
                csFileContent.AppendLine("namespace ExportedData");
                csFileContent.AppendLine("{");
                csFileContent.AppendLine("    public static class DataSetExample");
                csFileContent.AppendLine("    {");
                csFileContent.AppendLine("        public static ExportedDataSet GetData()");
                csFileContent.AppendLine("        {");
                csFileContent.AppendLine("            var dataSet = new ExportedDataSet();");

                foreach (DataGridViewRow row in DataGridView.Rows)
                {
                    if (!row.IsNewRow)
                    {
                        csFileContent.AppendLine("            dataSet.Rows.Add(new ExportedDataRow");
                        csFileContent.AppendLine("            {");
                        foreach (DataGridViewCell cell in row.Cells)
                        {
                            string propertyName = DataGridView.Columns[cell.ColumnIndex].Name.Replace(" ", "");
                            string value = cell.Value?.ToString() ?? "null";
                            csFileContent.AppendLine($"                {propertyName} = \"{value}\",");
                        }
                        csFileContent.AppendLine("            });");
                    }
                }

                csFileContent.AppendLine("            return dataSet;");
                csFileContent.AppendLine("        }");
                csFileContent.AppendLine("    }");
                csFileContent.AppendLine("}");

                // Save the content to the file
                System.IO.File.WriteAllText(filePath, csFileContent.ToString());
                MessageBox.Show($"C# file successfully created at {filePath}!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting data to C# file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }






        private void btnExport_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "C# Files (*.cs)|*.cs";
                saveFileDialog.Title = "Save as C# File";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    ExportDataGridViewToCSFile(saveFileDialog.FileName);
                }
            }
        }


        //        // Excel code Reference:
        //        //Codeproject.com. (2024). CodeProject. [online] Available at:Ironsoftware.com. (2024). How to Export Datagridview To Excel in C#. [online] Available at: https://ironsoftware.com/csharp/excel/blog/using-ironxl/export-datagridview-to-excel-csharp/ [Accessed 5 Dec. 2024].

        //‌
        //        //‌

        //        //From<https://www.mybib.com/tools/harvard-referencing-generator> 


    }















}

