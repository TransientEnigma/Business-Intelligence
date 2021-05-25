using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
//test commit with this comment
namespace BI_coursework
{
    public partial class Form1 : Form
    {

        //Leave this alone
        public Form1()
        {
            InitializeComponent();
        }

        //Leave this alone
        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void btnGetDataFromSource_Click(object sender, EventArgs e)
        {

            //From source
            lblSourceProgress.Visible = true;
            lblSourceProgress.ForeColor = System.Drawing.Color.Orange;
            lblSourceProgress.Text = "Getting products..";
            GetProductFromSource();
            lblSourceProgress.Text = "Getting customers..";
            GetCustomerFromSource();
            lblSourceProgress.Text = "Getting dates..";
            GetDatesFromSource();
            lblSourceProgress.Text = "Getting regions..";
            GetRegionsFromSource();
            lblSourceProgress.ForeColor = System.Drawing.Color.Green;
            lblSourceProgress.Text = "Finished.";

            //From dimension
            lblDestinationProgress.Visible = true;
            lblDestinationProgress.ForeColor = System.Drawing.Color.Orange;
            lblDestinationProgress.Text = "Getting products..";
            GetAllProductsFromDimension();
            lblDestinationProgress.Text = "Getting customers..";
            GetAllCustomersFromDimension();
            lblDestinationProgress.Text = "Getting dates..";
            GetAllDatesFromDimension();
            lblDestinationProgress.Text = "Getting regions..";
            GetAllRegionsFromDimension();
            lblDestinationProgress.ForeColor = System.Drawing.Color.Green;
            lblDestinationProgress.Text = "Finished.";

            //build fact table
            lblBuildProgress.Visible = true;
            lblBuildProgress.ForeColor = System.Drawing.Color.Orange;
            lblBuildProgress.Text = "Building the fact table..";
            BuildFactTable();
            lblBuildProgress.ForeColor = System.Drawing.Color.Green;
            lblBuildProgress.Text = "Finished.";


            //Get from fact table
            lblGetProgress.Visible = true;
            lblGetProgress.ForeColor = System.Drawing.Color.Orange;
            lblGetProgress.Text = "Getting from fact table..";
            GetFactTable();
            lblGetProgress.ForeColor = System.Drawing.Color.Green;
            lblGetProgress.Text = "Finished.";

            ////load all the blank charts
            btnLoadDateData_Click(sender, e);
            ProductDateChanged(5);
            ProductDateChanged(5);
            DateChanged(5);
            btnLoadCustomerData_Click(sender, e);
        }

        private void splitDates(string rawDate)
        {
            // split the date down and assign it to variables for later use
            string[] arrayDate = rawDate.Split('/');
            Int32 year = Convert.ToInt32(arrayDate[2]);
            Int32 month = Convert.ToInt32(arrayDate[1]);
            Int32 day = Convert.ToInt32(arrayDate[0]);

            DateTime myDate = new DateTime(year, month, day);

            String dayOfWeek = myDate.DayOfWeek.ToString();
            Int32 dayOfYear = myDate.DayOfYear;
            String monthName = myDate.ToString("MMMM");
            Int32 weekNumber = dayOfYear / 7;
            Boolean weekend = false;
            if (dayOfWeek == "Saturday" || dayOfWeek == "Sunday") weekend = true;

            // convert this to a database friendly format
            string dbDate = myDate.ToString("M/dd/yyyy");

            insertTimeDimension(dbDate, dayOfWeek, day, monthName, month, weekNumber, year, weekend, dayOfYear);
        }

        private void GetDatesFromSource()
        {
            //Create a list
            List<string> Dates = new List<string>();
            listBoxDates.DataSource = null;
            listBoxDates.Items.Clear();

            string connectionString = Properties.Settings.Default.Data_set_1ConnectionString;

            // create an instance of an object so we can connect to the database 
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                // open the database so we can get information out
                connection.Open();
                OleDbDataReader reader = null; // declare a reader (reads table in rows) object and assign it an initial null value  
                OleDbCommand getDates = new OleDbCommand("SELECT [Order Date], [Ship Date] from Sheet1", connection); // create an object to hold the sql query we want to execute

                reader = getDates.ExecuteReader();   // execute the sql query and assign it to reader
                while (reader.Read())                   // while reader reads the table/database
                {
                    Dates.Add(reader[0].ToString());   // add column 1 to the list and convert it to string
                    Dates.Add(reader[1].ToString());   // add column 2 to the list and convert it to string

                }
            }

            

            List<string> DatesFormatted = new List<string>();

            foreach (string date in Dates)
            {
                var dates = date.Split(new char[0], StringSplitOptions.RemoveEmptyEntries);
                DatesFormatted.Add(dates[0]);
            }

            listBoxDates.DataSource = DatesFormatted;


            // split the dates and insert every date in the list
            foreach (string date in DatesFormatted)
            {
                splitDates(date);
            }
        }

        private void GetProductFromSource()
        {
            // create a list to store my products in / list object to store the list of Products
            List<string> Products = new List<string>();
            // clear the listbox to ensure no old data is in there
            listBoxProducts.DataSource = null;
            listBoxProducts.Items.Clear();

            // create the database string
            string connectionString = Properties.Settings.Default.Data_set_1ConnectionString;

            // create an instance of an object so we can connect to the database 
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                // open the database so we can get information out
                connection.Open();
                OleDbDataReader reader = null; // declare a reader (reads table in rows) object and assign it an initial null value  
                OleDbCommand getProducts = new OleDbCommand("SELECT [Product ID], Category, [Sub-Category], [Product Name] from Sheet1", connection); // create an object to hold the sql query we want to execute

                reader = getProducts.ExecuteReader();   // execute the sql query and assign it to reader
                while (reader.Read())                   // while reader reads the table/database
                {
                    Products.Add(reader[0].ToString() + ", " + reader[1].ToString() + ", " + reader[2].ToString() + ", " + reader[3].ToString());   // - if index 3 is removed, Product ID won't show in listbox

                    string reference = Convert.ToString(reader["Product ID"]);
                    string category = Convert.ToString(reader["Category"]);
                    string subcategory = Convert.ToString(reader["Sub-Category"]);
                    string name = Convert.ToString(reader["Product Name"]);

                    insertProductDimension(category, subcategory, name, reference);

                }
            }

            
            listBoxProducts.DataSource = Products;        // assign Products list object to the listbox in the form as its source of data
        }

        private void GetCustomerFromSource()
        {
            // Create a list to store my customers in
            List<string> Customers = new List<string>();
            // Clear the listbox to ensure no old data is in there
            listBoxCustomers.DataSource = null;
            listBoxCustomers.Items.Clear();

            // Create the database string
            string connectionString = Properties.Settings.Default.Data_set_1ConnectionString;

            // Connect to the database using previous string
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                // Open the ACCESS connection
                connection.Open();
                OleDbDataReader reader = null;
                OleDbCommand getCustomers = new OleDbCommand("SELECT [Customer ID], [Customer Name], Country, City, State, [Postal Code], Region from Sheet1", connection);

                // Read the results (ExecuteReader gets only a single line)
                reader = getCustomers.ExecuteReader();
                while (reader.Read())
                {
                    // Read each column's data for a single line
                    Customers.Add(reader[0].ToString() + ", " + reader[1].ToString() + ", " + reader[2].ToString() + ", " + reader[3].ToString() + ", " + reader[4].ToString() + ", " + reader[5].ToString() + ", " + reader[6].ToString());

                    String name = Convert.ToString(reader["Customer Name"]);
                    String country = Convert.ToString(reader["Country"]);
                    String city = Convert.ToString(reader["City"]);
                    String state = Convert.ToString(reader["State"]);
                    String postalCode = Convert.ToString(reader["Postal Code"]);
                    String region = Convert.ToString(reader["Region"]);
                    String reference = Convert.ToString(reader["Customer ID"]);

                    insertCustomerDimension(name, country, city, state, postalCode, region, reference);
                }
            }


            // Bind results to the listbox
            listBoxCustomers.DataSource = Customers;
        }

        private void GetRegionsFromSource()
        {
            //create list to store data 
            List<string> Regions = new List<string>();
            //clear listbox to ensure old data is removed 
            listboxRegion.DataSource = null;
            listboxRegion.Items.Clear();
            //create database string 
            string connectionString = Properties.Settings.Default.Data_set_1ConnectionString;


            //retrieve data from table to show on listbox 
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                // open the connection
                connection.Open();
                OleDbDataReader reader = null;
                OleDbCommand getRegions = new OleDbCommand("SELECT DISTINCT Country, City, State, Region FROM sheet1", connection);

                //read the results
                reader = getRegions.ExecuteReader();
                while (reader.Read())
                {
                    Regions.Add(reader[0].ToString() + ", " + reader[1].ToString() + ", " + reader[2].ToString() + ", " + reader[3].ToString());

                    String country = Convert.ToString(reader["Country"]);
                    String city = Convert.ToString(reader["City"]);
                    String state = Convert.ToString(reader["State"]);
                    String region = Convert.ToString(reader["Region"]);

                    insertRegionDimension(country, city, state, region);
                }
            }

            
            //insert results into listbox
            listboxRegion.DataSource = Regions;
        }

        private void insertTimeDimension(string date, string dayName, int dayNumber, string monthName, int monthNumber, int weekNumber, int year, bool weekend, int dayOfYear)
        {
            //Create a connection to the MDF file
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
            {
                // open the sql connection
                myConnection.Open();
                // check if the product already exists in the database -  we do not need duplicates
                SqlCommand command = new SqlCommand("SELECT id FROM Time where date = @date", myConnection);
                command.Parameters.Add(new SqlParameter("@date", date));

                // create a variable and assign it as false by default
                Boolean exists = false;

                // run the command & read the results
                using (SqlDataReader reader = command.ExecuteReader())
                {

                    // if there are no rows it means the product exists, so update the var!
                    if (reader.HasRows) exists = true;
                }

                if (exists == false)
                {            //Insert the data
                    SqlCommand insertCommand = new SqlCommand(
                        "INSERT INTO Time (dayName, dayNumber, monthName, monthNumber, weekNumber, year, weekend, date, dayOfYear)" +
                        " VALUES (@dayName, @dayNumber, @monthName, @monthNumber, @weekNumber, @year, @weekend, @date, @dayOfYear) ", myConnection);
                    insertCommand.Parameters.Add(new SqlParameter("dayName", dayName));
                    insertCommand.Parameters.Add(new SqlParameter("dayNumber", dayNumber));
                    insertCommand.Parameters.Add(new SqlParameter("monthName", monthName));
                    insertCommand.Parameters.Add(new SqlParameter("monthNumber", monthNumber));
                    insertCommand.Parameters.Add(new SqlParameter("weekNumber", weekNumber));
                    insertCommand.Parameters.Add(new SqlParameter("year", year));
                    insertCommand.Parameters.Add(new SqlParameter("weekend", weekend));
                    insertCommand.Parameters.Add(new SqlParameter("date", date));
                    insertCommand.Parameters.Add(new SqlParameter("dayOfYear", dayOfYear));

                    // insert the line
                    int recordsAffected = insertCommand.ExecuteNonQuery();
                }

            }

        }

        private void insertProductDimension(string category, string subcategory, string name, string reference)
        {
            // create a connection to the mdf file
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            // create a variable and assign it as false by default
            Boolean exists = false;

            using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
            {
                // open the sql connection
                myConnection.Open();
                // check if the product already exists in the database -  we do not need duplicates
                SqlCommand command = new SqlCommand("SELECT id FROM Product where reference = @reference", myConnection);
                command.Parameters.Add(new SqlParameter("reference", reference));

                // run the command & read the results
                using (SqlDataReader reader = command.ExecuteReader())
                {

                    // if there are no rows it means the product exists, so update the var!
                    if (reader.HasRows) exists = true;
                }

                if (exists == false)
                {
                    SqlCommand insertCommand = new SqlCommand(
                        "INSERT INTO Product (category, subcategory, name, reference)" +
                        " VALUES (@category, @subcategory, @name, @reference) ", myConnection);
                    insertCommand.Parameters.Add(new SqlParameter("category", category));
                    insertCommand.Parameters.Add(new SqlParameter("subcategory", subcategory));
                    insertCommand.Parameters.Add(new SqlParameter("name", name));
                    insertCommand.Parameters.Add(new SqlParameter("reference", reference));

                    // insert the line
                    int recordsAffected = insertCommand.ExecuteNonQuery();
                }
            }
        }

        private void insertCustomerDimension(string name, string country, string city, string state, string postCode, string region, string reference)
        {
            // Create a connection to the MDF file
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            // Connect to the database using previous string
            using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
            {
                // Open the SQL connection
                myConnection.Open();
                // Check if the customer already exists in the database (for duplicates)
                SqlCommand command = new SqlCommand("SELECT reference FROM Customer WHERE reference = @reference", myConnection);
                // Add customer to the previous command
                command.Parameters.Add(new SqlParameter("reference", reference));

                // Create a variable and assign it as false by default
                Boolean exists = false;

                // Run the command & read the results
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    // If there are rows it means the customer exists 
                    if (reader.HasRows) exists = true;
                }

                // If the customer doesn't exist
                if (exists == false)
                {
                    SqlCommand insertCommand = new SqlCommand(
                        "INSERT INTO Customer (name, country, city, state, postalCode, region, reference)"
                        +
                        "VALUES (@name, @country, @city, @state, @postalCode, @region, @reference)", myConnection);
                    insertCommand.Parameters.Add(new SqlParameter("name", name));
                    insertCommand.Parameters.Add(new SqlParameter("country", country));
                    insertCommand.Parameters.Add(new SqlParameter("city", city));
                    insertCommand.Parameters.Add(new SqlParameter("state", state));
                    insertCommand.Parameters.Add(new SqlParameter("postalCode", postCode));
                    insertCommand.Parameters.Add(new SqlParameter("region", region));
                    insertCommand.Parameters.Add(new SqlParameter("reference", reference));

                    // Insert the line
                    int recordsAffected = insertCommand.ExecuteNonQuery();
                }
            }
        }

        private void insertRegionDimension(string country, string city, string state, string region)
        {
            //create a connection to the database 
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            //create a variable and assign it as false
            Boolean exists = false;

            //connect to the database using string
            using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
            {
                //open the SQL connection
                myConnection.Open();
                //check if the region already exists
                SqlCommand command = new SqlCommand("SELECT DISTINCT city FROM Region WHERE city = @city", myConnection);
                //add region to the previous command
                command.Parameters.Add(new SqlParameter("city", city));

                //run the command and read the results
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    //if there are rows it means the region exists 
                    if (reader.HasRows) exists = true;
                }

                //if the region doesnt exist
                if (exists == false)
                {
                    SqlCommand insertCommand = new SqlCommand(
                        "INSERT INTO Region (country, city, state, region) VALUES (@country, @city, @state, @region)", myConnection);
                    insertCommand.Parameters.Add(new SqlParameter("country", country));
                    insertCommand.Parameters.Add(new SqlParameter("city", city));
                    insertCommand.Parameters.Add(new SqlParameter("state", state));
                    insertCommand.Parameters.Add(new SqlParameter("region", region));

                    //show user that records have been affected 
                    int recordsAffected = insertCommand.ExecuteNonQuery();
                }
            }
        }

        //This needs replacing with whatever button we decide to make
        private void btnGetDataFromDestination_Click(object sender, EventArgs e)
        {
            lblDestinationProgress.Visible = true;
            lblDestinationProgress.ForeColor = System.Drawing.Color.Orange;
            lblDestinationProgress.Text = "Getting products..";
            GetAllProductsFromDimension();
            lblDestinationProgress.Text = "Getting customers..";
            GetAllCustomersFromDimension();
            lblDestinationProgress.Text = "Getting dates..";
            GetAllDatesFromDimension();
            lblDestinationProgress.Text = "Getting regions..";
            GetAllRegionsFromDimension();
            lblDestinationProgress.ForeColor = System.Drawing.Color.Green;
            lblDestinationProgress.Text = "Finished.";
        }

        private void GetAllDatesFromDimension()
        {
            // create a list to store the data in
            List<string> DestinationDates = new List<string>();

            // Clear the listbox to ensure no old data is in there
            listBoxDatesDimension.DataSource = null;
            listBoxDatesDimension.Items.Clear();

            // create a connection to the MDF file
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
            {
                // open the SQL connection 
                myConnection.Open();
                // check if the date alerady exists in the database - we do not need duplaictes!
                SqlCommand command = new SqlCommand("SELECT Id, dayName, dayNumber, monthName, monthNumber, weekNumber, year," +
                    " weekend, date, dayOfYear FROM Time", myConnection);

                using (SqlDataReader reader = command.ExecuteReader())
                {
                    if (reader.HasRows) // there is data so do something with it
                    {
                        while (reader.Read()) // this loop actually gets the data...
                        {
                            // do something with each record here
                            // build what I want the listbox to show (note everything is a string!)
                            string id = reader["Id"].ToString();
                            string dayName = reader["dayName"].ToString();
                            string dayNumber = reader["dayNumber"].ToString();
                            string monthName = reader["monthName"].ToString();
                            string monthNumber = reader["monthNumber"].ToString();
                            string weekNumber = reader["weekNumber"].ToString();
                            string year = reader["year"].ToString();
                            string weekend = reader["weekend"].ToString();
                            string date = reader["date"].ToString();
                            string dayOfYear = reader["dayOfYear"].ToString();

                            string text;

                            text = "ID = " + id + ", Day Name = " + dayName + ", Day Number = " + dayNumber +
                                ", Month Name = " + monthName + ", Month Number = " + monthNumber + ", Week Number = " + weekNumber +
                                ", Year = " + year + ", Weekend = " + weekend + ", Date = " + date + "Day of year = " + dayOfYear;

                            DestinationDates.Add(text);
                        }
                    }
                    else // there was no data - show an error maybe?
                    {
                        DestinationDates.Add("No data present in the Dates Dimension");
                    }
                }
            }

            // bind the listbox to the list
            listBoxDatesDimension.DataSource = DestinationDates;


        }
        private void GetAllProductsFromDimension()
        {
            // create a list to store the data in
            List<string> DestinationProducts = new List<string>();

            // Clear the listbox to ensure no old data is in there
            listBoxProductsDimension.DataSource = null;
            listBoxProductsDimension.Items.Clear();

            // create a connection to the MDF file
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
            {
                // open the SQL connection 
                myConnection.Open();
                // check if the date alerady exists in the database - we do not need duplaictes!
                SqlCommand command = new SqlCommand("SELECT Id, category, subcategory, name, reference FROM Product", myConnection);

                using (SqlDataReader reader = command.ExecuteReader())
                {
                    if (reader.HasRows) // there is data so do something with it
                    {
                        while (reader.Read()) // this loop actually gets the data...
                        {
                            // do something with each record here
                            // build what I want the listbox to show (note everything is a string!)
                            string id = reader["Id"].ToString();
                            string category = reader["category"].ToString();
                            string subcategory = reader["subcategory"].ToString();
                            string name = reader["name"].ToString();
                            string reference = reader["reference"].ToString();

                            string text;

                            text = "ID = " + id + ", Category = " + category + ", SubCategory = " + subcategory + ", Name = " + name + ", Reference = " + reference;

                            DestinationProducts.Add(text);
                        }
                    }
                    else // there was no data - show an error maybe?
                    {
                        DestinationProducts.Add("No data present in the Product Dimension");
                    }
                }
            }

            // bind the listbox to the list
            listBoxProductsDimension.DataSource = DestinationProducts;
        }

        private void GetAllCustomersFromDimension()
        {
            // Create a list to store the data in
            List<string> DestinationCustomers = new List<string>();
            // Clear the listbox to ensure no old data is in there
            listBoxCustomersDimension.DataSource = null;
            listBoxCustomersDimension.Items.Clear();

            // Create a connection to the MDF file
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            // Connect to the database using previous string
            using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
            {
                // Open the SQL connection
                myConnection.Open();
                // Check if the customer already exists in the database (for duplicates)
                SqlCommand command = new SqlCommand("SELECT Id, name, country, city, state, postalCode, region, reference FROM Customer", myConnection);

                using (SqlDataReader reader = command.ExecuteReader())
                {
                    // If there are rows it means there is data
                    if (reader.HasRows)
                    {
                        // Read the results
                        while (reader.Read())
                        {
                            // Build what I want the listbox to show
                            string Id = reader["Id"].ToString();
                            string name = reader["name"].ToString();
                            string country = reader["country"].ToString();
                            string city = reader["city"].ToString();
                            string state = reader["state"].ToString();
                            string postalCode = reader["postalCode"].ToString();
                            string region = reader["region"].ToString();
                            string reference = reader["reference"].ToString();

                            string text;

                            text = "ID = " + Id + ", Name = " + name + ", Country = " + country +
                                ", City = " + city + ", State = " + state + ", Postal Code = " +
                                postalCode + ", Region = " + region + ", Reference = " + reference;

                            DestinationCustomers.Add(text);
                        }
                    }
                    // If there was no data
                    else
                    {
                        DestinationCustomers.Add("No data present in the Customer Dimension");
                    }
                }
            }
            // Bind results to the listbox
            listBoxCustomersDimension.DataSource = DestinationCustomers;
        }

        private void GetAllRegionsFromDimension()
        {
            //create a list to store data
            List<string> RegionDimension = new List<string>();
            //clear the listbox to ensure no old data exists
            listboxRegionDimension.DataSource = null;
            listboxRegionDimension.Items.Clear();

            //create a connection to the database 
            string connectionString = Properties.Settings.Default.DestinationDatabaseConnectionString;

            //connect to the database using string
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                //open the SQL connection
                connection.Open();
                //check if the region already exists in the database
                SqlCommand command = new SqlCommand("SELECT DISTINCT (Id), country, city, state, region FROM Region", connection);

                using (SqlDataReader reader = command.ExecuteReader())
                {
                    //check for rows
                    if (reader.HasRows)
                    {
                        //read the results
                        while (reader.Read())
                        {
                            //build the listbox to show data
                            string Id = reader["Id"].ToString();
                            string country = reader["country"].ToString();
                            string city = reader["city"].ToString();
                            string state = reader["state"].ToString();
                            string region = reader["region"].ToString();

                            string text;

                            text = "ID = " + Id + ", Country = " + country + ", City = " + city + ", State = "
                                + state + ", Region = " + region;

                            RegionDimension.Add(text);
                        }
                    }
                    //if there was no data
                    else
                    {
                        RegionDimension.Add("No data present in the Region Dimension");
                    }
                }
            }
            //bind results to listbox
            listboxRegionDimension.DataSource = RegionDimension;
        }

        private int GetDateId(string date)
        {
            // Remove the time from the date
            var dateSplit = date.Split(new char[0], StringSplitOptions.RemoveEmptyEntries);
            // Overwrite the original value
            date = dateSplit[0];

            // Split the clean date down and assign it to variables for later use.
            string[] arrayDate = date.Split('/');
            Int32 year = Convert.ToInt32(arrayDate[2]);
            Int32 month = Convert.ToInt32(arrayDate[1]);
            Int32 day = Convert.ToInt32(arrayDate[0]);

            DateTime myDate = new DateTime(year, month, day);

            string dbDate = myDate.ToString("MM/dd/yyyy");

            // create a connection to the mdf file
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString; // this has reference + this is an independent function

            using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
            {
                // open the sql connection
                myConnection.Open();
                // check if the product already exists in the database -  we do not need duplicates
                SqlCommand command = new SqlCommand("SELECT id FROM Time WHERE date = @date", myConnection);
                command.Parameters.Add(new SqlParameter("@date", dbDate)); // unlike 'name' - can have two products with same name

                // run the command & read the results
                using (SqlDataReader reader = command.ExecuteReader())
                {

                    // if there are no rows it means the product exists, so update the var!
                    if (reader.HasRows)
                    {
                        // Read the results
                        while (reader.Read())
                        {
                            return Convert.ToInt32(reader["id"]);
                        }
                    }
                }
            }
            return 0;

        }
        private int GetProductId(string reference)
        {
            // create a connection to the mdf file
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString; // this has reference + this is an independent function

            using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
            {
                // open the sql connection
                myConnection.Open();
                // check if the product already exists in the database -  we do not need duplicates
                SqlCommand command = new SqlCommand("SELECT Id FROM Product where reference = @reference", myConnection);
                command.Parameters.Add(new SqlParameter("reference", reference)); // unlike 'name' - can have two products with same name

                // run the command & read the results
                using (SqlDataReader reader = command.ExecuteReader())
                {

                    // if there are no rows it means the product exists, so update the var!
                    if (reader.HasRows)
                    {
                        // Read the results
                        while (reader.Read())
                        {
                            return Convert.ToInt32(reader["id"]);
                        }
                    }
                }
            }

            return 0;
        }

        private int GetCustomerId(string reference)
        {
            // Create a connection to the MDF file
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            // Connect to the database using previous string
            using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
            {
                // Open the SQL connection
                myConnection.Open();
                // Check if the customer already exists in the database (for duplicates)
                SqlCommand command = new SqlCommand("SELECT Id FROM Customer WHERE reference = @reference", myConnection);
                // Add customer to the previous command
                command.Parameters.Add(new SqlParameter("reference", reference));

                // Run the command & read the results
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    // If there are rows it means the date exists
                    if (reader.HasRows)
                    {
                        // Read the results
                        while (reader.Read())
                        {
                            return Convert.ToInt32(reader["id"]);
                        }
                    }
                }
            }
            return 0;
        }

        private int GetRegionId(string city)
        {
            //create a connection to the MDF file
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            //connect to the database
            using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
            {
                //open the SQL connection
                myConnection.Open();
                //check if region already exists in database
                SqlCommand command = new SqlCommand("SELECT Id FROM Region WHERE city = @city", myConnection);
                //add ewgion to the previous command
                command.Parameters.Add(new SqlParameter("city", city));

                //run the command & read the results
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    //if there are rows it means the date exists
                    if (reader.HasRows)
                    {
                        //read the results
                        while (reader.Read())
                        {
                            return Convert.ToInt32(reader["id"]);
                        }
                    }
                }
            }
            return 0;
        }

        private void BuildFactTable()
        {
            // create the database string
            string connectionString = Properties.Settings.Default.Data_set_1ConnectionString;

            // create an instance of an object so we can connect to the database 
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                // open the database so we can get information out
                connection.Open();
                OleDbDataReader reader = null; // declare a reader (reads table in rows) object and assign it an initial null value  
                OleDbCommand getProducts = new OleDbCommand("SELECT ID, [Row ID], [Order ID], [Order Date], [Ship Date], " +
                    "[Ship Mode], [Customer ID], [Customer Name], Segment, Country, City, State, [Postal Code], [Product ID], " +
                    " Region, Category, [Sub-Category], [Product Name], Sales, Quantity, Profit, Discount from Sheet1", connection); // create an object to hold the sql query we want to execute

                reader = getProducts.ExecuteReader();   // execute the sql query and assign it to reader
                while (reader.Read())                   // while reader reads the table/database
                {
                    // get a line of daya from source

                    // get the numeric values
                    Double sales = Convert.ToDouble(reader["Sales"]);
                    Int32 quantity = Convert.ToInt32(reader["Quantity"]);
                    Double profit = Convert.ToDouble(reader["Profit"]);
                    Double discount = Convert.ToDouble(reader["Discount"]);

                    // get the dimension IDs
                    Int32 productId = GetProductId(reader["Product ID"].ToString());
                    Int32 timeId = GetDateId(reader["Order Date"].ToString());
                    Int32 customerId = GetCustomerId(reader["Customer ID"].ToString());
                    Int32 regionId = GetRegionId(reader["City"].ToString()); 

                    // Insert it into the database
                    insertIntoFactTable(productId, timeId, customerId, regionId, sales, discount, profit, quantity);
                }
            }

            
        }

        private void insertIntoFactTable(int productId, int timeId, int customerId, int regionId, double value, double discount, double profit, double quantity)
        {
            // Create a connection to the MDF file
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            // Connect to the database using previous string
            using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
            {
                // Open the SQL connection
                myConnection.Open();
                // Check if the data already exists in the database (for duplicates)
                SqlCommand command = new SqlCommand("SELECT productId, timeId, customerId, regionId FROM FactTableEmpty WHERE productId = @productId AND timeId = @timeId AND customerId = @customerId AND regionId = @regionId", myConnection);
                // Add data to the previous command
                command.Parameters.Add(new SqlParameter("productId", productId));
                command.Parameters.Add(new SqlParameter("timeId", timeId));
                command.Parameters.Add(new SqlParameter("customerId", customerId));
                command.Parameters.Add(new SqlParameter("regionId", regionId));


                // Create a variable and assign it as false by default
                Boolean exists = false;

                // Run the command & read the results
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    // If there are rows it means the data exists
                    if (reader.HasRows) exists = true;
                }

                // If the data doesn't exist
                if (exists == false)
                {
                    SqlCommand insertCommand = new SqlCommand(
                        "INSERT INTO FactTableEmpty (productId, timeId, customerId, regionId, value, discount, profit, quantity)"
                        +
                        "VALUES (@productId, @timeId, @customerId, @regionId, @value, @discount, @profit, @quantity)", myConnection);
                    insertCommand.Parameters.Add(new SqlParameter("productId", productId));
                    insertCommand.Parameters.Add(new SqlParameter("timeId", timeId));
                    insertCommand.Parameters.Add(new SqlParameter("customerId", customerId));
                    insertCommand.Parameters.Add(new SqlParameter("regionId", regionId));
                    insertCommand.Parameters.Add(new SqlParameter("value", value));
                    insertCommand.Parameters.Add(new SqlParameter("discount", discount));
                    insertCommand.Parameters.Add(new SqlParameter("profit", profit));
                    insertCommand.Parameters.Add(new SqlParameter("quantity", quantity));

                    // Insert the line
                    int recordsAffected = insertCommand.ExecuteNonQuery();
                }
            }
        }

        private void GetFactTable()
        {
            // Create a list to store the data in
            List<string> FactTable = new List<string>();
            // Clear the listbox to ensure no old data is in there
            listBoxFactTable.DataSource = null;
            listBoxFactTable.Items.Clear();

            // Create a connection to the MDF file
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            // Connect to the database using previous string
            using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
            {
                // Open the SQL connection
                myConnection.Open();
                // Check if the data already exists in the database (for duplicates)
                SqlCommand command = new SqlCommand("SELECT productId, timeId, customerId, regionId, value, discount, profit, quantity FROM FactTableEmpty", myConnection);

                using (SqlDataReader reader = command.ExecuteReader())
                {
                    // If there are rows it means there is data
                    if (reader.HasRows)
                    {
                        // Read the results
                        while (reader.Read())
                        {
                            // Build what I want the listbox to show
                            string productId = reader["productId"].ToString();
                            string timeId = reader["timeId"].ToString();
                            string customerId = reader["customerId"].ToString();
                            string regionId = reader["regionId"].ToString();
                            string value = reader["value"].ToString();
                            string discount = reader["discount"].ToString();
                            string profit = reader["profit"].ToString();
                            string quantity = reader["quantity"].ToString();

                            string text;

                            text = "Product ID = " + productId + ", Time ID = " + timeId + ", Customer ID = " +
                                customerId + ", Region ID = " + regionId + ", Sales = " + value + ", Discount = " + discount + ", Profit = " +
                                profit + ", Quantity = " + quantity;

                            FactTable.Add(text);
                        }
                    }
                    // If there was no data
                    else
                    {
                        FactTable.Add("No data present in the Fact Table");
                    }
                }
            }
            // Bind results to the listbox
            listBoxFactTable.DataSource = FactTable;
        }

        // event handler button for Get Products
        private void btnGetProducts_Click(object sender, EventArgs e)
        {
            GetProductFromSource();
        }

        private void btnGetCustomers_Click(object sender, EventArgs e)
        {
            GetCustomerFromSource();
        }

        private void btnGetDates_Click(object sender, EventArgs e)
        {
            GetDatesFromSource();
        }

        private void btnGetRegionsFromSource_Click(object sender, EventArgs e)
        {
            GetRegionsFromSource();
        }

        private void btnGetProductFromDimension_Click(object sender, EventArgs e)
        {
            GetAllProductsFromDimension();
        }

        private void btnGetCustomersFromDimension_Click(object sender, EventArgs e)
        {
            GetAllCustomersFromDimension();
        }

        private void btnGetDatesFromDimension_Click(object sender, EventArgs e)
        {
            GetAllDatesFromDimension();
        }

        private void btnGetRegionsFromDimension_Click(object sender, EventArgs e)
        {
            GetAllRegionsFromDimension();
        }

        private void btnGetFromFactTable_Click(object sender, EventArgs e)
        {
            lblGetProgress.Visible = true;
            lblGetProgress.ForeColor = System.Drawing.Color.Orange;
            lblGetProgress.Text = "Getting from fact table..";
            GetFactTable();
            lblGetProgress.ForeColor = System.Drawing.Color.Green;
            lblGetProgress.Text = "Finished.";
        }

        private void btnBuildFactTable_Click(object sender, EventArgs e)
        {
            lblBuildProgress.Visible = true;
            lblBuildProgress.ForeColor = System.Drawing.Color.Orange;
            lblBuildProgress.Text = "Building the fact table..";
            BuildFactTable();
            lblBuildProgress.ForeColor = System.Drawing.Color.Green;
            lblBuildProgress.Text = "Finished.";
        }

        private void CustomerDateChanged(Int32 CustomerAmount)
        {
            // Get the selected date
            string SelectedStart = Convert.ToString(monthCalendarCustomer.SelectionStart);
            string SelectedEnd = Convert.ToString(monthCalendarCustomer.SelectionEnd);

            // Split date and time
            var SelectedStartSplit = SelectedStart.Split(new char[0], StringSplitOptions.RemoveEmptyEntries);
            var SelectedEndSplit = SelectedEnd.Split(new char[0], StringSplitOptions.RemoveEmptyEntries);

            // New array with D/M/Y split
            string[] arrayStart = SelectedStartSplit[0].Split('/');
            string[] arrayEnd = SelectedEndSplit[0].Split('/');

            // Assign to the appropriate variables
            Int32 StartYear = Convert.ToInt32(arrayStart[2]);
            Int32 StartMonth = Convert.ToInt32(arrayStart[1]);
            Int32 StartDay = Convert.ToInt32(arrayStart[0]);

            Int32 EndYear = Convert.ToInt32(arrayEnd[2]);
            Int32 EndMonth = Convert.ToInt32(arrayEnd[1]);
            Int32 EndDay = Convert.ToInt32(arrayEnd[0]);

            DateTime StartingDate = new DateTime(StartYear, StartMonth, StartDay);
            DateTime EndingDate = new DateTime(EndYear, EndMonth, EndDay);

            // Convert this to a database friendly format
            string StartDate = StartingDate.ToString("MM/dd/yyyy");
            string EndDate = EndingDate.ToString("MM/dd/yyyy");

            MostProfitableCustomers(StartDate, EndDate, CustomerAmount);
            CustomersWithMostDiscounts(StartDate, EndDate, CustomerAmount);
            ProfitOnDate(StartDate, EndDate);
            ActiveCustomersByDate(StartDate, EndDate);
        }

        private void monthCalendarCustomer_DateSelected(object sender, DateRangeEventArgs e)
        {
            // Get the customer amount
            Int32 CustomerAmount = Convert.ToInt32(numericUpDownCustomer.Value);
            CustomerDateChanged(CustomerAmount);
        }

        private void numericUpDownCustomer_ValueChanged(object sender, EventArgs e)
        {
            // Get the customer amount
            Int32 CustomerAmount = Convert.ToInt32(numericUpDownCustomer.Value);
            CustomerDateChanged(CustomerAmount);
        }

        private void MostProfitableCustomers(string StartDate, string EndDate, Int32 CustomerAmount)
        {
            // Dictionary to store customer name and profit
            Dictionary<String, Double> CustomerProfit = new Dictionary<String, Double>();
            // Dictionary to store customer name and value
            Dictionary<String, Double> CustomerValue = new Dictionary<String, Double>();

            // Create a connection to the MDF file
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            // Connect to the database using previous string
            using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
            {
                // Open the SQL connection
                myConnection.Open();
                SqlCommand command = new SqlCommand("SELECT TOP (@CustomerAmount) Customer.name AS CustomerName, SUM(FactTableEmpty.profit) AS Profit, SUM(FactTableEmpty.value) AS Value FROM FactTableEmpty JOIN Customer ON FactTableEmpty.customerId = Customer.Id JOIN Time ON FactTableEmpty.timeId = Time.id WHERE Time.date BETWEEN @StartDate AND @EndDate GROUP BY Customer.name ORDER BY Profit DESC; ", myConnection);
                command.Parameters.Add(new SqlParameter("@StartDate", StartDate));
                command.Parameters.Add(new SqlParameter("@EndDate", EndDate));
                command.Parameters.Add(new SqlParameter("@CustomerAmount", CustomerAmount));

                using (SqlDataReader reader = command.ExecuteReader())
                {
                    // If there are rows it means there is data
                    if (reader.HasRows)
                    {
                        // Read the results
                        while (reader.Read())
                        {
                            // Add the profit for this customer
                            CustomerProfit.Add(Convert.ToString(reader["CustomerName"]), Convert.ToInt32(reader["Profit"])); //converted to int to display as whole number on chart
                            // Add the value for this customer
                            CustomerValue.Add(Convert.ToString(reader["CustomerName"]), Convert.ToInt32(reader["Value"])); //converted to int to display as whole number on chart
                        }
                    }
                    // If there was no data
                    else
                    {
                        CustomerProfit.Add("No data found", 0);
                        CustomerValue.Add("No data found", 0);
                    }
                }
            }

            // Ensure that the chart is clear
            chartMostProfitableCustomers.Series[0].Points.Clear();
            chartMostProfitableCustomers.Series[1].Points.Clear();

            // To build a chart:
            chartMostProfitableCustomers.DataSource = CustomerProfit;
            // Add the profits
            foreach (var item in CustomerProfit)
            {
                chartMostProfitableCustomers.Series[0].Points.AddXY(item.Key, item.Value);
            }
            // Add the values
            foreach (var item in CustomerValue)
            {
                chartMostProfitableCustomers.Series[1].Points.AddXY(item.Key, item.Value);
            }
            chartMostProfitableCustomers.DataBind();

            // If profit falls below 0 make it stand out
            foreach (DataPoint dp in chartMostProfitableCustomers.Series[0].Points)
            {
                if (dp.YValues[0] < 0)
                {
                    dp.Color = Color.Red;
                }
            }
        }

        private void CustomersWithMostDiscounts(string StartDate, string EndDate, Int32 CustomerAmount)
        {
            // Dictionary to store customer name and discount
            Dictionary<String, Double> CustomerDiscount = new Dictionary<String, Double>();

            // Create a connection to the MDF file
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            // Connect to the database using previous string
            using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
            {
                // Open the SQL connection
                myConnection.Open();
                SqlCommand command = new SqlCommand("SELECT TOP (@CustomerAmount) Customer.name AS CustomerName, SUM(FactTableEmpty.discount) AS Discount FROM FactTableEmpty JOIN Customer ON FactTableEmpty.customerId = Customer.Id JOIN Time ON FactTableEmpty.timeId = Time.id WHERE Time.date BETWEEN @StartDate AND @EndDate GROUP BY Customer.name ORDER BY Discount DESC; ", myConnection);
                command.Parameters.Add(new SqlParameter("@StartDate", StartDate));
                command.Parameters.Add(new SqlParameter("@EndDate", EndDate));
                command.Parameters.Add(new SqlParameter("@CustomerAmount", CustomerAmount));

                using (SqlDataReader reader = command.ExecuteReader())
                {
                    // If there are rows it means there is data
                    if (reader.HasRows)
                    {
                        // Read the results
                        while (reader.Read())
                        {
                            // Add the discount for this customer
                            CustomerDiscount.Add(Convert.ToString(reader["CustomerName"]), Convert.ToDouble(reader["Discount"]));
                        }
                    }
                    // If there was no data
                    else
                    {
                        CustomerDiscount.Add("No data found", 0);
                    }
                }
            }

            // To build a chart:
            chartCustomersWithMostDiscounts.DataSource = CustomerDiscount;
            // Add the customer names
            chartCustomersWithMostDiscounts.Series[0].XValueMember = "Key";
            // Add the discounts
            chartCustomersWithMostDiscounts.Series[0].YValueMembers = "Value";
            chartCustomersWithMostDiscounts.DataBind();
        }

        private void ProfitOnDate(string StartDate, string EndDate)
        {
            // Dictionary to store date and profit
            Dictionary<String, Double> DateProfit = new Dictionary<String, Double>();

            // Create a connection to the MDF file
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            // Connect to the database using previous string
            using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
            {
                // Open the SQL connection
                myConnection.Open();
                SqlCommand command = new SqlCommand("SELECT Time.date AS Date, SUM(profit) AS Profit FROM FactTableEmpty JOIN Time ON FactTableEmpty.timeId = Time.id WHERE Time.date BETWEEN @StartDate AND @EndDate GROUP BY Time.date; ", myConnection);
                command.Parameters.Add(new SqlParameter("@StartDate", StartDate));
                command.Parameters.Add(new SqlParameter("@EndDate", EndDate));

                using (SqlDataReader reader = command.ExecuteReader())
                {
                    // If there are rows it means there is data
                    if (reader.HasRows)
                    {
                        // Read the results
                        while (reader.Read())
                        {
                            // Split date and time
                            var Date = reader["Date"].ToString();
                            var SplitDate = Date.Split(new char[0], StringSplitOptions.RemoveEmptyEntries);

                            // Add the profit for this date
                            DateProfit.Add(SplitDate[0], Convert.ToDouble(reader["Profit"]));
                        }
                    }
                    // If there was no data
                    else
                    {
                        DateProfit.Add("No data found", 0);
                    }
                }
            }

            // To build a chart:
            chartProfitOnDate.DataSource = DateProfit;
            // Add the dates
            chartProfitOnDate.Series[0].XValueMember = "Key";
            // Add the profits
            chartProfitOnDate.Series[0].YValueMembers = "Value";
            chartProfitOnDate.DataBind();

            // If profit falls below or above 0 make it stand out
            foreach (DataPoint dp in chartProfitOnDate.Series[0].Points)
            {
                if (dp.YValues[0] < 0)
                {
                    dp.LabelForeColor = Color.Red;
                    dp.Color = Color.Red;
                }
                if (dp.YValues[0] > 0)
                {
                    dp.LabelForeColor = Color.Green;
                    dp.Color = Color.Green;
                }
            }
        }

        private void ActiveCustomersByDate(string StartDate, string EndDate)
        {
            // Dictionary to store date and customer amount
            Dictionary<String, Int32> DateAmount = new Dictionary<String, Int32>();

            // Create a connection to the MDF file
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            // Connect to the database using previous string
            using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
            {
                // Open the SQL connection
                myConnection.Open();
                SqlCommand command = new SqlCommand("SELECT Time.date AS Date, COUNT(DISTINCT customerId) AS Customers FROM FactTableEmpty JOIN Time ON FactTableEmpty.timeId = Time.id WHERE Time.date BETWEEN @StartDate AND @EndDate GROUP BY Time.date; ", myConnection);
                command.Parameters.Add(new SqlParameter("@StartDate", StartDate));
                command.Parameters.Add(new SqlParameter("@EndDate", EndDate));

                using (SqlDataReader reader = command.ExecuteReader())
                {
                    // If there are rows it means there is data
                    if (reader.HasRows)
                    {
                        // Read the results
                        while (reader.Read())
                        {
                            // Split date and time
                            var Date = reader["Date"].ToString();
                            var SplitDate = Date.Split(new char[0], StringSplitOptions.RemoveEmptyEntries);

                            // Add the customer amount for this date
                            DateAmount.Add(SplitDate[0], Convert.ToInt32(reader["Customers"]));
                        }
                    }
                    // If there was no data
                    else
                    {
                        DateAmount.Add("No data found", 0);
                    }
                }
            }

            // To build a chart:
            chartActiveCustomersByDate.DataSource = DateAmount;
            // Add the dates
            chartActiveCustomersByDate.Series[0].XValueMember = "Key";
            // Add the customer amount
            chartActiveCustomersByDate.Series[0].YValueMembers = "Value";
            chartActiveCustomersByDate.DataBind();
        }

        private void btnLoadCustomerData_Click(object sender, EventArgs e)
        {
            // Load all neccesary objects
            lblCustomerSelectDate.Visible = true;
            monthCalendarCustomer.Visible = true;
            lblCustomerAmount.Visible = true;
            numericUpDownCustomer.Visible = true;
            lblMostProfitableCustomers.Visible = true;
            chartMostProfitableCustomers.Visible = true;
            lblCustomersWithMostDiscounts.Visible = true;
            chartCustomersWithMostDiscounts.Visible = true;
            lblProfitOnDate.Visible = true;
            chartProfitOnDate.Visible = true;
            lblActiveCustomersByDate.Visible = true;
            chartActiveCustomersByDate.Visible = true;
            // Preset the customer amount to 3
            CustomerDateChanged(Convert.ToInt32(numericUpDownCustomer.Text));
        }

        // Event handler for when a tab is changed
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // If Customer Dashboard tab is selected
            if (tabControl1.SelectedTab.Name == "tabCustomerDashboard")
            {
                // Set the calendar's date to 2014 instead of 2019 for dashboard testing convenience
                DateTime olddateStart = Convert.ToDateTime("01/01/2014");
                DateTime olddateEnd = Convert.ToDateTime("31/01/2014");
                monthCalendarCustomer.SelectionStart = olddateStart;
                monthCalendarCustomer.SelectionEnd = olddateEnd;
                // Maximize the window
                this.WindowState = FormWindowState.Maximized;
            }
            // If another tab is selected
            else
            {
                // Return to the original size
                this.WindowState = FormWindowState.Normal;
            }
        }

         private void MostProfitableProduct(string StartDate, string EndDate, Int32 ProductAmount)
        {
            Dictionary<String, Decimal> ProductProfits = new Dictionary<String, Decimal>();

            String connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            // Connect to the database using previous string
            using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
            {
                // Open the SQL connection
                myConnection.Open();
                SqlCommand command = new SqlCommand(
                "SELECT TOP (@ProductAmount) Product.subcategory AS ProductSubcategory, SUM(FactTableEmpty.profit) AS Profit FROM FactTableEmpty JOIN Product ON FactTableEmpty.productId = Product.Id JOIN Time ON FactTableEmpty.timeId = Time.id WHERE Time.date BETWEEN @StartDate AND @EndDate GROUP BY Product.subcategory ORDER BY Profit DESC; ", myConnection );
                // Add data to the previous command
                command.Parameters.Add(new SqlParameter("@StartDate", StartDate));
                command.Parameters.Add(new SqlParameter("@EndDate", EndDate));
                command.Parameters.Add(new SqlParameter("@ProductAmount", ProductAmount));

                using (SqlDataReader reader = command.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            ProductProfits.Add(Convert.ToString(reader["ProductSubcategory"]), Convert.ToInt32(reader["Profit"]));//convert to int to deisplay as whole number on chart
                        }
                    }

                    else
                    {
                        ProductProfits.Add("No data found", 0);
                    }
                }                 
            }

            chartMostProfitableProducts.DataSource = ProductProfits;
            chartMostProfitableProducts.Series[0].XValueMember = "Key";
            chartMostProfitableProducts.Series[0].YValueMembers = "Value";
            chartMostProfitableProducts.DataBind();

            // If profit falls below or above 0 make it stand out
            foreach (DataPoint dp in chartMostProfitableProducts.Series[0].Points)
            {
                if (dp.YValues[0] < 0)
                {
                    dp.LabelForeColor = Color.Red;
                    dp.Color = Color.Red;
                }
                if (dp.YValues[0] > 0)
                {
                    dp.LabelForeColor = Color.Green;
                    dp.Color = Color.Green;
                }
            }

        }

        private void MostValuableProduct(string StartDate, string EndDate, Int32 ProductAmount)
        {
            Dictionary<String, Decimal> ProductValue = new Dictionary<String, Decimal>();

            String connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            // Connect to the database using previous string
            using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
            {
                // Open the SQL connection
                myConnection.Open();
                SqlCommand command = new SqlCommand(
                "SELECT TOP (@ProductAmount) Product.subcategory AS ProductSubcategory, SUM(FactTableEmpty.value) AS Value FROM FactTableEmpty JOIN Product ON FactTableEmpty.productId = Product.Id JOIN Time ON FactTableEmpty.timeId = Time.id WHERE Time.date BETWEEN @StartDate AND @EndDate GROUP BY Product.subcategory ORDER BY Value DESC; ", myConnection);
                // Add data to the previous command
                command.Parameters.Add(new SqlParameter("@StartDate", StartDate));
                command.Parameters.Add(new SqlParameter("@EndDate", EndDate));
                command.Parameters.Add(new SqlParameter("@ProductAmount", ProductAmount));

                using (SqlDataReader reader = command.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            ProductValue.Add(Convert.ToString(reader["ProductSubcategory"]), Convert.ToInt32(reader["Value"]));
                        }
                    }

                    else
                    {
                        ProductValue.Add("No data found", 0);
                    }
                }
            }

            chartMostValuableProducts.DataSource = ProductValue;
            chartMostValuableProducts.Series[0].XValueMember = "Key";
            chartMostValuableProducts.Series[0].YValueMembers = "Value";
            chartMostValuableProducts.DataBind();
        }

        private void ActiveProductsByDate(string StartDate, string EndDate)
        {
            // Dictionary to store date and customer amount
            Dictionary<String, Int32> ProductDateAmount = new Dictionary<String, Int32>();

            // Create a connection to the MDF file
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            // Connect to the database using previous string
            using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
            {
                // Open the SQL connection
                myConnection.Open();
                SqlCommand command = new SqlCommand("SELECT Time.date AS Date, COUNT(DISTINCT productId) AS Products FROM FactTableEmpty JOIN Time ON FactTableEmpty.timeId = Time.id WHERE Time.date BETWEEN @StartDate AND @EndDate GROUP BY Time.date; ", myConnection);
                command.Parameters.Add(new SqlParameter("@StartDate", StartDate));
                command.Parameters.Add(new SqlParameter("@EndDate", EndDate));

                using (SqlDataReader reader = command.ExecuteReader())
                {
                    // If there are rows it means there is data
                    if (reader.HasRows)
                    {
                        // Read the results
                        while (reader.Read())
                        {
                            // Split date and time
                            var Date = reader["Date"].ToString();
                            var SplitDate = Date.Split(new char[0], StringSplitOptions.RemoveEmptyEntries);

                            // Add the customer amount for this date
                            ProductDateAmount.Add(SplitDate[0], Convert.ToInt32(reader["Products"]));
                        }
                    }
                    // If there was no data
                    else
                    {
                        ProductDateAmount.Add("No data found", 0);
                    }
                }
            }

            // To build a chart:
            chartProductByDate.DataSource = ProductDateAmount;
            // Add the dates
            chartProductByDate.Series[0].XValueMember = "Key";
            // Add the customer amount
            chartProductByDate.Series[0].YValueMembers = "Value";
            chartProductByDate.DataBind();
        }

        private void ProductDateChanged(Int32 ProductAmount)
        {
            // Get the selected date
            string SelectedStart = Convert.ToString(dtpProductStart.Value);
            string SelectedEnd = Convert.ToString(dtpProductEnd.Value);

            // Split date and time
            var SelectedStartSplit = SelectedStart.Split(new char[0], StringSplitOptions.RemoveEmptyEntries);
            var SelectedEndSplit = SelectedEnd.Split(new char[0], StringSplitOptions.RemoveEmptyEntries);

            // New array with D/M/Y split
            string[] arrayStart = SelectedStartSplit[0].Split('/');
            string[] arrayEnd = SelectedEndSplit[0].Split('/');

            // Assign to the appropriate variables
            Int32 StartYear = Convert.ToInt32(arrayStart[2]);
            Int32 StartMonth = Convert.ToInt32(arrayStart[1]);
            Int32 StartDay = Convert.ToInt32(arrayStart[0]);

            Int32 EndYear = Convert.ToInt32(arrayEnd[2]);
            Int32 EndMonth = Convert.ToInt32(arrayEnd[1]);
            Int32 EndDay = Convert.ToInt32(arrayEnd[0]);

            DateTime StartingDate = new DateTime(StartYear, StartMonth, StartDay);
            DateTime EndingDate = new DateTime(EndYear, EndMonth, EndDay);

            // Convert this to a database friendly format
            string StartDate = StartingDate.ToString("MM/dd/yyyy");
            string EndDate = EndingDate.ToString("MM/dd/yyyy");

            // Methods
            MostProfitableProduct(StartDate, EndDate, ProductAmount);
            MostValuableProduct(StartDate, EndDate, ProductAmount);
            ActiveProductsByDate(StartDate, EndDate);
        }

        private void btnLoadProductData_Click(object sender, EventArgs e)
        {
            
        }

        private void btnLoadDateData_Click(object sender, EventArgs e)
        {
            // This is a hardcoded week - the lowest grade
            // Ideally this range would come from your database or elsewhere to allow the user to pick which dates they want to see
            // A good idea could be to create an empty list and then add in the week of dates you need
            List<String> dateList = new List<String>(new String[] { "06/01/2014", "07/01/2014", "08/01/2014", "09/01/2014", "10/01/2014", "11/01/2014", "12/01/2014" });

            // I need somewhere to hold the information pulled from the database, so this is an empty dictionary
            // I am using a dictionary as I can then manually set my own "key"
            // Rather than it being accessed through [0], [1], etc, I can access it via the date
            // The dictionary type is string, int - date, number of sales
            Dictionary<String, Int32> salesCount = new Dictionary<String, Int32>();

            // create a connection to the mdf file. We only need this once so it is outside my loop
            String connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            //Run this code once for each date in my list - in my case 7 times
            foreach (String date in dateList)
            {
                // Split the date down and assign it to variables for later use
                String[] arrayDate = date.Split('/');
                Int32 year = Convert.ToInt32(arrayDate[2]);
                Int32 month = Convert.ToInt32(arrayDate[1]);
                Int32 day = Convert.ToInt32(arrayDate[0]);

                DateTime myDate = new DateTime(year, month, day);

                // convert this to a database friendly format
                string dbDate = myDate.ToString("M/dd/yyyy");

                // Connect to the database using previous string
                using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
                {
                    // Open the SQL connection
                    myConnection.Open();
                    SqlCommand command = new SqlCommand("SELECT COUNT(*) as SalesNumber FROM FactTableEmpty JOIN Time " +
                        " ON FactTableEmpty.timeId = Time.id WHERE Time.date = @date;", myConnection);
                    // Add data to the previous command
                    command.Parameters.Add(new SqlParameter("@date", dbDate));

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        if (reader.HasRows) // check if there are any results
                        {
                            // Make charts visible
                           // datechart1.Visible = true;
                            datechart2.Visible = true;

                            while (reader.Read()) // there are results, so read them out
                            {
                                // add the numberof sales on this date
                                salesCount.Add(date, Int32.Parse(reader["SalesNumber"].ToString()));
                            }
                        }
                        else // there were no resuls
                        {
                            // there were 0 sales on this date
                            salesCount.Add(date, 0);
                        }
                    }
                }

            }

            // End of the foreach loop. We now have a (hopefully) filled array

            //To build a bar chart:
            datechart2.DataSource = salesCount;
            datechart2.Series[0].XValueMember = "Key";
            datechart2.Series[0].YValueMembers = "Value";
            datechart2.DataBind();

            // Or a pie chart
            //datechart1.DataSource = salesCount;
            //datechart1.Series[0].XValueMember = "Key";
            //datechart1.Series[0].YValueMembers = "Value";
            //datechart1.DataBind();
        }

        private void btnLoadData_Click(object sender, EventArgs e)
        {
            
        }

        private void MostProfitableRegion(string StartDate, string EndDate, Int32 RegionProfit)
        {
            Dictionary<String, Decimal> RegionProfits = new Dictionary<String, Decimal>();

            String connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            //connect to the database
            using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
            {
                //open the SQL connection
                myConnection.Open();
                SqlCommand command = new SqlCommand(
                "SELECT TOP (@RegionProfit) Region.region AS Region, SUM(FactTableEmpty.profit) AS Profit FROM FactTableEmpty JOIN Region ON FactTableEmpty.regionId = RegionId JOIN Product ON FactTableEmpty.productId = Product.Id JOIN Time ON FactTableEmpty.timeId = Time.id WHERE Time.date BETWEEN @StartDate AND @EndDate GROUP BY Region.region ORDER BY Profit DESC; ", myConnection);
                //add data to the previous command
                command.Parameters.Add(new SqlParameter("@StartDate", StartDate));
                command.Parameters.Add(new SqlParameter("@EndDate", EndDate));
                command.Parameters.Add(new SqlParameter("@RegionProfit", RegionProfit));

                using (SqlDataReader reader = command.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            RegionProfits.Add(Convert.ToString(reader["Region"]), Convert.ToInt32(reader["Profit"])); //converted to int to display as whole number on chart
                        }
                    }

                    else
                    {
                        RegionProfits.Add("No data found", 0);
                    }
                }
            }

            crtRegionSales.DataSource = RegionProfits;
            crtRegionSales.Series[0].XValueMember = "Key";
            crtRegionSales.Series[0].YValueMembers = "Value";
            crtRegionSales.DataBind();

            // If profit falls below or above 0 make it stand out
            foreach (DataPoint dp in crtRegionSales.Series[0].Points)
            {
                if (dp.YValues[0] < 0)
                {
                    dp.LabelForeColor = Color.Red;
                    dp.Color = Color.Red;
                }
                if (dp.YValues[0] > 0)
                {
                    dp.LabelForeColor = Color.Green;
                    dp.Color = Color.Green;
                }
            }

        }

        private void DateChanged(Int32 RegionProfit)
        {
            //get selected date
            string Start = Convert.ToString(dtpStartDate.Value);
            string End = Convert.ToString(dtpEndDate.Value);

            //split date and time
            var StartSplit = Start.Split(new char[0], StringSplitOptions.RemoveEmptyEntries);
            var EndSplit = End.Split(new char[0], StringSplitOptions.RemoveEmptyEntries);

            //new array with D/M/Y split
            string[] arrayStart = StartSplit[0].Split('/');
            string[] arrayEnd = EndSplit[0].Split('/');

            //assign to the appropriate variables
            Int32 StartYear = Convert.ToInt32(arrayStart[2]);
            Int32 StartMonth = Convert.ToInt32(arrayStart[1]);
            Int32 StartDay = Convert.ToInt32(arrayStart[0]);

            Int32 EndYear = Convert.ToInt32(arrayEnd[2]);
            Int32 EndMonth = Convert.ToInt32(arrayEnd[1]);
            Int32 EndDay = Convert.ToInt32(arrayEnd[0]);

            DateTime StartingDate = new DateTime(StartYear, StartMonth, StartDay);
            DateTime EndingDate = new DateTime(EndYear, EndMonth, EndDay);

            //convert this to database format
            string StartDate = StartingDate.ToString("MM/dd/yyyy");
            string EndDate = EndingDate.ToString("MM/dd/yyyy");

            //methods
            MostProfitableRegion(StartDate, EndDate, RegionProfit);
        }

        private void lblProfitOnDate_Click(object sender, EventArgs e)
        {

        }

        private void datechart2_Click(object sender, EventArgs e)
        {

        }

        private void tabCustomerDashboard_Click(object sender, EventArgs e)
        {

        }

        private void btnLoadData_Click_1(object sender, EventArgs e)
        {
            DateChanged(5);
        }

        private void btnLoadProductData_Click_1(object sender, EventArgs e)
        {
            ProductDateChanged(5);
        }

        private void btnLoadDateData_Click_1(object sender, EventArgs e)
        {
            //load all the data
            btnLoadDateData_Click(sender, e);
            ProductDateChanged(5);
            ProductDateChanged(5);
            DateChanged(5);
            btnLoadCustomerData_Click(sender, e);
        }

        private void chartActiveCustomersByDate_Click(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void lblFactTable_Click(object sender, EventArgs e)
        {

        }

        private void lblBuildProgress_Click(object sender, EventArgs e)
        {

        }

        private void lblGetProgress_Click(object sender, EventArgs e)
        {

        }
    }
}