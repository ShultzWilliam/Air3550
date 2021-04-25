using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace Air3550
{
    /// <summary>
    /// Interaction logic for FlightHistory.xaml
    /// Place where a customer can view any flight that they have taken
    /// </summary>
    public partial class FlightHistory : Page
    {
        string flightID, identification; //initialize global variables
        int userIDRow;
        Functions functions = new Functions();

        public FlightHistory()
        {
            InitializeComponent();
        }
        public class flightItem
        { //class used to insert flights into the data grid
            public string ID { get; set; }
            public string Origin { get; set; }
            public string Destination { get; set; }
            public string Departure { get; set; }
            public string Arrival { get; set; }
            public string Price { get; set; }
        }
        public FlightHistory(string id) : base()
        { //define the user ID
            InitializeComponent();
            identification = id;
            
        }
        private void Window_Loaded(object sender, RoutedEventArgs e)
        { //Load in the data grid
            Flights.Items.Clear();
            //define the excel variables
            //Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = functions.database_connect();
            Excel._Worksheet xlWorksheet1 = xlWorkbook.Sheets[1];
            Excel.Range xlRange1 = xlWorksheet1.UsedRange;
            int rowCount1 = functions.getRows(1);
            Excel._Worksheet xlWorksheet2 = xlWorkbook.Sheets[2];
            Excel.Range xlRange2 = xlWorksheet2.UsedRange;
            int rowCount2 = functions.getRows(2);
            userIDRow = functions.getIDRow(identification, 1); //get the ID Rows for the flight and user IDs
            int[] flightArray = new int[rowCount2];
            int numOfFlights = 0;
            if (functions.isEmpty(1, userIDRow, 22) == false)
            { //if the flights cell isn't empty
                string flights = xlRange1.Cells[userIDRow, 22].Value2.ToString(); //get the user's flights

                for (int i = 1; i <= rowCount2; i++)
                {
                    if (flights.Contains(xlRange2.Cells[i, 1].Value2.ToString()))
                    { //if the user is scheduled for the flight
                        flightArray[numOfFlights] = i; //save the index
                        numOfFlights++; //increment the number of flights
                    }
                }

                if (numOfFlights == 0)
                { //if num of flights = 0, then we didn't find any flights
                    Warning.Text = "No flights found";
                }
                else
                { //otherwise
                    for (int i = 0; i < numOfFlights; i++)
                    { //for each flight we found
                        var item = new flightItem
                        {
                            ID = xlRange2.Cells[flightArray[i], 1].Value2.ToString(),
                            Origin = functions.getAirport(xlRange2.Cells[flightArray[i], 5].Value2.ToString()),
                            Destination = functions.getAirport(xlRange2.Cells[flightArray[i], 6].Value2.ToString()),
                            Departure = DateTime.FromOADate(xlRange2.Cells[flightArray[i], 7].Value2).ToString("MM/dd/yyyy") + " " + DateTime.FromOADate(xlRange2.Cells[flightArray[i], 8].Value2).ToString("h:mm tt"),
                            Arrival = DateTime.FromOADate(xlRange2.Cells[flightArray[i], 10].Value2).ToString("MM/dd/yyyy") + " " + DateTime.FromOADate(xlRange2.Cells[flightArray[i], 11].Value2).ToString("h:mm tt"),
                            Price = "$" + xlRange2.Cells[flightArray[i], 17].Value2.ToString()
                        }; //create a new flight item to insert into the data grid
                        Flights.Items.Add(item);
                    }
                }
            }
            else
            { //otherwise
                Warning.Text = "You have not been on any flights";
            }
            xlWorkbook.Close();
        }

            private void Sign_Out(object sender, RoutedEventArgs e)
        { //sign out of the application
            SignIn signIn = new SignIn();
            this.NavigationService.Navigate(signIn);
        }
        private void Submit_Click(object sender, RoutedEventArgs e)
        { //Go to the selected flight
            flightID = FlightID.Text;

            //create the excel variables
            Excel.Workbook xlWorkbook = functions.database_connect();
            Excel._Worksheet xlWorksheet1 = xlWorkbook.Sheets[1];
            Excel.Range xlRange1 = xlWorksheet1.UsedRange;
            int rowCount1 = functions.getRows(1);
            string myFlights = xlRange1.Cells[userIDRow, 22].Value2.ToString(); //get the user's flights


            if ((functions.isNum(flightID) == true) && (functions.isFlight(flightID) == true) && (myFlights.Contains(flightID)))
            { //if the flight ID exists, go to the flight
                xlWorkbook.Close();
                OldFlight oldFlight = new OldFlight(identification, flightID);
                this.NavigationService.Navigate(oldFlight);
            }
            else
            { //otherwise, display an error
                xlWorkbook.Close();
                Warning.Text = "Invalid Flight ID";
            }
        }
        private void Main_Menu(object sender, RoutedEventArgs e)
        { //to return to the main menu
            int IDrow = functions.getIDRow(identification, 1);
            string userType = functions.getUserType(IDrow);
            if (userType == "Customer")
            {
                MainMenuCustomer mainMenu = new MainMenuCustomer(identification); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
            else if (userType == "Load Engineer")
            {
                MainMenuLoadEngineer mainMenu = new MainMenuLoadEngineer(identification); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
            else if (userType == "Accountant")
            {
                MainMenuAccountant mainMenu = new MainMenuAccountant(identification); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
            else if (userType == "Marketing Manager")
            {
                MainMenuMarketingManager mainMenu = new MainMenuMarketingManager(identification); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
            else if (userType == "Flight Manager")
            {
                MainMenuFlightManager mainMenu = new MainMenuFlightManager(identification); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
        }
    }
}
