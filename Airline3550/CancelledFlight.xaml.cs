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
    /// Interaction logic for CancelledFlight.xaml
    /// </summary>
    public partial class CancelledFlight : Page
    {
        string flightID, Identification; //initialize global variables
        Functions functions = new Functions(); //get the necessary functions
        public CancelledFlight()
        {
            InitializeComponent();
        }

        public CancelledFlight(string identification, string id)
        { //define the user and flight IDs
            InitializeComponent();
            flightID = id;
            Identification = identification;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        { //when the window is loaded, load in the flight info

            //get the necessary excel variables
            Excel.Workbook xlWorkbook = functions.database_connect();
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[2];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = functions.getRows(2);
            int IDRow = functions.getIDRow(flightID, 2);
            Excel._Worksheet xlWorksheet1 = xlWorkbook.Sheets[1];
            Excel.Range xlRange1 = xlWorksheet1.UsedRange;
            int rowCount1 = functions.getRows(1);
            int userIDRow = functions.getIDRow(Identification, 1); //get the ID Rows for the flight and user IDs
            string userFlights = xlRange1.Cells[userIDRow, 23].Value2.ToString(); //get the original userFlights
            string paymentMethods = xlRange1.Cells[userIDRow, 24].Value2.ToString();
            string paidWith = "";
            string[] userFlightsArray = userFlights.Split(' '); //create arrays to find the indeces of the flight and payment method
            string[] paymentMethodsArray = paymentMethods.Split(' ');
            for (int i = 0; i < userFlightsArray.Length; i++)
            { //go through the array until we find the user's flight index to find the payment method and price
                if (userFlightsArray[i] == flightID)
                {
                    paidWith = paymentMethodsArray[i];
                    break;
                }

            }

            //load in the information to the page
            FlightID.Text = flightID;
            Origin.Text = functions.getAirport(xlRange.Cells[IDRow, 5].Value2.ToString());
            Destination.Text = functions.getAirport(xlRange.Cells[IDRow, 6].Value2.ToString());
            Departure_Date.Text = (DateTime.FromOADate(xlRange.Cells[IDRow, 7].Value2)).ToString("MM/dd/yyyy");
            Departure_Time.Text = (DateTime.FromOADate(xlRange.Cells[IDRow, 8].Value2)).ToString("h:mm tt");
            Departure_Terminal.Text = xlRange.Cells[IDRow, 9].Value2.ToString();
            Arrival_Date.Text = (DateTime.FromOADate(xlRange.Cells[IDRow, 10].Value2)).ToString("MM/dd/yyyy");
            Arrival_Time.Text = (DateTime.FromOADate(xlRange.Cells[IDRow, 11].Value2)).ToString("h:mm tt");
            Arrival_Terminal.Text = xlRange.Cells[IDRow, 12].Value2.ToString();
            Price.Text = "$" + xlRange.Cells[IDRow, 17].Value2.ToString();
            Plane.Text = xlRange.Cells[IDRow, 14].Value2.ToString();
            Paid.Text = paidWith;
            xlWorkbook.Close(true);
        }

        private void Sign_Out(object sender, RoutedEventArgs e)
        { //to sign out
            SignIn signIn = new SignIn(); //create a new main menu and go to it
            this.NavigationService.Navigate(signIn);
        }

        private void Main_Menu(object sender, RoutedEventArgs e)
        { //to return to the main menu
            int IDrow = functions.getIDRow(Identification, 1);
            string userType = functions.getUserType(IDrow);
            if (userType == "Customer")
            {
                MainMenuCustomer mainMenu = new MainMenuCustomer(Identification); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
            else if (userType == "Load Engineer")
            {
                MainMenuLoadEngineer mainMenu = new MainMenuLoadEngineer(Identification); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
            else if (userType == "Accountant")
            {
                MainMenuAccountant mainMenu = new MainMenuAccountant(Identification); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
            else if (userType == "Marketing Manager")
            {
                MainMenuMarketingManager mainMenu = new MainMenuMarketingManager(Identification); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
            else if (userType == "Flight Manager")
            {
                MainMenuFlightManager mainMenu = new MainMenuFlightManager(Identification); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
        }
    }
}
