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
    /// Interaction logic for FlightDetails.xaml
    /// Place where a customer can view the details of a yet to be taken flight
    /// </summary>
    public partial class FlightDetails : Page
    {
        string flightID, Identification; //initialize global variables
        Functions functions = new Functions(); //get the necessary functions

        public FlightDetails()
        {
            InitializeComponent();
        }

        public FlightDetails(string identification, string id)
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

            //load in the information to the page
            FlightID.Text = flightID;
            Origin.Text = functions.getAirport(xlWorksheet.Cells[IDRow, 5].Value2.ToString());
            Destination.Text = functions.getAirport(xlWorksheet.Cells[IDRow, 6].Value2.ToString());
            Departure_Date.Text = (DateTime.FromOADate(xlRange.Cells[IDRow, 7].Value2)).ToString("MM/dd/yyyy");
            Departure_Time.Text = (DateTime.FromOADate(xlWorksheet.Cells[IDRow, 8].Value2)).ToString("h:mm tt");
            Departure_Terminal.Text = xlWorksheet.Cells[IDRow, 9].Value2.ToString();
            Arrival_Date.Text = (DateTime.FromOADate(xlRange.Cells[IDRow, 10].Value2)).ToString("MM/dd/yyyy");
            Arrival_Time.Text = (DateTime.FromOADate(xlWorksheet.Cells[IDRow, 11].Value2)).ToString("h:mm tt");
            Arrival_Terminal.Text = xlWorksheet.Cells[IDRow, 12].Value2.ToString();
            Price.Text = "$" + xlWorksheet.Cells[IDRow, 17].Value2.ToString();
            Plane.Text = xlWorksheet.Cells[IDRow, 14].Value2.ToString();
            xlWorkbook.Close(true);
        }
        private void Cancel_Flight(object sender, RoutedEventArgs e)
        { //to book the flight
            //cancel the flight and give the customer credit or their points back
            MyFlights myFlights = new MyFlights(Identification); //create a new main menu and go to it
            this.NavigationService.Navigate(myFlights);
        }
        private void Sign_Out(object sender, RoutedEventArgs e)
        { //to sign out
            SignIn signIn = new SignIn(); //create a new main menu and go to it
            this.NavigationService.Navigate(signIn);
        }


        private void Print_Pass(object sender, RoutedEventArgs e)
        { //Print the boarding pass
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
