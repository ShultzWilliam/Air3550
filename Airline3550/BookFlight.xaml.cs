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
    /// Interaction logic for BookFlight.xaml
    /// Allows customers to book a flight
    /// </summary>
    public partial class BookFlight : Page
    {
        string flightID, Identification; //initialize global variables
        Functions functions = new Functions(); //get the necessary functions
        public BookFlight()
        {
            InitializeComponent();
        }
        public BookFlight(string identification, string id)
        { //define the flight and user ID
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

            //Birth.SelectedDate = DateTime.FromOADate(xlWorksheet.Cells[IDRow, 11].Value2); //Set the birth date in the text box
            //Credit.Text = xlWorksheet.Cells[IDRow, 12].Value2.ToString(); //Set the credit card number in the text box
            FlightID.Text = flightID;
            Origin.Text = functions.getAirport(xlWorksheet.Cells[IDRow, 5].Value2.ToString());
            Destination.Text = functions.getAirport(xlWorksheet.Cells[IDRow, 6].Value2.ToString());
            Departure_Date.Text = DateTime.FromOADate(xlWorksheet.Cells[IDRow, 7].Value2);
            Departure_Time.Text = xlWorksheet.Cells[IDRow, 8].Value2.ToString();
            Departure_Terminal.Text = xlWorksheet.Cells[IDRow, 9].Value2.ToString();
            Arrival_Date.Text = DateTime.FromOADate(xlWorksheet.Cells[IDRow, 10].Value2);
            Arrival_Time.Text = xlWorksheet.Cells[IDRow, 11].Value2.ToString();
            Arrival_Terminal.Text = xlWorksheet.Cells[IDRow, 12].Value2.ToString();
            Price.Text = xlWorksheet.Cells[IDRow, 17].Value2.ToString();
            Plane.Text = xlWorksheet.Cells[IDRow, 14].Value2.ToString();
            xlWorkbook.Close(true);
        }
        private void Book_Click(object sender, RoutedEventArgs e)
        { //to book the flight
            if (((bool)Credit.IsChecked == true && (bool)CreditCard.IsChecked == false && (bool)Points.IsChecked == false) ||
                ((bool)Credit.IsChecked == false && (bool)CreditCard.IsChecked == true && (bool)Points.IsChecked == false) ||
                ((bool)Credit.IsChecked == false && (bool)CreditCard.IsChecked == false && (bool)Points.IsChecked == true))
            { //check that only one check box is checked
                MainMenuCustomer mainMenu = new MainMenuCustomer(Identification); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
            else
            { //otherwise, display a warning
                Warning.Text = "Cannot select more than one payment method";
            }

        }
        private void Sign_Out(object sender, RoutedEventArgs e)
        { //to sign out
            SignIn signIn = new SignIn(); //create a new main menu and go to it
            this.NavigationService.Navigate(signIn);
        }

        private void Main_Menu(object sender, RoutedEventArgs e)
        { //to return to the main menu
            string userType = "Customer";
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
