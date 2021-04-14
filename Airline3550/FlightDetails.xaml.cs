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

namespace Air3550
{
    /// <summary>
    /// Interaction logic for FlightDetails.xaml
    /// Place where a customer can view the details of a yet to be taken flight
    /// </summary>
    public partial class FlightDetails : Page
    {
        string flightID, Identification; //initialize global variables

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

            //load in the data from the database (for now I have placeholder data)

            FlightID.Text = flightID;
            Origin.Text = "Origin";
            Destination.Text = "Destination";
            Departure_Date.Text = "4/29/2021";
            Departure_Time.Text = "5:00 AM";
            Departure_Terminal.Text = "A";
            Arrival_Date.Text = "4/29/2021";
            Arrival_Time.Text = "5:00 PM";
            Arrival_Terminal.Text = "B";
            Price.Text = "$" + Convert.ToString(95);
            Plane.Text = "737";
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
