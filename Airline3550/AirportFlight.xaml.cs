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
    /// Interaction logic for AirportFlight.xaml
    /// Displays the flight details for the accountant user class
    /// </summary>
    public partial class AirportFlight : Page
    {
        string flightID, Identification; //strings that will be useful across the page

        public AirportFlight()
        {
            InitializeComponent();
        }
        public AirportFlight(string identification, string id)
        {
            //define the flight and user ID
            InitializeComponent();
            flightID = id;
            Identification = identification;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        { //when the window is loaded, load in the flight info

            //load in the data from the database (for now I have placeholder data)

            FlightID.Text = flightID;
            Capacity.Text = Convert.ToString(80) + "/" + Convert.ToString(100);
            Profit.Text = Convert.ToString(10000);
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


        private void Print(object sender, RoutedEventArgs e)
        { //Print the flight details
        }

        private void Main_Menu(object sender, RoutedEventArgs e)
        { //to return to the main menu
            MainMenuAccountant mainMenu = new MainMenuAccountant(Identification); //create a new main menu and go to it
            this.NavigationService.Navigate(mainMenu);
            
        }
    }
}
