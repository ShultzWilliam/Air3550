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
    /// Interaction logic for Airport.xaml
    /// Displays the history of the entire airport for the accountant user class and
    /// allows them to pick an flight for more information
    /// </summary>
    public partial class Airport : Page
    {
        string airportID, identification, flightID;

        public Airport()
        {
            InitializeComponent();
        }
        public Airport(string id, string AirportID) : base()
        { //initialize the airport ID and user ID
            InitializeComponent();
            identification = id;
            airportID = AirportID;
        }

        private void Sign_Out(object sender, RoutedEventArgs e)
        { //sign out of the application
            SignIn signIn = new SignIn();
            this.NavigationService.Navigate(signIn);
        }
        private void Submit_Click(object sender, RoutedEventArgs e)
        { //Go to the airport selected
            flightID = Flight.Text;
            Functions functions = new Functions();
            if (functions.isNum(flightID) == true)
            { //check the the flight ID is a number (later we'll have to check if the flight ID exists
                AirportFlight airportFlight = new AirportFlight(identification, flightID);
                this.NavigationService.Navigate(airportFlight);
            }
            else
            {
                Warning.Text = "Incorrect Flight ID";
            }
        }
        private void Main_Menu(object sender, RoutedEventArgs e)
        { //to return to the main menu
            MainMenuAccountant mainMenu = new MainMenuAccountant(identification); //create a new main menu and go to it
            this.NavigationService.Navigate(mainMenu);
        }
    }
}
