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
    /// Interaction logic for MainMenuFlightManager.xaml
    /// Main menu for the flight manager user type
    /// </summary>
    public partial class MainMenuFlightManager : Page
    {

        public string Identification;
        public MainMenuFlightManager()
        {
            InitializeComponent();
        }

        public MainMenuFlightManager(string id) : base()
        { //Load in the user ID
            InitializeComponent();
            Identification = id; //set the global variable to the passed in ID
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        { //load in the user information
            Functions functions = new Functions();
            int IDRow = functions.getIDRow(Identification, 1); //get the ID column for the user
            User.Text = functions.getName(IDRow); //Print the passed in ID
        }

        private void Sign_Out(object sender, RoutedEventArgs e)
        { //sign out of the application
            SignIn signIn = new SignIn();
            this.NavigationService.Navigate(signIn);
        }
        private void Edit_Profile(object sender, RoutedEventArgs e)
        { //edit your profile
            Profile profile = new Profile(Identification);
            this.NavigationService.Navigate(profile);
        }

        private void Book_Flight(object sender, RoutedEventArgs e)
        { //book a flight
            SearchFlight searchFlight = new SearchFlight(Identification);
            this.NavigationService.Navigate(searchFlight);
        }


        private void My_Flights(object sender, RoutedEventArgs e)
        { //Go to scheduled flights
            MyFlights myFlights = new MyFlights(Identification);
            this.NavigationService.Navigate(myFlights);
        }
        private void My_History_Click(object sender, RoutedEventArgs e)
        { //Go to taken flights
            FlightHistory myHistory = new FlightHistory(Identification);
            this.NavigationService.Navigate(myHistory);
        }

        private void Flight_Manifest(object sender, RoutedEventArgs e)
        {// go to the flightManifest screen
            FlightManifest flightManifest = new FlightManifest(Identification);
            this.NavigationService.Navigate(flightManifest);

        }
    }
}
