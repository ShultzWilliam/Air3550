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
    /// Interaction logic for FlightHistory.xaml
    /// Place where a customer can view any flight that they have taken
    /// </summary>
    public partial class FlightHistory : Page
    {
        string flightID, identification; //initialize global variables

        public FlightHistory()
        {
            InitializeComponent();
        }
        public FlightHistory(string id) : base()
        { //define the user ID
            InitializeComponent();
            identification = id;
        }

        private void Sign_Out(object sender, RoutedEventArgs e)
        { //sign out of the application
            SignIn signIn = new SignIn();
            this.NavigationService.Navigate(signIn);
        }
        private void Submit_Click(object sender, RoutedEventArgs e)
        { //Go to the selected flight
            flightID = FlightID.Text;

            //check if the selected flight exists

            Functions functions = new Functions();
            if (functions.isNum(flightID) == true)
            { //if the flight ID exists, go to the flight
                OldFlight oldFlight = new OldFlight(identification, flightID);
                this.NavigationService.Navigate(oldFlight);
            }
            else
            { //otherwise, display an error
                Warning.Text = "Invalid Flight ID";
            }
        }
        private void Main_Menu(object sender, RoutedEventArgs e)
        { //to return to the main menu
            string userType = "Customer";
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
