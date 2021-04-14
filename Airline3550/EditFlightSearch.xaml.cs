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
    /// Interaction logic for EditFlightSearch.xaml
    /// Area where a load engineer can search for flights to edit
    /// </summary>
    public partial class EditFlightSearch : Page
    {
        string flightID, identification; //initialize global variables
        public EditFlightSearch()
        {
            InitializeComponent();
        }
        public EditFlightSearch(string id) : base()
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
        { //Go back to the main menu
            flightID = FlightID.Text;

            //check if the flightID exists

            Functions functions = new Functions();
            if (functions.isNum(flightID) == true)
            { //if the flight ID exists, go to the flight
                EditFlight editFlight = new EditFlight(identification, flightID);
                this.NavigationService.Navigate(editFlight);
            }
            else
            { //otherwise, display an error
                Warning.Text = "Invalid Flight ID";
            }
        }
        private void Main_Menu(object sender, RoutedEventArgs e)
        { //to return to the main menu
            MainMenuLoadEngineer mainMenu = new MainMenuLoadEngineer(identification); //create a new main menu and go to it
            this.NavigationService.Navigate(mainMenu);
        }
    }
}
