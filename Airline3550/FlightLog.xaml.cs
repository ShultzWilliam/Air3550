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
    /// Interaction logic for FlightLog.xaml
    /// Place where a flight manager can look for flights to print a manifest for
    /// </summary>
    public partial class FlightLog : Page
    {
        string flightID, identification; //initialize global variables

        public FlightLog()
        {
            InitializeComponent();
        }
        public FlightLog(string id) : base()
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
        { //Go to the flight manifest
            flightID = FlightID.Text;

            //check if the flightID exists

            Functions functions = new Functions();
            if (functions.isNum(flightID) == true)
            { //if the flight ID exists, go to the flightManifest
                FlightManifest flightManifest = new FlightManifest(identification, flightID);
                this.NavigationService.Navigate(flightManifest);
            }
            else
            { //otherwise, display an error
                Warning.Text = "Invalid Flight ID";
            }
        }
        private void Main_Menu(object sender, RoutedEventArgs e)
        { //to return to the main menu
            MainMenuFlightManager mainMenu = new MainMenuFlightManager(identification); //create a new main menu and go to it
            this.NavigationService.Navigate(mainMenu);
        }
    }
}
