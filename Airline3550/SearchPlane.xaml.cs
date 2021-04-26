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
    /// Interaction logic for SearchPlane.xaml
    /// Place where marketing manager can search for a flight to select a plane for
    /// </summary>
    public partial class SearchPlane : Page
    {
        string flightID, identification; //initialize global variables
        public SearchPlane()
        {
            InitializeComponent();
        }
        public SearchPlane(string id) : base()
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
            int numRows = functions.getRows(2);
            if (functions.isFlight(flightID) == true)
            { //if the flight ID exists, go to the flight

                SchedulePlane schedulePlane = new SchedulePlane(identification, flightID);
                this.NavigationService.Navigate(schedulePlane);
            }
            else
            { //otherwise, display an error
                Warning.Text = "Invalid Flight ID";
            }
        }
        private void Main_Menu(object sender, RoutedEventArgs e)
        { //to return to the main menu

            MainMenuMarketingManager mainMenu = new MainMenuMarketingManager(identification); //create a new main menu and go to it
            this.NavigationService.Navigate(mainMenu);

        }
    }
}
