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
    /// Interaction logic for FlightManifest.xaml
    /// Place where a flight manager can print a flight manifest or search for customer records
    /// </summary>
    public partial class FlightManifest : Page
    {
        string flightID, identification, userID; //initialize global variables

        public FlightManifest()
        {
            InitializeComponent();
        }
        public FlightManifest(string id, string ID) : base()
        { //define the user and flight IDs
            InitializeComponent();
            identification = id;
            flightID = ID;
        }

        private void Sign_Out(object sender, RoutedEventArgs e)
        { //sign out of the application
            SignIn signIn = new SignIn();
            this.NavigationService.Navigate(signIn);
        }
        private void Submit_Click(object sender, RoutedEventArgs e)
        { //Go back to the main menu
            userID = UserID.Text;
            //check if the flightID exists

            Functions functions = new Functions();
            if (functions.isNum(userID) == true)
            { //if the user ID exists, go to the user
                UserRecord flightManifest = new UserRecord(identification, userID);
                this.NavigationService.Navigate(flightManifest);
            }
            else
            { //otherwise, display an error
                Warning.Text = "Invalid user ID";
            }
        }
        private void Main_Menu(object sender, RoutedEventArgs e)
        { //to return to the main menu
            MainMenuFlightManager mainMenu = new MainMenuFlightManager(identification); //create a new main menu and go to it
            this.NavigationService.Navigate(mainMenu);
        }
        private void Print(object sender, RoutedEventArgs e)
        { //to print the customer record
        }
    }
}
