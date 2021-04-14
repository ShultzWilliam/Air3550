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
    /// Interaction logic for SchedulePlane.xaml
    /// Place where a Marketing Manager can schedule planes for flights
    /// </summary>
    public partial class SchedulePlane : Page
    {
        string flightID, Identification; //initialize global variables

        public SchedulePlane()
        {
            InitializeComponent();
        }
        public SchedulePlane(string identification, string id)
        { //define the flight and user ID
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
            Departure_Date.Text = "4/29/2021";
            Departure_Time.Text = "5:00 PM";
            Departure_Terminal.Text = "B";
            Price.Text = "$" + Convert.ToString(95);
        }
        private void Schedule_Click(object sender, RoutedEventArgs e)
        { //to book the flight

            //schedule the flight

            MainMenuMarketingManager mainMenu = new MainMenuMarketingManager(Identification); //create a new main menu and go to it
            this.NavigationService.Navigate(mainMenu);
        }
        private void Sign_Out(object sender, RoutedEventArgs e)
        { //to sign out
            SignIn signIn = new SignIn(); //create a new main menu and go to it
            this.NavigationService.Navigate(signIn);
        }

        private void Main_Menu(object sender, RoutedEventArgs e)
        { //to return to the main menu

            MainMenuMarketingManager mainMenu = new MainMenuMarketingManager(Identification); //create a new main menu and go to it
            this.NavigationService.Navigate(mainMenu);
            
        }
    }
}
