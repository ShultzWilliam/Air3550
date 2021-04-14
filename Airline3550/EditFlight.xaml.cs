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
    /// Interaction logic for EditFlight.xaml
    /// Place where a load engineer can edit a existing flight
    /// </summary>
    public partial class EditFlight : Page
    {
        string Identification, flightID; //initialize global variables
        public EditFlight()
        {
            InitializeComponent();
        }
        public EditFlight(string identification, string FlightID)
        { //define the flight and user ID
            InitializeComponent();
            Identification = identification;
            flightID = FlightID;
        }

        private void Sign_Out(object sender, RoutedEventArgs e)
        { //to sign out
            SignIn signIn = new SignIn(); //create a new main menu and go to it
            this.NavigationService.Navigate(signIn);
        }
        private void Main_Menu(object sender, RoutedEventArgs e)
        { //to return to the main menu
            MainMenuLoadEngineer mainMenu = new MainMenuLoadEngineer(Identification); //create a new main menu and go to it
            this.NavigationService.Navigate(mainMenu);
        }
        private void Submit_Click(object sender, RoutedEventArgs e)
        { //to submit flight changes

            //function to save the changes to the database

            Functions functions = new Functions();
            if (functions.isNum(Price.Text) && functions.isTime(Arrival_Time.Text) && functions.isTime(Departure_Time.Text))
            { //if the inputs are correct
                MainMenuLoadEngineer mainMenu = new MainMenuLoadEngineer(Identification); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
            else if (!(functions.isNum(Price.Text)))
            { //if the price input is incorrect
                Warning.Text = "Incorrect price input";
            }
            else
            { //if the time input is incorrect
                Warning.Text = "Incorrect time input";
            }
        }
        private void Calculate_Click(object sender, RoutedEventArgs e)
        { //to calculate the price of the flight
            Price.Text = "95";
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        { //to load in the flight's information
            Origin.SelectedIndex = 1; //placeholders, replace once we have a database
            Destination.SelectedIndex = 0;
            Departure_Date.SelectedDate = DateTime.Today;
            Departure_Time.Text = "Departure Time";
            Departure_Terminal.Text = "Departure Terminal";
            Arrival_Date.SelectedDate = DateTime.Today;
            Arrival_Time.Text = "Departure Time";
            Arrival_Terminal.Text = "Departure Terminal";
            Price.Text = "Too Expensive";
        }
    }
}
