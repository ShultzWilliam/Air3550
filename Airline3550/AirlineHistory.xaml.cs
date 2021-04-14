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
    /// Interaction logic for AirlineHistory.xaml
    /// Displays the history of the entire airline for the accountant user class and
    /// allows them to pick an airport for more information
    /// </summary>
    public partial class AirlineHistory : Page
    {
        string airportID, identification;

        public AirlineHistory()
        {
            InitializeComponent();
        }
        public AirlineHistory(string id) : base()
        { //initialize the user ID
            InitializeComponent();
            identification = id;
        }

        private void Sign_Out(object sender, RoutedEventArgs e)
        { //sign out of the application
            SignIn signIn = new SignIn();
            this.NavigationService.Navigate(signIn);
        }
        private void Submit_Click(object sender, RoutedEventArgs e)
        { //Go to the airport selected
            airportID = Airport.Text;
            Airport airport = new Airport(identification, airportID);
            this.NavigationService.Navigate(airport);
        }
        private void Main_Menu(object sender, RoutedEventArgs e)
        { //to return to the main menu
            
            MainMenuAccountant mainMenu = new MainMenuAccountant(identification); //create a new main menu and go to it
            this.NavigationService.Navigate(mainMenu);
            
        }
    }
}
