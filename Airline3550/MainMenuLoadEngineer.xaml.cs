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
    /// Interaction logic for MainMenuLoadEngineer.xaml
    /// Main menu for the load engineer user type
    /// </summary>
    public partial class MainMenuLoadEngineer : Page
    {

        public string Identification; //define global ID
        public MainMenuLoadEngineer()
        {
            InitializeComponent();
        }

        public MainMenuLoadEngineer(string id) : base()
        { //Load in the user ID
            InitializeComponent();
            Identification = id; //set the global variable to the passed in ID
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        { //load the user information
            Functions functions = new Functions();
            int IDRow = functions.getIDRow(Identification, 1); //get the ID column for the user
            User.Text = functions.getName(IDRow); //Print the passed in ID
        }

        private void Sign_Out(object sender, RoutedEventArgs e)
        { //sign out of the application
            SignIn signIn = new SignIn();
            this.NavigationService.Navigate(signIn);
        }

        private void Create_Flight(object sender, RoutedEventArgs e)
        { //To schedule a new flight
            CreateFlight createFlight = new CreateFlight(Identification);
            this.NavigationService.Navigate(createFlight);

        }

        private void Edit_Flights(object sender, RoutedEventArgs e)
        { //To schedule a new flight
            EditFlightSearch editFlightSearch = new EditFlightSearch(Identification);
            this.NavigationService.Navigate(editFlightSearch);
        }

        private void Profile(object sender, RoutedEventArgs e)
        { //edit your profile
            Profile profile = new Profile(Identification);
            this.NavigationService.Navigate(profile);
        }

    }
}
