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
    /// Interaction logic for Profile.xaml
    /// Where a user can view and edit their profile
    /// </summary>
    public partial class Profile : Page
    {
        string Identification; //initialize the user ID
        public Profile()
        {
            InitializeComponent();
        }
        public Profile(string id) : base()
        { //Load in the user ID
            InitializeComponent();
            Identification = id; //set the global variable to the passed in ID
        }

        private void Name_Loaded(object sender, RoutedEventArgs e)
        { //Load in the name
            User_Name.Text = "Name"; //Set the name in the text box
        }

        private void Address_Loaded(object sender, RoutedEventArgs e)
        { //Load in the address
            Address.Text = "Address"; //Set the address in the text box
        }

        private void Phone_Loaded(object sender, RoutedEventArgs e)
        { //Load in the phone
            Phone.Text = "Phone"; //Set the phone in the text box
        }

        private void City_Loaded(object sender, RoutedEventArgs e)
        { //Load in the city
            City.Text = "City"; //Set the city in the text box
        }
        private void State_Loaded(object sender, RoutedEventArgs e)
        { //Load in the state
            State.Text = "State"; //Set the state in the text box
        }
        private void ZIP_Loaded(object sender, RoutedEventArgs e)
        { //Load in the zip code
            Zip.Text = "Zip Code"; //Set the Zip Code in the text box
        }
        private void Email_Loaded(object sender, RoutedEventArgs e)
        { //Load in the email
            Email.Text = "Email"; //Set the email in the text box
        }
        private void Birth_Loaded(object sender, RoutedEventArgs e)
        { //Load in the birth date
            Birth.SelectedDate = DateTime.Today; //Set the birth date in the text box
        }
        private void Credit_Loaded(object sender, RoutedEventArgs e)
        { //Load in the credit card number
            Credit.Text = "Credit card number"; //Set the credit card number in the text box
        }
        private void Exp_Loaded(object sender, RoutedEventArgs e)
        { //Load in the credit card experation date
            Expiration.Text = "Exp Date"; //Set the expiration date in the text box
        }
        private void CSV_Loaded(object sender, RoutedEventArgs e)
        { //Load in the Credit card CSV number
            CSV.Text = "Credit Card CSV"; //Set the Credit Card CSV in the text box
        }
        private void Sign_Out(object sender, RoutedEventArgs e)
        { //sign out of the application
            SignIn signIn = new SignIn();
            this.NavigationService.Navigate(signIn);
        }
        private void Submit_Click(object sender, RoutedEventArgs e)
        { //to return to the main menu
            string userType = "Customer";
            if (userType == "Customer")
            {
                MainMenuCustomer mainMenu = new MainMenuCustomer(Identification); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
            else if (userType == "Load Engineer")
            {
                MainMenuLoadEngineer mainMenu = new MainMenuLoadEngineer(Identification); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
            else if (userType == "Accountant")
            {
                MainMenuAccountant mainMenu = new MainMenuAccountant(Identification); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
            else if (userType == "Marketing Manager")
            {
                MainMenuMarketingManager mainMenu = new MainMenuMarketingManager(Identification); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
            else if (userType == "Flight Manager")
            {
                MainMenuFlightManager mainMenu = new MainMenuFlightManager(Identification); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
        }
    }
}
