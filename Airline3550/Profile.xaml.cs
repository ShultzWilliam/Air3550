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
using System.Security.Cryptography;


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

        private void Window_Loaded(object sender, RoutedEventArgs e)
        { //Load in the name
            User_Name.Text = "Name"; //Set the name in the text box
            Address.Text = "Address"; //Set the address in the text box
            Phone.Text = "Phone"; //Set the phone in the text box
            City.Text = "City"; //Set the city in the text box
            State.Text = "State"; //Set the state in the text box
            Zip.Text = "Zip Code"; //Set the Zip Code in the text box
            Email.Text = "Email"; //Set the email in the text box
            Birth.SelectedDate = DateTime.Today; //Set the birth date in the text box
            CreditCard.Text = "Credit card number"; //Set the credit card number in the text box
            Expiration.Text = "Exp Date"; //Set the expiration date in the text box
            CSV.Text = "Credit Card CSV"; //Set the Credit Card CSV in the text box
            Password.Text = "Password";
            Credit.Text = "Credit";
            Points.Text = "Points";
        }
        private void Sign_Out(object sender, RoutedEventArgs e)
        { //sign out of the application
            SignIn signIn = new SignIn();
            this.NavigationService.Navigate(signIn);
        }
        private void Submit_Click(object sender, RoutedEventArgs e)
        { //to return to the main menu

            //save the information to the database

            byte[] password; //to save the password
            using (SHA512 shaM = new SHA512Managed())
            { //save the password as a SHA512 hash
                password = shaM.ComputeHash(Encoding.UTF8.GetBytes(Password.Text));
            }

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
