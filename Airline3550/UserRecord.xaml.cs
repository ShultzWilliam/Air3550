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
    /// Interaction logic for UserRecord.xaml
    /// Where a flight manager can view a customer's record
    /// </summary>
    public partial class UserRecord : Page
    {
        string userID, Identification; //initialize global variables
        public UserRecord()
        {
            InitializeComponent();
        }
        public UserRecord(string identification, string id)
        { //define user and flight ID
            InitializeComponent();
            userID = id;
            Identification = identification;
        }
        private void Window_Loaded(object sender, RoutedEventArgs e)
        { //when the window is loaded, load in the flight info
            User_Name.Text = userID;
            CreditCard.Text = "LOL";
            Paid.Text = "Too Much";
        }
        private void Sign_Out(object sender, RoutedEventArgs e)
        { //to sign out
            SignIn signIn = new SignIn(); //create a new main menu and go to it
            this.NavigationService.Navigate(signIn);
        }
        private void Main_Menu(object sender, RoutedEventArgs e)
        { //to return to the main menu
            MainMenuFlightManager mainMenu = new MainMenuFlightManager(Identification); //create a new main menu and go to it
            this.NavigationService.Navigate(mainMenu);
        }

        private void Print(object sender, RoutedEventArgs e)
        { //to print the customer record
        }
    }
}
