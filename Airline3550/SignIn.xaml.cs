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
    /// Interaction logic for SignIn.xaml
    /// Where a user signs in
    /// </summary>
    public partial class SignIn : Page
    {
        public SignIn()
        {
            InitializeComponent();
        }

        private void Submit_Click(object sender, RoutedEventArgs e)
        { //when going to the main menu
            String id, password; //initialize strings for the ID and Password
            id = ID.Text; //set the strings
            password = Password.Text;
            string userType = password; //use the password to check each user type for now
            byte[] passwordHash; //to save the password
            using (SHA512 shaM = new SHA512Managed())
            { //save the password as a SHA512 hash
                passwordHash = shaM.ComputeHash(Encoding.UTF8.GetBytes(Password.Text));
            }
            if (userType == "Customer")
            {
                MainMenuCustomer mainMenu = new MainMenuCustomer(id); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
            else if (userType == "Load Engineer")
            {
                MainMenuLoadEngineer mainMenu = new MainMenuLoadEngineer(id); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
            else if (userType == "Accountant")
            {
                MainMenuAccountant mainMenu = new MainMenuAccountant(id); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
            else if (userType == "Marketing Manager")
            {
                MainMenuMarketingManager mainMenu = new MainMenuMarketingManager(id); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
            else if (userType == "Flight Manager")
            {
                MainMenuFlightManager mainMenu = new MainMenuFlightManager(id); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
            //MainMenuCustomer mainMenu = new MainMenuCustomer(id); //create a new main menu and go to it
            //this.NavigationService.Navigate(mainMenu);

        }

        private void Create_Click(object sender, RoutedEventArgs e)
        { //when going to create account
            ID.Clear(); //clear the ID and password
            Password.Clear();
            CreateAccount createAccount = new CreateAccount(); //navigate to create account
            this.NavigationService.Navigate(createAccount);
        }

        private void Forgot_Click(object sender, RoutedEventArgs e)
        { //when going to forgot password
            ID.Clear(); //clear the ID and password
            Password.Clear();
            ForgotPassword forgotPassword = new ForgotPassword(); //navigate to forgot password
            this.NavigationService.Navigate(forgotPassword);
        }

        private void ID_Enter(object sender, RoutedEventArgs e)
        { //when the user clicks on the ID box
            ID.Clear(); //clear the ID 
        }
        private void Password_Enter(object sender, RoutedEventArgs e)
        { //when the user clicks on the password box
            Password.Clear(); //clear the password
        }
    }
}
