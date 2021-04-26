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
    /// Interaction logic for AccountentFlightRecord.xaml
    /// </summary>
    public partial class AccountentFlightRecord : Page
    {
        string Identification;
        public AccountentFlightRecord(string id)
        {
            Identification = id;
            InitializeComponent();
        }

        private void Sign_Out(object sender, RoutedEventArgs e)
        { //sign out of the application
            SignIn signIn = new SignIn();
            this.NavigationService.Navigate(signIn);
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

        }

        private void Main_Menu(object sender, RoutedEventArgs e)
        {
            MainMenuAccountant mainMenu = new MainMenuAccountant(Identification); //create a new main menu and go to it
            this.NavigationService.Navigate(mainMenu);
        }

        public class flightItem
        { //class used to insert flights into the data grid
            public string ID { get; set; }
            public string Origin { get; set; }
            public string Destination { get; set; }
            public string Departure { get; set; }
            public string Arrival { get; set; }
            public string Price { get; set; }
        }

        private void Submit_Click(object sender, RoutedEventArgs e)
        { //Go to the airport selected
            Functions functions = new Functions();
           // if (functions.isNum(flightID) == true)
           // { //check the the flight ID is a number (later we'll have to check if the flight ID exists
           // }
           // else
            {
                Warning.Text = "Incorrect Flight ID";
            }
        }
    }
}
