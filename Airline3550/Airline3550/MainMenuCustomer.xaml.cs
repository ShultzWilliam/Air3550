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

namespace Airline3550
{
    /// <summary>
    /// Interaction logic for MainMenuCustomer.xaml
    /// </summary>
    public partial class MainMenuCustomer : Page
    {
        public string Identification;
        public MainMenuCustomer()
        {
            InitializeComponent();
        }

        public MainMenuCustomer(string id) : base()
        { //Load in the user ID
            InitializeComponent();
            Identification = id; //set the global variable to the passed in ID
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            User.Text = Identification; //Print the passed in ID
            //User.Text = "Name";
        }

        private void Sign_Out(object sender, RoutedEventArgs e)
        {
            SignIn signIn = new SignIn();
            this.NavigationService.Navigate(signIn);
        }
        private void Edit_Profile(object sender, RoutedEventArgs e)
        {
            Profile profile = new Profile(Identification);
            this.NavigationService.Navigate(profile);
        }
    }
}
