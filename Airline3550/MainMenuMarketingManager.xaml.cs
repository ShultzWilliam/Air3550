﻿using System;
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
    /// Interaction logic for MainMenuMarketingManager.xaml
    /// Main menu for the marketing manager user type
    /// </summary>
    public partial class MainMenuMarketingManager : Page
    {

        public string Identification; //define the global ID

        public MainMenuMarketingManager()
        {
            InitializeComponent();
        }
        public MainMenuMarketingManager(string id) : base()
        { //Load in the user ID
            InitializeComponent();
            Identification = id; //set the global variable to the passed in ID
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        { //load in the user information
            Functions functions = new Functions();
            int IDRow = functions.getIDRow(Identification, 1); //get the ID column for the user
            User.Text = functions.getName(IDRow); //Print the passed in ID
        }

        private void Sign_Out(object sender, RoutedEventArgs e)
        { //sign out of the application
            SignIn signIn = new SignIn();
            this.NavigationService.Navigate(signIn);
        }
        private void Edit_Profile(object sender, RoutedEventArgs e)
        { //edit your profile
            Profile profile = new Profile(Identification);
            this.NavigationService.Navigate(profile);
        }

        private void Schedule(object sender, RoutedEventArgs e)
        { //Search for flights to book planes for
            SearchPlane searchPlane = new SearchPlane(Identification);
            this.NavigationService.Navigate(searchPlane);
        }
    }
}
