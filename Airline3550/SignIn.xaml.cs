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
using System.Security.Cryptography;
using Excel = Microsoft.Office.Interop.Excel;

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
            Functions functions = new Functions();

            //define the excel variables
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = functions.database_connect();
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            int IDcolumn = 0;
            byte[] passwordHash;
            byte[] foundPassword;

            String id, password; //initialize strings for the ID and Password
            //bool IDfound = false;
            id = ID.Text; //set the strings
            password = Password.Text;
            using (SHA512 shaM = new SHA512Managed())
            { //save the password as a SHA512 hash
                passwordHash = shaM.ComputeHash(Encoding.UTF8.GetBytes(Password.Text));
            }
            password = passwordHash.ToString();
            IDcolumn = functions.getIDColumn(id);

            if (IDcolumn != 0)
            { //grab the password from the database
                password = xlRange.Cells[IDcolumn, 20].Value2;
                foundPassword = Encoding.ASCII.GetBytes(password);

                if ((foundPassword.ToString() == passwordHash.ToString()))
                {
                    string userType = functions.getUserType(IDcolumn); //get the user type from the database
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
                else
                {
                    //Warning.Text = "Incorrect Password \n" + passwordHash + "\n" + foundPassword;
                    Warning.Text = "Incorrect Password";
                }
            }
            else if (IDcolumn == 0)
            {
                Warning.Text = "ID not found ";
            }

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
