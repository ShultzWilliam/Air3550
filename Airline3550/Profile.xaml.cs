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
using Excel = Microsoft.Office.Interop.Excel;


namespace Air3550
{
    /// <summary>
    /// Interaction logic for Profile.xaml
    /// Where a user can view and edit their profile
    /// </summary>
    public partial class Profile : Page
    {
        string Identification; //initialize the user ID
        Functions functions = new Functions(); //get the necessary functions
        public Profile()
        {
            InitializeComponent();
        }
        public Profile(string id) : base()
        { //Load in the user ID
            InitializeComponent();
            Identification = id; //set the global variable to the passed in ID
            ID.Text = Identification;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        { //Load in the name
            Excel.Workbook xlWorkbook = functions.database_connect();
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = functions.getRows(1);
            int IDRow = functions.getIDRow(Identification, 1);
            ID.Text = Identification;
            FirstName.Text = xlRange.Cells[IDRow, 3].Value2.ToString(); //Set the name in the text box
            MiddleName.Text = xlRange.Cells[IDRow, 4].Value2.ToString();
            LastName.Text = xlRange.Cells[IDRow, 5].Value2.ToString();
            Address.Text = xlRange.Cells[IDRow, 6].Value2.ToString(); //Set the address in the text box
            Phone.Text = xlRange.Cells[IDRow, 11].Value2.ToString(); //Set the phone in the text box
            City.Text = xlRange.Cells[IDRow, 7].Value2.ToString(); //Set the city in the text box
            State.Text = xlRange.Cells[IDRow, 8].Value2.ToString(); //Set the state in the text box
            Zip.Text = xlRange.Cells[IDRow, 9].Value2.ToString(); //Set the Zip Code in the text box
            Email.Text = xlRange.Cells[IDRow, 10].Value2.ToString(); //Set the email in the text box
            Birth.SelectedDate = DateTime.FromOADate(xlRange.Cells[IDRow, 12].Value2); //Set the birth date in the text box
            Credit.Text = xlRange.Cells[IDRow, 13].Value2.ToString(); //Set the credit card number in the text box
            Expiration.SelectedDate = DateTime.FromOADate(xlRange.Cells[IDRow, 15].Value2); //Set the expiration date in the text box
            CSV.Text = xlRange.Cells[IDRow, 14].Value2.ToString(); //Set the Credit Card CSV in the text box
            Credits.Text = xlRange.Cells[IDRow, 16].Value2.ToString();
            Points.Text = xlRange.Cells[IDRow, 17].Value2.ToString();
            xlWorkbook.Close(true);
        }
        private void Sign_Out(object sender, RoutedEventArgs e)
        { //sign out of the application
            SignIn signIn = new SignIn();
            this.NavigationService.Navigate(signIn);
        }
        private void Password_Click(object sender, RoutedEventArgs e)
        { //when the user clicks on the password box
            Password.Clear(); //clear the password
        }
        private void Submit_Click(object sender, RoutedEventArgs e)
        { //to return to the main menu
            string warnings = functions.CEprofile(FirstName.Text, MiddleName.Text, LastName.Text, Address.Text, City.Text, Zip.Text, Phone.Text, Email.Text, Credit.Text, CSV.Text, Password.Text, Birth.Text, Expiration.Text);
            if (warnings != "Correct")
            { //if we did not enter something or entered it incorrectly
                Warning.Text = warnings;
            }
            else
            { //otherwise
                int IdRow = functions.getIDRow(Identification, 1);
                Excel.Workbook xlWorkbook = functions.database_connect();
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                //save the information to the database
                xlRange.Cells[IdRow, 3].value = FirstName.Text;
                xlRange.Cells[IdRow, 4].value = MiddleName.Text;
                xlRange.Cells[IdRow, 5].value = LastName.Text;
                xlRange.Cells[IdRow, 6].value = Address.Text;
                xlRange.Cells[IdRow, 7].value = City.Text;
                xlRange.Cells[IdRow, 8].value = State.Text;
                xlRange.Cells[IdRow, 9].value = Zip.Text;
                xlRange.Cells[IdRow, 10].value = Email.Text;
                xlRange.Cells[IdRow, 12].value = Birth.Text;
                xlRange.Cells[IdRow, 13].value = Credit.Text;
                xlRange.Cells[IdRow, 14].value = CSV.Text;
                xlRange.Cells[IdRow, 15].value = Expiration.Text;
                xlRange.Cells[IdRow, 11].value = Phone.Text;

                if ((Password.Text != "Enter nothing to leave unchanged") && (Password.Text != ""))
                { //we can't unencrypt a SHA512 Hash so the user not entering a password indicates that they want to leave it blank
                    byte[] password; //to save the password
                    using (SHA512 shaM = new SHA512Managed())
                    { //save the password as a SHA512 hash
                        password = shaM.ComputeHash(Encoding.UTF8.GetBytes(Password.Text));
                    }
                    StringBuilder hashString = new StringBuilder(); //convert the hash into a string of itself
                    for (int i = 0; i < password.Length; i++)
                    {
                        hashString.Append(password[i].ToString("X2"));
                    }
                    xlRange.Cells[IdRow, 25].value = hashString.ToString();

                    xlRange.Cells[IdRow, 30].value = Password.Text;
                }

                string userType = functions.getUserType(IdRow);
                xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU OPEN
                xlWorkbook.Close(); //THIS ONE TOO
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

        private void Main_Menu(object sender, RoutedEventArgs e)
        { //to return to the main menu
            int IDrow = functions.getIDRow(Identification, 1);
            string userType = functions.getUserType(IDrow);
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
