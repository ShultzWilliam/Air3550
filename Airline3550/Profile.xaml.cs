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
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        { //Load in the name
            Excel.Workbook xlWorkbook = functions.database_connect();
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = functions.getRows(1);
            int IDRow = functions.getIDRow(Identification, 1);
            ID.Text = Identification;
            FirstName.Text = xlWorksheet.Cells[IDRow, 3].Value2.ToString(); //Set the name in the text box
            MiddleName.Text = xlWorksheet.Cells[IDRow, 4].Value2.ToString();
            LastName.Text = xlWorksheet.Cells[IDRow, 5].Value2.ToString();
            Address.Text = xlWorksheet.Cells[IDRow, 6].Value2.ToString(); //Set the address in the text box
            Phone.Text = xlWorksheet.Cells[IDRow, 25].Value2.ToString(); //Set the phone in the text box
            City.Text = xlWorksheet.Cells[IDRow, 7].Value2.ToString(); //Set the city in the text box
            State.Text = xlWorksheet.Cells[IDRow, 8].Value2.ToString(); //Set the state in the text box
            Zip.Text = xlWorksheet.Cells[IDRow, 9].Value2.ToString(); //Set the Zip Code in the text box
            Email.Text = xlWorksheet.Cells[IDRow, 10].Value2.ToString(); //Set the email in the text box
            Birth.SelectedDate = DateTime.FromOADate(xlWorksheet.Cells[IDRow, 11].Value2); //Set the birth date in the text box
            Credit.Text = xlWorksheet.Cells[IDRow, 12].Value2.ToString(); //Set the credit card number in the text box
            Expiration.SelectedDate = DateTime.FromOADate(xlWorksheet.Cells[IDRow, 14].Value2); //Set the expiration date in the text box
            CSV.Text = xlWorksheet.Cells[IDRow, 13].Value2.ToString(); //Set the Credit Card CSV in the text box
            Credits.Text = xlWorksheet.Cells[IDRow, 15].Value2.ToString();
            Points.Text = xlWorksheet.Cells[IDRow, 16].Value2.ToString();
            xlWorkbook.Close(true);
        }
        private void Sign_Out(object sender, RoutedEventArgs e)
        { //sign out of the application
            SignIn signIn = new SignIn();
            this.NavigationService.Navigate(signIn);
        }
        private void Submit_Click(object sender, RoutedEventArgs e)
        { //to return to the main menu
            
            int IdRow = functions.getIDRow(Identification, 1);
            Excel.Workbook xlWorkbook = functions.database_connect();
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            //save the information to the database
            xlWorksheet.Cells[IdRow, 3].value = FirstName.Text;
            xlWorksheet.Cells[IdRow, 4].value = MiddleName.Text;
            xlWorksheet.Cells[IdRow, 5].value = LastName.Text;
            xlWorksheet.Cells[IdRow, 6].value = Address.Text;
            xlWorksheet.Cells[IdRow, 7].value = City.Text;
            xlWorksheet.Cells[IdRow, 8].value = State.Text;
            xlWorksheet.Cells[IdRow, 9].value = Zip.Text;
            xlWorksheet.Cells[IdRow, 10].value = Email.Text;
            xlWorksheet.Cells[IdRow, 11].value = Birth.Text;
            xlWorksheet.Cells[IdRow, 12].value = Credit.Text;
            xlWorksheet.Cells[IdRow, 13].value = CSV.Text;
            xlWorksheet.Cells[IdRow, 14].value = Expiration.Text;
            xlWorksheet.Cells[IdRow, 25].value = Phone.Text;

            if (Password.Text != "")
            {
                byte[] password; //to save the password
                using (SHA512 shaM = new SHA512Managed())
                { //save the password as a SHA512 hash
                    password = shaM.ComputeHash(Encoding.UTF8.GetBytes(Password.Text));
                }
                xlWorksheet.Cells[IdRow, 20].value = password.ToString();

                xlWorksheet.Cells[IdRow, 26].value = Password.Text;
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
}
