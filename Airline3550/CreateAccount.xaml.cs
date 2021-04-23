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
    /// Interaction logic for CreateAccount.xaml
    /// Allows the user to create an account
    /// </summary>
    public partial class CreateAccount : Page
    {
        Functions functions = new Functions(); //get the necessary functions
        string Identification;
        public CreateAccount()
        {
            InitializeComponent();
            //create the excel variables
            
            Excel.Workbook xlWorkbook = functions.database_connect();
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = functions.getRows(1);
            int colCount = xlRange.Columns.Count;
            bool taken = true;
            Random r = new Random(); //create a random number
            //we need to create a random six digit ID number that hasn't been taken already
            //therefore, we'll create a random number, loop through the users table, compare,
            //and, if it hasn't been taken, assign it. Otherwise, we'll try again
            while(taken == true)
            {
                taken = false;
                int id;
                id = r.Next(100000, 999999); //get a random number
                Identification = id.ToString(); //convert it to a string
                for (int i = 1; i <= rowCount; i++)
                {
                    if (xlRange.Cells[i, 1].Value2.ToString() == Identification)
                    {
                        taken = true;
                    }
                }
            }
            ID.Text = Identification;
            xlWorkbook.Close();
            
        }

        private void Submit_Click(object sender, RoutedEventArgs e)
        { //click to create the account
            string warnings = functions.CEprofile(FirstName.Text, MiddleName.Text, LastName.Text, Address.Text, City.Text, Zip.Text, Phone.Text, Email.Text, Credit.Text, CSV.Text, Password.Text, Birth.Text, Expiration.Text);
            if (warnings != "Correct")
            { //if we did not enter something or entered it incorrectly
                Warning.Text = warnings;
            }
            else
            { //otherwise
            //create the excel variables
            Excel.Workbook xlWorkbook = functions.database_connect();
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            byte[] passwordHash; //to save the password
            using (SHA512 shaM = new SHA512Managed())
            { //save the password as a SHA512 hash
                passwordHash = shaM.ComputeHash(Encoding.UTF8.GetBytes(Password.Text));
            }
            StringBuilder hashString = new StringBuilder(); //convert the hash into a string of itself
            for (int i = 0; i < passwordHash.Length; i++)
            {
                hashString.Append(passwordHash[i].ToString("X2"));
            }
            string password = hashString.ToString();
            xlRange.Cells[rowCount + 1, 1].value = Identification;
                xlRange.Cells[rowCount + 1, 2].value = "Customer";
                xlRange.Cells[rowCount + 1, 3].value = FirstName.Text;
                xlRange.Cells[rowCount + 1, 4].value = MiddleName.Text;
                xlRange.Cells[rowCount + 1, 5].value = LastName.Text;
                xlRange.Cells[rowCount + 1, 6].value = Address.Text;
                xlRange.Cells[rowCount + 1, 7].value = City.Text;
                xlRange.Cells[rowCount + 1, 8].value = State.Text;
                xlRange.Cells[rowCount + 1, 9].value = Zip.Text;
                xlRange.Cells[rowCount + 1, 10].value = Email.Text;
                xlRange.Cells[rowCount + 1, 11].value = Phone.Text;
                xlRange.Cells[rowCount + 1, 12].value = Birth.Text;
                xlRange.Cells[rowCount + 1, 13].value = Credit.Text;
                xlRange.Cells[rowCount + 1, 14].value = CSV.Text;
                xlRange.Cells[rowCount + 1, 15].value = Expiration.Text;
                xlRange.Cells[rowCount + 1, 22].value = password;
                xlRange.Cells[rowCount + 1, 27].value = Password.Text;
                xlRange.Cells[rowCount + 1, 16].value = "0"; //Credit
                xlRange.Cells[rowCount + 1, 17].value = "0"; //Points
                xlRange.Cells[rowCount + 1, 18].value = "0"; //Money Spent
            xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU OPEN
            xlWorkbook.Close();

            MainMenuCustomer mainMenu = new MainMenuCustomer(Identification); //create a new main menu and go to it
            this.NavigationService.Navigate(mainMenu);
            
            }
            

        }

    }
}
