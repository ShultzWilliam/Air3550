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
            State.Text = "Ohio";
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
            if (warnings != "Correct" && Password.Text == "")
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
            byte[] password; //to save the password
            using (SHA512 shaM = new SHA512Managed())
            { //save the password as a SHA512 hash
                password = shaM.ComputeHash(Encoding.UTF8.GetBytes(Password.Text));
            }
            xlWorksheet.Cells[rowCount + 1, 1].value = Identification;
            xlWorksheet.Cells[rowCount + 1, 2].value = "Customer";
            xlWorksheet.Cells[rowCount + 1, 3].value = FirstName.Text;
            xlWorksheet.Cells[rowCount + 1, 4].value = MiddleName.Text;
            xlWorksheet.Cells[rowCount + 1, 5].value = LastName.Text;
            xlWorksheet.Cells[rowCount + 1, 6].value = Address.Text;
            xlWorksheet.Cells[rowCount + 1, 7].value = City.Text;
            xlWorksheet.Cells[rowCount + 1, 8].value = State.Text;
            xlWorksheet.Cells[rowCount + 1, 9].value = Zip.Text;
            xlWorksheet.Cells[rowCount + 1, 10].value = Email.Text;
            xlWorksheet.Cells[rowCount + 1, 11].value = Phone.Text;
            xlWorksheet.Cells[rowCount + 1, 12].value = Birth.Text;
            xlWorksheet.Cells[rowCount + 1, 13].value = Credit.Text;
            xlWorksheet.Cells[rowCount + 1, 14].value = CSV.Text;
            xlWorksheet.Cells[rowCount + 1, 15].value = Expiration.Text;
            xlWorksheet.Cells[rowCount + 1, 22].value = password.ToString();
            xlWorksheet.Cells[rowCount + 1, 27].value = Password.Text;
            xlWorksheet.Cells[rowCount + 1, 16].value = "0"; //Credit
            xlWorksheet.Cells[rowCount + 1, 17].value = "0"; //Points
            xlWorksheet.Cells[rowCount + 1, 18].value = "0"; //Money Spent
            xlWorkbook.Close(true);

            MainMenuCustomer mainMenu = new MainMenuCustomer(Identification); //create a new main menu and go to it
            this.NavigationService.Navigate(mainMenu);
            
            }
            

        }

    }
}
