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
    /// Interaction logic for ResetPassword.xaml
    /// Where a user can reset their password
    /// </summary>
    public partial class ResetPassword : Page
    {
        string Email;
        int IDRow;
        public ResetPassword()
        {
            InitializeComponent();
        }

        public ResetPassword(int IdRow)
        { //get the email and set the parameters
            InitializeComponent();
            //get the email address
            IDRow = IdRow; //define the ID Row
            Functions functions = new Functions();

            //define the excel variables
            //Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = functions.database_connect();
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = functions.getRows(1);
            ID.Text = xlRange.Cells[IDRow, 1].Value2.ToString(); //display the user's ID number

            xlWorkbook.Close(true);
        }
        
        private void Submit_Click(object sender, RoutedEventArgs e)
        { //change the password
            if ((Password1.Text != "") && (Password2.Text != "") && (Password1.Text == Password2.Text))
            { //if the passwords match, save changes

                //save changes
                Functions functions = new Functions();

                //define the excel variables
                //Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = functions.database_connect();
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                int rowCount = functions.getRows(1);
                byte[] passwordHash; //to save the password
                using (SHA512 shaM = new SHA512Managed())
                { //save the password as a SHA512 hash
                    passwordHash = shaM.ComputeHash(Encoding.UTF8.GetBytes(Password1.Text));
                }
                StringBuilder hashString = new StringBuilder(); //convert the hash into a string of itself
                for (int i = 0; i < passwordHash.Length; i++)
                {
                    hashString.Append(passwordHash[i].ToString("X2"));
                }
                string password = hashString.ToString();
                xlRange.Cells[IDRow, 25].value = password;

                //xlRange.Cells[IDRow, 30].value = Password1.Text;

                xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU OPEN
                xlWorkbook.Close(); //THIS ONE TOO

                SignIn signIn = new SignIn();
                this.NavigationService.Navigate(signIn);
            }
            else
            { //otherwise, display a warning
                Warning.Text = "Passwords don't match or are empty, try again";
            }
            
        }
    }
}
