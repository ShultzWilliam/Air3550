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
using Excel = Microsoft.Office.Interop.Excel;

namespace Air3550
{
    /// <summary>
    /// Interaction logic for ForgotPassword.xaml
    /// Screen users are taken to if they forgot their password
    /// </summary>
    public partial class ForgotPassword : Page
    {
        public ForgotPassword()
        {
            InitializeComponent();
        }

        private void Submit_Click(object sender, RoutedEventArgs e)
        { //submit the password

            //check if the email exists
            Functions functions = new Functions();

            //define the excel variables
            //Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = functions.database_connect();
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = functions.getRows(1);
            int IDRow = 0;
            string email = Email.Text;
            for (int i = 1; i <= rowCount; i++)
            { //go through the rows and find the email address
                if (email == xlRange.Cells[i, 10].Value2.ToString())
                { //if we find the email address
                    IDRow = i;
                }
            }

            xlWorkbook.Close(true);
            if (IDRow != 0)
            { //if the IDRow is not zero, we found the email address
                ResetPassword resetPassword = new ResetPassword(IDRow);
                this.NavigationService.Navigate(resetPassword);
            }
            else
            { //If it is zero, print a warning that we didn't find the email address
                Warning.Text = "Specified Email Address was not found in the database";
            }
            
        }
    }
}
