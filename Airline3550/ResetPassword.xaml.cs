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
    /// Interaction logic for ResetPassword.xaml
    /// Where a user can reset their password
    /// </summary>
    public partial class ResetPassword : Page
    {
        string Email;
        public ResetPassword()
        {
            InitializeComponent();
        }

        public ResetPassword(string email)
        { //get the email and set the parameters
            InitializeComponent();
            //get the email address
            Email = email;
        }
        
        private void Submit_Click(object sender, RoutedEventArgs e)
        { //change the password
            if (Password1.Text == Password2.Text)
            { //if the passwords match, save changes

                //save changes

                SignIn signIn = new SignIn();
                this.NavigationService.Navigate(signIn);
            }
            else
            { //otherwise, display a warning
                Warning.Text = "Passwords don't match, try again";
            }
            
        }
    }
}
