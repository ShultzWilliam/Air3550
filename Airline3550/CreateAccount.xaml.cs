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

namespace Air3550
{
    /// <summary>
    /// Interaction logic for CreateAccount.xaml
    /// Allows the user to create an account
    /// </summary>
    public partial class CreateAccount : Page
    {
        public CreateAccount()
        {
            InitializeComponent();
        }

        private void Submit_Click(object sender, RoutedEventArgs e)
        {
            byte[] password; //to save the password
            using (SHA512 shaM = new SHA512Managed())
            { //save the password as a SHA512 hash
                password = shaM.ComputeHash(Encoding.UTF8.GetBytes(Password.Text));
            }
        }

    }
}
