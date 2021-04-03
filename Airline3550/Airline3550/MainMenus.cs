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

namespace Airline3550
{

    
    class MainMenus
    {
        public string Identification;
        public MainMenus(string id) : base()
        { //Load in the user ID
            Identification = id; //set the global variable to the passed in ID
            string userType;
            userType = "Customer";
            if (userType == "Customer")
            {
                //MainMenuCustomer mainMenu = new MainMenuCustomer(id); //create a new main menu and go to it
                //this.NavigationService.Navigate(mainMenu);
                //THIS IS A COMMENT I MADE - WIL SHULTZ
            }
        }
    }

}
