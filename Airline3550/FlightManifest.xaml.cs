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
    /// Interaction logic for FlightManifest.xaml
    /// Place where a flight manager can print a flight manifest or search for customer records
    /// </summary>
    public partial class FlightManifest : Page
    {
        string flightID, identification, IDCHECK; //initialize global variables

        public FlightManifest()
        {
            InitializeComponent();
        }
        public FlightManifest(string id, string ID) : base()
        { //define the user and flight IDs
            InitializeComponent();
            identification = id;
            flightID = ID;
        }

        private void Sign_Out(object sender, RoutedEventArgs e)
        { //sign out of the application
            SignIn signIn = new SignIn();
            this.NavigationService.Navigate(signIn);
        }
        private void Submit_Click(object sender, RoutedEventArgs e)
        { //Go back to the main menu
            IDCHECK = FlightID.Text;
            //check if the flightID exists

            Functions functions = new Functions();
            if (functions.isNum(IDCHECK) == true)
            { //if the user ID exists, go to the user
                FlightID.Text = string.Empty;
                UserRecord flightManifest = new UserRecord(identification, IDCHECK);
                this.NavigationService.Navigate(flightManifest);
            }
            else
            { //otherwise, display an error
                Warning.Text = "Invalid Flight ID";
            }
        }
        private void Main_Menu(object sender, RoutedEventArgs e)
        { //to return to the main menu
            MainMenuFlightManager mainMenu = new MainMenuFlightManager(identification); //create a new main menu and go to it
            this.NavigationService.Navigate(mainMenu);
        }
        private void Print(object sender, RoutedEventArgs e)
        {
            IDCHECK = FlightID.Text;
            //check if the flightID exists

            Functions functions = new Functions();
            if (functions.isNum(IDCHECK) == true)
            {
                // finding all the crew members and attendents
                Excel.Workbook xlWorkbook = functions.database_connect();
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[2];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                int numRows = functions.getRows(2);
                Crew.Text = "Crew members: ";
                for (int i = 2; i <= numRows; i++)
                {//find the flight ID row
                    string temp = xlRange.Cells[i, 1].Value2.ToString();

                    if (temp == FlightID.Text)
                    {
                        string crew, attendants;
                        crew = xlRange.Cells[i, 3].Value2.ToString();
                        attendants = xlRange.Cells[i, 4].Value2.ToString();
                        Crew.Text += crew + ", " + attendants;
                    }
                }
                xlWorkbook.Close(true);

                // finding all the customers
               Excel.Workbook xlWorkbook2 = functions.database_connect();
               Excel._Worksheet xlWorksheet2 = xlWorkbook2.Sheets[1];
               Excel.Range xlRange2 = xlWorksheet2.UsedRange;
               Passengers.Text = "Passengers: ";
               int numRows2 = functions.getRows(1);
              // Passengers.Text += numRows2;
               for (int i = 2; i < numRows2; i++)
               {// find who is attending the flight
                    if (functions.isEmpty(1, i, 19) == false)
                    { // if the user has any purchased flights
                        string flights = xlRange2.Cells[i, 19].Value2.ToString();
                        if (flights.Contains(FlightID.Text))
                        {   //if the user is scheduled for the flight
                            Passengers.Text += xlRange2.Cells[i, 3].Value2.ToString() + " " + xlRange2.Cells[i, 5].Value2.ToString() + ", ";
                        }
                    }
                    else if (functions.isEmpty(1, i, 22) == false)
                    { // if the user has any purchased flights
                        string flights = xlRange2.Cells[i, 22].Value2.ToString();
                        if (flights.Contains(FlightID.Text))
                        {   //if the user was scheduled for the flight
                            Passengers.Text += xlRange2.Cells[i, 3].Value2.ToString() + " " + xlRange2.Cells[i, 5].Value2.ToString() + ", ";
                        }
                    }
                }
                Passengers.Text = Passengers.Text.Substring(0, Passengers.Text.Length - 2);
                xlWorkbook2.Close(true); 
            }
            else
            { //otherwise, display an error
                Warning.Text = "Invalid Flight ID";
            }
        }
    }
}
