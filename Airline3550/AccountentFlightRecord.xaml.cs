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
    /// Interaction logic for AccountentFlightRecord.xaml
    /// </summary>
    public partial class AccountentFlightRecord : Page
    {
        string Identification, IDCHECK; //initialize global variables
        int attendence;
        double price, totalProfit;
        public AccountentFlightRecord(string id)
        {
            Identification = id;
            InitializeComponent();
        }

        private void Sign_Out(object sender, RoutedEventArgs e)
        { //sign out of the application
            SignIn signIn = new SignIn();
            this.NavigationService.Navigate(signIn);
        }

        private void Main_Menu(object sender, RoutedEventArgs e)
        {
            MainMenuAccountant mainMenu = new MainMenuAccountant(Identification); //create a new main menu and go to it
            this.NavigationService.Navigate(mainMenu);
        }

        public class flightItem
        { //class used to insert flights into the data grid
            public string ID { get; set; }
            public string Origin { get; set; }
            public string Destination { get; set; }
            public string Departure { get; set; }
            public string Arrival { get; set; }
            public string Price { get; set; }
        }
        private void Submit_Click(object sender, RoutedEventArgs e)
        { //Go to the airport selected
            IDCHECK = FlightID.Text;
            Functions functions = new Functions();
            if (functions.isNum(IDCHECK) == true)
            { //check the the flight ID is a number (later we'll have to check if the flight ID exists
                Excel.Workbook xlWorkbook2 = functions.database_connect();
                Excel._Worksheet xlWorksheet2 = xlWorkbook2.Sheets[1];
                Excel.Range xlRange2 = xlWorksheet2.UsedRange;
                int numRows2 = functions.getRows(1);
                attendence = 0;
                totalProfit = 0.0;
                // Passengers.Text += numRows2;
                for (int i = 2; i < numRows2; i++)
                {// find who is attending the flight
                    if (functions.isEmpty(1, i, 19) == false)
                    { // if the user has any purchased flights
                        string flights = xlRange2.Cells[i, 19].Value2.ToString();
                        if (flights.Contains(FlightID.Text))
                        {   //if the user is scheduled for the flight
                            attendence++;
                        }
                    }
                    else if (functions.isEmpty(1, i, 22) == false)
                    { // if the user has any purchased flights
                        string flights = xlRange2.Cells[i, 22].Value2.ToString();
                        if (flights.Contains(FlightID.Text))
                        {   //if the user was scheduled for the flight
                            attendence++;
                        }
                    }
                }
                Attendence.Text = attendence.ToString();
                xlWorkbook2.Close(true);

                //calculating profit of flight and displaying the flight to the grid
                Flights.Items.Clear();
                Excel.Workbook xlWorkbook = functions.database_connect();
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[2];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                int numRows = functions.getRows(2);
                for (int i = 2; i <= numRows; i++)
                {//find the flight ID row
                    string temp = xlRange.Cells[i, 1].Value2.ToString();

                    if (temp == FlightID.Text)
                    {
                        var item = new flightItem
                        {
                            ID = FlightID.Text,
                            Origin = functions.getAirport(xlRange.Cells[i, 5].Value2.ToString()),
                            Destination = functions.getAirport(xlRange.Cells[i, 6].Value2.ToString()),
                            Departure = DateTime.FromOADate(xlRange.Cells[i, 7].Value2).ToString("MM/dd/yyyy") + " " + DateTime.FromOADate(xlRange.Cells[i, 8].Value2).ToString("h:mm tt"),
                            Arrival = DateTime.FromOADate(xlRange.Cells[i, 10].Value2).ToString("MM/dd/yyyy") + " " + DateTime.FromOADate(xlRange.Cells[i, 11].Value2).ToString("h:mm tt"),
                            Price = "$" + xlRange.Cells[i, 17].Value2
                        };
                        Flights.Items.Add(item);
                        price = xlRange.Cells[i, 17].Value2;
                    }
                }
                totalProfit = price * (double)attendence;
                Profit.Text = totalProfit.ToString();
                xlWorkbook.Close(true);


            }
            else
            {
                Warning.Text = "Incorrect Flight ID";
            }
        }
    }
}
