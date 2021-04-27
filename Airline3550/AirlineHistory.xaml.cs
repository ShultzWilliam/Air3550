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
    /// Interaction logic for AirlineHistory.xaml
    /// Displays the history of the entire airline for the accountant user class and
    /// allows them to pick an airport for more information
    /// </summary>
    public partial class AirlineHistory : Page
    {
        string identification, planeID;
        int counter, attendence;
        double totalProfit, attendenceRatio, sumProfit, planeCapacity;
        public AirlineHistory(string id) : base()
        { //initialize the user ID
            InitializeComponent();
            identification = id;

        }

        private void Sign_Out(object sender, RoutedEventArgs e)
        { //sign out of the application
            SignIn signIn = new SignIn();
            this.NavigationService.Navigate(signIn);
        }

        public class flightItem
        { //class used to insert flights into the data grid
            public string ID { get; set; }
            public string Origin { get; set; }
            public string Destination { get; set; }
            public string Departure { get; set; }
            public string Arrival { get; set; }
            public string Price { get; set; }
            public string Attendence { get; set; }
            public string AttendenceRatio { get; set; }
            public string Profit { get; set; }
        }
        private void Submit_Click(object sender, RoutedEventArgs e)
        {
            Functions functions = new Functions();
            Excel.Workbook xlWorkbook = functions.database_connect();
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[2];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int numRows = functions.getRows(2);
            for (int i = 2; i <= numRows; i++)
            {// print all the flights to the flight history


                //calculating attendence of a given flight
                //open up the users sheet of the database
                Excel.Workbook xlWorkbook2 = functions.database_connect();
                Excel._Worksheet xlWorksheet2 = xlWorkbook2.Sheets[1];
                Excel.Range xlRange2 = xlWorksheet2.UsedRange;
                int numRows2 = functions.getRows(1);

                //start the attendence and attendence ratio at 0
                attendence = 0;
                totalProfit = 0.0;

                for (int j = 2; j < numRows2; j++)
                {// cycle through the users tofind who is attending the flight
                    if (functions.isEmpty(1, j, 19) == false)
                    { // if the user has any purchased flights
                        string flights = xlRange2.Cells[j, 19].Value2.ToString();
                        if (flights.Contains(xlRange.Cells[i, 1].Value2.ToString()))
                        {   //if the user is scheduled for the flight
                            attendence++;
                        }
                    }
                    else if (functions.isEmpty(1, j, 22) == false)
                    { // if the user has any purchased flights
                        string flights = xlRange2.Cells[j, 22].Value2.ToString();
                        if (flights.Contains(xlRange.Cells[i, 1].Value2.ToString()))
                        {   //if the user was scheduled for the flight
                            attendence++;
                        }
                    }
                }
                xlWorkbook2.Close(true);

                //finding the attendence ratio of the flight
                //open up the planes sheet
                Excel.Workbook xlWorkbook3 = functions.database_connect();
                Excel._Worksheet xlWorksheet3 = xlWorkbook3.Sheets[3];
                Excel.Range xlRange3 = xlWorksheet3.UsedRange;

                //this is the id of the plane used in a given flight
                planeID = xlRange.Cells[i, 14].Value2.ToString();
                int numRows3 = functions.getRows(3);
                for (int k = 2; k < numRows3; k++)
                {//cycle through the planes page to find the plane used in a given flight
                    string tempID = xlRange3.Cells[k, 1].Value2.ToString();
                    //if we find the correct plane
                    if (tempID.Equals(planeID))
                    {
                        planeCapacity = xlRange3.Cells[k, 3].Value2;
                    }
                }
                xlWorkbook3.Close(true);

                if (attendence == 0) attendenceRatio = 0;
                else attendenceRatio = attendence / planeCapacity;
                totalProfit = attendence * xlRange.Cells[i, 17].Value2;

                // display the information about a given flight on the datagrid
                string temp = xlRange.Cells[i, 1].Value2.ToString();
                var item = new flightItem
                {
                    ID = xlRange.Cells[i, 1].Value2.ToString(),
                    Origin = functions.getAirport(xlRange.Cells[i, 5].Value2.ToString()),
                    Destination = functions.getAirport(xlRange.Cells[i, 6].Value2.ToString()),
                    Departure = DateTime.FromOADate(xlRange.Cells[i, 7].Value2).ToString("MM/dd/yyyy") + " " + DateTime.FromOADate(xlRange.Cells[i, 8].Value2).ToString("h:mm tt"),
                    Arrival = DateTime.FromOADate(xlRange.Cells[i, 10].Value2).ToString("MM/dd/yyyy") + " " + DateTime.FromOADate(xlRange.Cells[i, 11].Value2).ToString("h:mm tt"),
                    Price = "$" + xlRange.Cells[i, 17].Value2,
                    Attendence = attendence.ToString(),
                    AttendenceRatio = attendenceRatio.ToString(),
                    Profit = totalProfit.ToString()
                };
                Flights.Items.Add(item);
                counter++;
                sumProfit += totalProfit;
            }
            Total_Flights.Text = counter.ToString();
            Total_Profit.Text = sumProfit.ToString();
            counter = 0;
            sumProfit = 0;
            attendenceRatio = 0;
            xlWorkbook.Close(true);
        }

        private void Main_Menu(object sender, RoutedEventArgs e)
        { //to return to the main menu

            MainMenuAccountant mainMenu = new MainMenuAccountant(identification); //create a new main menu and go to it
            this.NavigationService.Navigate(mainMenu);

        }
    }
}
