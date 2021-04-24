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
    /// Interaction logic for BoardingPass.xaml
    /// </summary>
    public partial class BoardingPass : Window
    {
        string flightID, Identification; //initialize global variables
        Functions functions = new Functions(); //get the necessary functions
        public BoardingPass()
        {
            InitializeComponent();
        }
        public BoardingPass(string identification, string id)
        { //define the user and flight IDs
            InitializeComponent();
            flightID = id;
            Identification = identification;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        { //when the window is loaded, load in the flight info

            //get the necessary excel variables
            Excel.Workbook xlWorkbook = functions.database_connect();
            Excel._Worksheet xlWorksheet1 = xlWorkbook.Sheets[1];
            Excel.Range xlRange1 = xlWorksheet1.UsedRange;
            int rowCount = functions.getRows(1);
            Excel._Worksheet xlWorksheet2 = xlWorkbook.Sheets[2];
            Excel.Range xlRange2 = xlWorksheet2.UsedRange;
            int rowCount2 = functions.getRows(2);
            int userIDRow = functions.getIDRow(Identification, 1);
            int IDRow = functions.getIDRow(flightID, 2);

            //load in the information to the page
            FlightID.Text = flightID;
            Name.Text = xlRange1.Cells[userIDRow, 3].Value2.ToString() + " " + xlRange1.Cells[userIDRow, 5].Value2.ToString();
            Origin.Text = functions.getAirport(xlRange2.Cells[IDRow, 5].Value2.ToString());
            Destination.Text = functions.getAirport(xlRange2.Cells[IDRow, 6].Value2.ToString());
            Departure_Date.Text = (DateTime.FromOADate(xlRange2.Cells[IDRow, 7].Value2)).ToString("MM/dd/yyyy");
            Departure_Time.Text = (DateTime.FromOADate(xlRange2.Cells[IDRow, 8].Value2)).ToString("h:mm tt");
            Departure_Terminal.Text = xlRange2.Cells[IDRow, 9].Value2.ToString();
            Arrival_Date.Text = (DateTime.FromOADate(xlRange2.Cells[IDRow, 10].Value2)).ToString("MM/dd/yyyy");
            Arrival_Time.Text = (DateTime.FromOADate(xlRange2.Cells[IDRow, 11].Value2)).ToString("h:mm tt");
            Arrival_Terminal.Text = xlRange2.Cells[IDRow, 12].Value2.ToString();
            UserID.Text = xlRange1.Cells[userIDRow, 1].Value2.ToString();
            xlWorkbook.Close(true);
        }
    }
}
