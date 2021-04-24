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
    /// Interaction logic for CreateFlight.xaml
    /// Allows the load engineer to create a flight
    /// </summary>
    public partial class CreateFlight : Page
    {
        string Identification; //initialize the user ID
        public CreateFlight()
        {
            InitializeComponent();
        }
        public CreateFlight(string identification)
        { //define the user ID
            InitializeComponent();
            Identification = identification;
        }

        private void Sign_Out(object sender, RoutedEventArgs e)
        { //to sign out
            SignIn signIn = new SignIn(); //create a new main menu and go to it
            this.NavigationService.Navigate(signIn);
        }
        private void Main_Menu(object sender, RoutedEventArgs e)
        { //to return to the main menu
            MainMenuLoadEngineer mainMenu = new MainMenuLoadEngineer(Identification); //create a new main menu and go to it
            this.NavigationService.Navigate(mainMenu);
        }
        private void Submit_Click(object sender, RoutedEventArgs e)
        { //to create the flight

            int NumAirports = 14;

            //function to save the flight to the database
            Functions functions = new Functions();
            if (functions.isNum(Price.Text) && functions.isTime(Arrival_Time.Text) && functions.isTime(Departure_Time.Text)
                && Departure_Date.SelectedDate.HasValue && Arrival_Date.SelectedDate.HasValue)
            { //if the inputs are correct

                //Add flight to excel doc
                //create the excel variables
                Excel.Workbook xlWorkbook = functions.database_connect();
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[2];
                Excel._Worksheet xlWorksheetAir = xlWorkbook.Sheets[4];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                Excel.Range xlRangeAir = xlWorksheetAir.UsedRange;

                //Get distance from Airport table
                string sDistance = null;
                int rowCount = functions.getRows(2);
                int DistanceRow = 2;
                string Location = Origin.Text;
                string Test;

                for (int i = 2; i <= NumAirports; i++)
                {//Once origin row is found
                    Test = xlRangeAir.Cells[i, 4].Value2.ToString();
                    if (xlRangeAir.Cells[i, 4].Value2.ToString() == Location)
                    {//Look for ending location
                        DistanceRow = i;
                        break;
                    }
                }

                Location = Destination.Text;
                for (int j = 10; j < j + NumAirports; j++)
                {//Once ending location is found
                    Test = xlRangeAir.Cells[1, j].Value2.ToString();
                    if (xlRangeAir.Cells[1, j].Value2.ToString() == Location)
                    {//Get distance
                        sDistance = xlRangeAir.Cells[DistanceRow, j].Value2.ToString();
                        break;
                    }
                }


                string flightID = "opsie";
                bool taken = true;
                Random r = new Random(); //create a random number
                                         //we need to create a random six digit ID number that hasn't been taken already
                                         //therefore, we'll create a random number, loop through the users table, compare,
                                         //and, if it hasn't been taken, assign it. Otherwise, we'll try again
                while (taken == true)
                {
                    taken = false;
                    int id;
                    id = r.Next(100000, 999999); //get a random number
                    flightID = id.ToString(); //convert it to a string
                    for (int i = 1; i <= rowCount; i++)
                    {
                        if (xlRange.Cells[i, 1].Value2.ToString() == flightID)
                        {
                            taken = true;
                        }
                    }
                }

                xlRange.Cells[rowCount, 1] = flightID;

                xlRange.Cells[rowCount, 5] = Origin.Text;
                xlRange.Cells[rowCount, 6].Value2 = Destination.Text;
                //xlRange.Cells[rowCount, 7].value = Departure_Date.SelectedDate.ToString();
                xlRange.Cells[rowCount, 8].Value2 = Departure_Time.Text;
                xlRange.Cells[rowCount, 9].Value2 = Departure_Terminal.Text;
                xlRange.Cells[rowCount, 10].Value2 = Arrival_Date.SelectedDate.ToString();
                //xlRange.Cells[rowCount, 11].value = Arrival_Time.Text;
                xlRange.Cells[rowCount, 12].Value2 = Arrival_Terminal.Text;
                xlRange.Cells[rowCount, 13].Value2 = sDistance;

                xlRange.Cells[rowCount, 17].Value2 = Price.Text;

                xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU OPEN
                xlWorkbook.Close(); //THIS ONE TOO

                MainMenuLoadEngineer mainMenu = new MainMenuLoadEngineer(Identification); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
            else if (!(functions.isNum(Price.Text)))
            { //if the price input is incorrect
                Warning.Text = "Incorrect price input";
            }
            else if (!(functions.isTime(Arrival_Time.Text) && functions.isTime(Departure_Time.Text)))
            { //if the time input is incorrect
                Warning.Text = "Incorrect time input";
            }
            else
            {
                Warning.Text = "Missing Destination or Arrival Location";
            }

        }
        private void Calculate_Click(object sender, RoutedEventArgs e)
        { //to calculate the price of the flight
            Functions functions = new Functions();
            if (!(functions.isTime(Arrival_Time.Text) && functions.isTime(Departure_Time.Text)))
            {
                Warning.Text = "Incorrect time input";
            }
            else
            {
                double price = 50; //set the base price
                //calculate the total price (12 cents per mile)
                //if a flight is a two leg, add $8

                int arrival, departure; //save values to get the aspects of the time and price 
                string arrivalHour, arrivalMinute, departureHour, departureMinute;
                double arrivalTime, departureTime;
                arrival = Arrival_Time.Text.IndexOf(":"); //get the index of ":" in the arrival and departure times
                departure = Departure_Time.Text.IndexOf(":");
                arrivalHour = Arrival_Time.Text.Substring(0, arrival); //get the arrival and departure hours and times
                arrivalMinute = Arrival_Time.Text.Substring(arrival + 1, 2);
                departureHour = Departure_Time.Text.Substring(0, departure);
                departureMinute = Departure_Time.Text.Substring(departure + 1, 2);
                arrivalTime = Int32.Parse(arrivalHour) + ((double)Int32.Parse(arrivalMinute) / 60); //convert the time to an integer, with minutes as decimal
                departureTime = Int32.Parse(departureHour) + ((double)Int32.Parse(departureMinute) / 60); //convert the time to an integer, with minutes as decimal
                if (Arrival_Time.Text.Contains("PM"))
                { //if it includes PM, add 12 to the time value
                    arrivalTime = arrivalTime + 12;
                }
                else if (Arrival_Time.Text.Contains("AM") && arrivalHour == "12")
                { //if it's 12 AM
                    arrivalTime = arrivalTime - 12;
                }
                if (Departure_Time.Text.Contains("PM"))
                { //if it includes PM, add 12 to the time value
                    departureTime = departureTime + 12;
                }
                else if (Departure_Time.Text.Contains("AM") && departureHour == "12")
                { //if it's 12 AM
                    departureTime = departureTime - 12;
                }
                if (arrivalTime < 5.0 || departureTime < 5.0)
                { //if they're arriving between midnight and 5AM, give them the 20% red eye discount
                    price = price * 0.8;
                }
                else if (departureTime < 8.0 || arrivalTime > 19.0)
                { //for the off peak discount
                    price = price * 0.9;
                }

                Price.Text = price.ToString(); //display the price
            }
        }
    }
}
