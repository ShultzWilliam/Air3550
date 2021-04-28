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

            //function to save the flight to the database
            Functions functions = new Functions();

            if (!(functions.isNum(Price.Text)))
            { //if the price input is incorrect
                Warning.Text = "Incorrect price input";
            }
            else if (!functions.isTime(Departure_Time.Text))
            { //if the time input is incorrect
                Warning.Text = "Incorrect Departure Time";
            }
            else if (!Departure_Date.SelectedDate.HasValue)
            {
                Warning.Text = "Missing Destination Location";
            }
            else if (String.IsNullOrEmpty(Arrival_Time.Text))
            {
                Warning.Text = "Please calculate an Arrival Time";
            }
            if (String.IsNullOrEmpty(Departure_Terminal.Text) || String.IsNullOrEmpty(Arrival_Terminal.Text))
            {
                Warning.Text = "Missing Terminal";
            }
            else
            { //if the inputs are correct

                //Add flight to excel doc
                //create the excel variables
                Excel.Workbook xlWorkbook = functions.database_connect();
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[2];
                Excel.Range xlRange = xlWorksheet.UsedRange;


                ////Get distance from Airport table
                string sDistance;
                int rowCount = functions.getRows(2), iArrivalInfoDivide;
                string sDZip = null, sOZip = null, sArrivalInfo = null, sArrivalTime, sArrivalDate, sDepartureDate;
               

                sOZip = functions.getZip(Origin.Text);
                sDZip = functions.getZip(Destination.Text);
                sDistance = functions.getDistance(Origin.Text, Destination.Text);
                sArrivalInfo = functions.getArrival(sDistance, Departure_Date.Text, Departure_Time.Text);
                iArrivalInfoDivide = sArrivalInfo.IndexOf("M");
                sArrivalTime = sArrivalInfo.Substring(0, iArrivalInfoDivide + 1);
                sArrivalDate = sArrivalInfo.Substring(iArrivalInfoDivide + 2, sArrivalInfo.Length - (iArrivalInfoDivide + 2));



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

                //For some reason the calandar box adds 12:00 AM to the end of all the dates
                sDepartureDate = Departure_Date.SelectedDate.ToString();
                int EndofDate = sDepartureDate.IndexOf(" ");
                sDepartureDate = sDepartureDate.Substring(0, EndofDate);

                xlRange.Cells[rowCount + 1, 1].value = flightID;
                xlRange.Cells[rowCount + 1, 2].value = "FALSE";
                xlRange.Cells[rowCount + 1, 5].value = sOZip;
                xlRange.Cells[rowCount + 1, 6].value = sDZip;
                xlRange.Cells[rowCount + 1, 7].value = sDepartureDate;
                xlRange.Cells[rowCount + 1, 8].value = Departure_Time.Text;
                xlRange.Cells[rowCount + 1, 9].value = Departure_Terminal.Text;
                xlRange.Cells[rowCount + 1, 10].value = sArrivalDate;
                xlRange.Cells[rowCount + 1, 11].value = sArrivalTime;
                xlRange.Cells[rowCount + 1, 12].value = Arrival_Terminal.Text;
                xlRange.Cells[rowCount + 1, 13].value = sDistance;

                xlRange.Cells[rowCount + 1, 16].value = "0";
                xlRange.Cells[rowCount + 1, 17].value = Price.Text;
                xlRange.Cells[rowCount + 1, 18].value = "0";

                xlRange.Cells[rowCount + 1, 20].value = Identification;

                xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU OPEN
                xlWorkbook.Close(); //THIS ONE TOO

                MainMenuLoadEngineer mainMenu = new MainMenuLoadEngineer(Identification); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }

        }

        private void Calculate_Click(object sender, RoutedEventArgs e)
        { //to calculate the price of the flight

            string sArrivalInfo, sDistance;
            string sArrivalTime;
            int iArrivalInfoDivide;

            Functions functions = new Functions();

            if (!functions.isTime(Departure_Time.Text))
            { //if the time input is incorrect
                Warning.Text = "Incorrect Departure Time";
            }
            else if (!Departure_Date.SelectedDate.HasValue)
            {
                Warning.Text = "Missing Destination Location";
            }
            else
            {

                sDistance = functions.getDistance(Origin.Text, Destination.Text);
                sArrivalInfo = functions.getArrival(sDistance, Departure_Date.Text, Departure_Time.Text);
                iArrivalInfoDivide = sArrivalInfo.IndexOf("M");
                sArrivalTime = sArrivalInfo.Substring(0, iArrivalInfoDivide + 1);


                double price = 50; //set the base price
                //calculate the total price (12 cents per mile)
                //if a flight is a two leg, add $8

                price += (0.12 * double.Parse(sDistance));

                int arrival, departure; //save values to get the aspects of the time and price 
                string arrivalHour, arrivalMinute, departureHour, departureMinute;
                double arrivalTime, departureTime;
                arrival = sArrivalTime.IndexOf(":"); //get the index of ":" in the arrival and departure times
                departure = Departure_Time.Text.IndexOf(":");
                arrivalHour = sArrivalTime.Substring(0, arrival); //get the arrival and departure hours and times
                arrivalMinute = sArrivalTime.Substring(arrival + 1, 2);
                departureHour = Departure_Time.Text.Substring(0, departure);
                departureMinute = Departure_Time.Text.Substring(departure + 1, 2);
                arrivalTime = Int32.Parse(arrivalHour) + ((double)Int32.Parse(arrivalMinute) / 60); //convert the time to an integer, with minutes as decimal
                departureTime = Int32.Parse(departureHour) + ((double)Int32.Parse(departureMinute) / 60); //convert the time to an integer, with minutes as decimal

                if (sArrivalTime.Contains("PM"))
                { //if it includes PM, add 12 to the time value
                    arrivalTime = arrivalTime + 12;
                }
                else if (sArrivalTime.Contains("AM") && arrivalHour == "12")
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

                price = Math.Round(price, 2);
                Price.Text = price.ToString(); //display the price
                Arrival_Time.Text = sArrivalInfo; //display the arrival info
            }
        }
    }
}
