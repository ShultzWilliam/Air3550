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
    /// Interaction logic for EditFlight.xaml
    /// Place where a load engineer can edit a existing flight
    /// </summary>
    public partial class EditFlight : Page
    {
        string Identification, flightID; //initialize global variables
        public EditFlight()
        {
            InitializeComponent();
        }
        public EditFlight(string identification, string FlightID)
        { //define the flight and user ID
            InitializeComponent();
            Identification = identification;
            flightID = FlightID;
            Window_Loaded();
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
        { //to submit flight changes

            //function to save the flight to the database
            Functions functions = new Functions();
            if (functions.isNum(Price.Text) && functions.isTime(Departure_Time.Text)
                && Departure_Date.SelectedDate.HasValue)
            { //if the inputs are correct

                //Add flight to excel doc
                //create the excel variables
                Excel.Workbook xlWorkbook = functions.database_connect();
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[2];
                Excel.Range xlRange = xlWorksheet.UsedRange;


                ////Get distance from Airport table
                string sDistance;
                int rowCount = functions.getRows(2), iArrivalInfoDivide, IDrow = -1;
                string sDZip = null, sOZip = null, sArrivalInfo = null, sArrivalTime, sArrivalDate, sDepartureDate;


                sOZip = functions.getZip(Origin_Textbox.Text);
                sDZip = functions.getZip(Destination_TextBox.Text);
                sDistance = functions.getDistance(Origin_Textbox.Text, Destination_TextBox.Text);
                sArrivalInfo = functions.getArrival(sDistance, Departure_Date.Text, Departure_Time.Text);
                iArrivalInfoDivide = sArrivalInfo.IndexOf("M");
                sArrivalTime = sArrivalInfo.Substring(0, iArrivalInfoDivide + 1);
                sArrivalDate = sArrivalInfo.Substring(iArrivalInfoDivide + 2, sArrivalInfo.Length - (iArrivalInfoDivide + 2));

                //Find the flight ID

                for (int i = 2; i <= rowCount; i++)
                {//find the flight id row
                    string temp = xlRange.Cells[i, 1].Value2.ToString();

                    if (temp == flightID)
                    {//get the row
                        IDrow = i;
                        break;
                    }

                }

                if (IDrow < 0)
                {
                    Warning.Text = "Flight ID does not exist";
                    return;
                }

                //For some reason the calandar box adds 12:00 AM to the end of all the dates
                sDepartureDate = Departure_Date.SelectedDate.ToString();
                int EndofDate = sDepartureDate.IndexOf(" ");
                sDepartureDate = sDepartureDate.Substring(0, EndofDate);

                xlRange.Cells[IDrow, 1].value = flightID;
                xlRange.Cells[IDrow, 2].value = "FALSE";
                xlRange.Cells[IDrow, 5].value = sOZip;
                xlRange.Cells[IDrow, 6].value = sDZip;
                xlRange.Cells[IDrow, 7].value = sDepartureDate;
                xlRange.Cells[IDrow, 8].value = Departure_Time.Text;
                xlRange.Cells[IDrow, 9].value = Departure_Terminal.Text;
                xlRange.Cells[IDrow, 10].value = sArrivalDate;
                xlRange.Cells[IDrow, 11].value = sArrivalTime;
                xlRange.Cells[IDrow, 12].value = Arrival_Terminal.Text;
                xlRange.Cells[IDrow, 13].value = sDistance;

                xlRange.Cells[rowCount + 1, 17].value = Price.Text;

                xlRange.Cells[rowCount + 1, 20].value = Identification;

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
            string sArrivalInfo, sDistance;
            string sArrivalTime;
            int iArrivalInfoDivide;

            Functions functions = new Functions();
            if (!functions.isTime(Departure_Time.Text))
            {
                Warning.Text = "Incorrect time input";
            }
            else
            {

                sDistance = functions.getDistance(Origin_Textbox.Text, Destination_TextBox.Text);
                sArrivalInfo = functions.getArrival(sDistance, Departure_Date.Text, Departure_Time.Text);
                iArrivalInfoDivide = sArrivalInfo.IndexOf("M");
                sArrivalTime = sArrivalInfo.Substring(0, iArrivalInfoDivide + 1);


                double price = 50; //set the base price
                //calculate the total price (12 cents per mile)
                //if a flight is a two leg, add $8

                price += (0.12 * double.Parse(sDistance));
                price = Math.Round(price, 2);

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

                Price.Text = price.ToString(); //display the price
                Arrival_Time.Text = sArrivalInfo; //display the arrival info
            }
        }

        private void Window_Loaded()
        { //to load in the flight's information

            Functions functions = new Functions();

            //Add flight to excel doc
            //create the excel variables
            Excel.Workbook xlWorkbook = functions.database_connect();
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[2];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int numRows = functions.getRows(2);


            for (int i = 2; i <= numRows; i++)
            {//find the flight id row
                string temp = xlRange.Cells[i, 1].Value2.ToString();

                if (temp == flightID)
                {//get the data from the sheet and add it to the gui

                    string Departure, Arrival, DepartureTime, DepartureDate, ArrivalTime, ArrivalDate;

                    Departure = DateTime.FromOADate(xlRange.Cells[i, 7].Value2).ToString("MM/dd/yyyy") + " " + DateTime.FromOADate(xlRange.Cells[i, 8].Value2).ToString("h:mm tt");
                    Arrival = DateTime.FromOADate(xlRange.Cells[i, 10].Value2).ToString("MM/dd/yyyy") + " " + DateTime.FromOADate(xlRange.Cells[i, 11].Value2).ToString("h:mm tt");

                    int dCutOff = Departure.IndexOf(" ");
                    DepartureTime = Departure.Substring(dCutOff + 1);

                    dCutOff = Arrival.IndexOf(" ");
                    ArrivalTime = Arrival.Substring(dCutOff + 1);

                    dCutOff = Departure.IndexOf(" ");
                    DepartureDate = Departure.Substring(0, dCutOff);
                    DateTime date = DateTime.Parse(DepartureDate);

                    dCutOff = Arrival.IndexOf(" ");
                    ArrivalDate = Arrival.Substring(0, dCutOff);


                    Origin_Textbox.Text = functions.getLocation(xlRange.Cells[i, 5].Value2.ToString());
                    Destination_TextBox.Text = functions.getLocation(xlRange.Cells[i, 6].Value2.ToString());
                    Departure_Date.SelectedDate = date;
                    Departure_Time.Text = DepartureTime;
                    Departure_Terminal.Text = xlRange.Cells[i, 9].Value2.ToString();
                    Arrival_Time.Text = ArrivalTime + " " + ArrivalDate;
                    Arrival_Terminal.Text = xlRange.Cells[i, 12].Value2.ToString();
                    Price.Text = xlRange.Cells[i, 17].Value2.ToString();

                    Origin_Textbox.IsReadOnly = true;
                    Destination_TextBox.IsReadOnly = true;
                    Departure_Terminal.IsReadOnly = true;
                    Arrival_Terminal.IsReadOnly = true;



                    //xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU OPEN
                    xlWorkbook.Close(); //THIS ONE TOO
                    return;
                }
            }

            Departure_Date.SelectedDate = DateTime.Today;
            Departure_Time.Text = "Departure Time";
            Departure_Terminal.Text = "Departure Terminal";
            Arrival_Time.Text = "Departure Time";
            Arrival_Terminal.Text = "Departure Terminal";
            Price.Text = "Too Expensive";

        }
        private void Delete_Click(object sender, RoutedEventArgs e)
        { //Remove flight from database

            Functions functions = new Functions();

            //Add flight to excel doc
            //create the excel variables
            Excel.Workbook xlWorkbook = functions.database_connect();
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[2];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int numRows = functions.getRows(2);

            for (int i = 2; i <= numRows; i++)
            {//find the flight id row
                string temp = xlRange.Cells[i, 1].Value2.ToString();

                if (temp == flightID)
                {//get the data from the sheet and add it to the gui

                    // 1. To Delete Entire Row - below rows will shift up
                    //xlRange.EntireRow.Delete(Type.Missing);
                    xlRange.Rows[i].Delete(Type.Missing);

                    xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU OPEN
                    xlWorkbook.Close(); //THIS ONE TOO

                    MainMenuLoadEngineer mainMenu = new MainMenuLoadEngineer(Identification); //create a new main menu and go to it
                    this.NavigationService.Navigate(mainMenu);
                    return;
                }
            }

            Warning.Text = "Flight does not Exist";
            //xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU OPEN
            xlWorkbook.Close(); //THIS ONE TOO
        }
    }
}
