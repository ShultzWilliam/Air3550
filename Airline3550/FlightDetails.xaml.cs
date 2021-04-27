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
    /// Interaction logic for FlightDetails.xaml
    /// Place where a customer can view the details of a yet to be taken flight
    /// </summary>
    public partial class FlightDetails : Page
    {
        string flightID, Identification; //initialize global variables
        Functions functions = new Functions(); //get the necessary functions
        bool booked = false;

        public FlightDetails()
        {
            InitializeComponent();
        }

        public FlightDetails(string identification, string id)
        { //define the user and flight IDs
            InitializeComponent();
            flightID = id;
            Identification = identification;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        { //when the window is loaded, load in the flight info

            //get the necessary excel variables
            Excel.Workbook xlWorkbook = functions.database_connect();
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[2];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = functions.getRows(2);
            int IDRow = functions.getIDRow(flightID, 2);
            Excel._Worksheet xlWorksheet1 = xlWorkbook.Sheets[1];
            Excel.Range xlRange1 = xlWorksheet1.UsedRange;
            int rowCount1 = functions.getRows(1);
            int userIDRow = functions.getIDRow(Identification, 1); //get the ID Rows for the flight and user IDs
            string userFlights = "";
            if (!(functions.isEmpty(1, userIDRow, 19)))
            {
                userFlights = xlRange1.Cells[userIDRow, 19].Value2.ToString(); //get the original userFlights

            }
            if (userFlights.Contains(flightID))
            { //make sure the user is booked for this flight, in the event they cancel a flight and press the back arrow
                booked = true;
                string paymentMethods = xlRange1.Cells[userIDRow, 20].Value2.ToString();
                string prices = xlRange1.Cells[userIDRow, 21].Value2.ToString();
                string paidWith = "";
                string price = "";
                string[] userFlightsArray = userFlights.Split(' '); //create arrays to find the indeces of the flight and payment method
                string[] paymentMethodsArray = paymentMethods.Split(' ');
                string[] pricesArray = prices.Split(' ');
                for (int i = 0; i < userFlightsArray.Length; i++)
                { //go through the array until we find the user's flight index to find the payment method
                    if (userFlightsArray[i] == flightID)
                    {
                        paidWith = paymentMethodsArray[i];
                        price = pricesArray[i];
                        break;
                    }

                }

                //load in the information to the page
                FlightID.Text = flightID;
                Origin.Text = functions.getAirport(xlRange.Cells[IDRow, 5].Value2.ToString());
                Destination.Text = functions.getAirport(xlRange.Cells[IDRow, 6].Value2.ToString());
                Departure_Date.Text = (DateTime.FromOADate(xlRange.Cells[IDRow, 7].Value2)).ToString("MM/dd/yyyy");
                Departure_Time.Text = (DateTime.FromOADate(xlRange.Cells[IDRow, 8].Value2)).ToString("h:mm tt");
                Departure_Terminal.Text = xlRange.Cells[IDRow, 9].Value2.ToString();
                Arrival_Date.Text = (DateTime.FromOADate(xlRange.Cells[IDRow, 10].Value2)).ToString("MM/dd/yyyy");
                Arrival_Time.Text = (DateTime.FromOADate(xlRange.Cells[IDRow, 11].Value2)).ToString("h:mm tt");
                Arrival_Terminal.Text = xlRange.Cells[IDRow, 12].Value2.ToString();
                Price.Text = "$" + price;
                Plane.Text = xlRange.Cells[IDRow, 14].Value2.ToString();
                Paid.Text = paidWith;
            }
            else
            {
                Warning.Text = "You are no longer booked for this flight";
            }
            xlWorkbook.Close(true);
        }
        private void Cancel_Flight(object sender, RoutedEventArgs e)
        { //to book the flight
          //cancel the flight and give the customer credit or their points back
          //check if we are less than an hour to departure to see if we can cancel yet
            Warning.Text = ""; //clear warning

            //define the excel variables to read from both the flight and user databases
            Excel.Workbook xlWorkbook = functions.database_connect();
            Excel._Worksheet xlWorksheet1 = xlWorkbook.Sheets[1];
            Excel.Range xlRange1 = xlWorksheet1.UsedRange;
            int rowCount1 = functions.getRows(1);
            Excel._Worksheet xlWorksheet2 = xlWorkbook.Sheets[2];
            Excel.Range xlRange2 = xlWorksheet2.UsedRange;
            int rowCount2 = functions.getRows(2);
            string userFlights, userFlights2, flightPassengers, paymentMethods, paidWith, paymentMethods2, prices, prices2; //create strings to read and modify the user's flights and the flight's passengers
            int userIDRow = functions.getIDRow(Identification, 1); //get the ID Rows for the flight and user IDs
            if (booked == true)
            { //make sure the user is booked for this flight, in the event they cancel a flight and press the back arrow
                int flightIDRow = functions.getIDRow(flightID, 2);
                double price = 0;
                userFlights = xlRange1.Cells[userIDRow, 19].Value2.ToString(); //get the original userFlights, flightPassengers, and payment methods prior to 
                flightPassengers = xlRange2.Cells[flightIDRow, 15].Value2.ToString();
                paymentMethods = xlRange1.Cells[userIDRow, 20].Value2.ToString();
                prices = xlRange1.Cells[userIDRow, 21].Value2.ToString();
                //paidWith = xlRange1.Cells[userIDRow, 20].Value2.ToString(); //save what the flight was paid with
                paymentMethods2 = ""; //initialize these two strings
                userFlights2 = "";
                prices2 = "";
                paidWith = paymentMethods;

                if ((DateTime.Parse(Departure_Date.Text + " " + Departure_Time.Text)).AddHours(-1) <= DateTime.Now)
                {
                if (flightID.Length == userFlights.Length)
                { //if the lengths of the two strings are equal, then this is the user's only flight
                    if (functions.isEmpty(1, userIDRow, 23))
                    { //if we haven't cancelled any flights yet, just save the info as it is
                        userFlights = userFlights.Replace(flightID, ""); //remove the flight for the user
                        xlRange1.Cells[userIDRow, 19].Value = userFlights; //save the new flights to the database
                        xlRange1.Cells[userIDRow, 20].Value = ""; //clear what the flight was paid with
                        xlRange1.Cells[userIDRow, 21].Value = ""; //clear what the flight was paid for
                                                                  //xlRange2.Cells[flightIDRow, 15].Value = flightPassengers.Replace(Identification, ""); //remove the passenger from the flight
                        xlRange1.Cells[userIDRow, 23].Value = flightID; //save to the cancelled info columns
                        xlRange1.Cells[userIDRow, 24].Value = paymentMethods;
                    }
                    else
                    { //otherwise, add on the info to what's already there
                        userFlights = userFlights.Replace(flightID, ""); //remove the flight for the user
                        xlRange1.Cells[userIDRow, 19].Value = userFlights; //save the new flights to the database
                        xlRange1.Cells[userIDRow, 20].Value = ""; //clear what the flight was paid with
                        xlRange1.Cells[userIDRow, 21].Value = ""; //clear what the flight was paid for

                        //xlRange2.Cells[flightIDRow, 15].Value = flightPassengers.Replace(Identification, ""); //remove the passenger from the flight
                        xlRange1.Cells[userIDRow, 23].Value = xlRange1.Cells[userIDRow, 23].Value2.ToString() + " " + flightID; //save to the cancelled info columns
                        xlRange1.Cells[userIDRow, 24].Value = xlRange1.Cells[userIDRow, 24].Value2.ToString() + " " + paymentMethods;
                    }
                    paidWith = paymentMethods;
                    price = Double.Parse(prices);
                    paymentMethods2 = "";
                    userFlights2 = "";
                    prices2 = "";
                }
                else
                { //otherwise, we have to accomodate for other flights

                    string[] userFlightsArray = userFlights.Split(' '); //create arrays to find the indeces of the flight and payment method
                    string[] paymentMethodsArray = paymentMethods.Split(' ');
                    string[] pricesArray = prices.Split(' ');
                    int flightIndex = 0; //save the index of the flight
                    int paymentLength = 0; //to help us make the new paymentMethods string
                    for (int i = 0; i < userFlightsArray.Length; i++)
                    { //go through the array until we find the user's flight index to find and remove the payment method
                        if (userFlightsArray[i] == flightID)
                        {
                            flightIndex = i; //save the index
                            paidWith = paymentMethodsArray[i];
                            price = Double.Parse(pricesArray[i]);
                        }
                        else
                        { //otherwise, we write to and create a new paymentMethods string without the method of the flight we're cancelling
                            if (paymentLength == 0)
                            { //if we haven't written to paymentMethods2 yet, just write the method
                                paymentMethods2 = paymentMethodsArray[i];
                                userFlights2 = userFlightsArray[i];
                                prices2 = pricesArray[i];
                                paymentLength++;
                            }
                            else
                            { //otherwise, write a space and a method
                                paymentMethods2 = paymentMethods2 + " " + paymentMethodsArray[i];
                                userFlights2 = userFlights2 + " " + userFlightsArray[i];
                                prices2 = prices2 + " " + pricesArray[i];
                                paymentLength++;
                            }
                        }

                    }



                    if (functions.isEmpty(1, userIDRow, 23))
                    { //if we haven't cancelled any flights yet, just save the info as it is

                        //userFlights = userFlights.Replace(flightID, ""); //remove the flight for the user
                        xlRange1.Cells[userIDRow, 19].Value = userFlights2; //save the new flights to the database
                        xlRange1.Cells[userIDRow, 20].Value = paymentMethods2; //clear what the flight was paid with
                        xlRange1.Cells[userIDRow, 21].Value = prices2; //clear what the flight was paid with

                        // xlRange2.Cells[flightIDRow, 15].Value = flightPassengers.Replace(Identification, ""); //remove the passenger from the flight
                        xlRange1.Cells[userIDRow, 23].Value = flightID; //save to the cancelled info columns
                        xlRange1.Cells[userIDRow, 24].Value = paidWith;
                    }
                    else
                    { //otherwise, add on the info to what's already there

                        //userFlights = userFlights.Replace(flightID, ""); //remove the flight for the user
                        xlRange1.Cells[userIDRow, 19].Value = userFlights2; //save the new flights to the database
                        xlRange1.Cells[userIDRow, 20].Value = paymentMethods2; //clear what the flight was paid with
                        xlRange1.Cells[userIDRow, 21].Value = prices2; //clear what the flight was paid with

                        //xlRange2.Cells[flightIDRow, 15].Value = flightPassengers.Replace(Identification, ""); //remove the passenger from the flight
                        xlRange1.Cells[userIDRow, 23].Value = xlRange1.Cells[userIDRow, 23].Value2.ToString() + " " + flightID; //save to the cancelled info columns
                        xlRange1.Cells[userIDRow, 24].Value = xlRange1.Cells[userIDRow, 24].Value2.ToString() + " " + paidWith;
                    }
                }

                string[] passengersArray = flightPassengers.Split(' '); //create an array of the passengers
                if (passengersArray[0] == Identification)
                { //if the user's ID is the first in the array, we need to remove just the user ID
                    xlRange2.Cells[flightIDRow, 15].Value = flightPassengers.Replace(Identification, ""); //remove the passenger from the flight
                }
                else
                { //otherwise, we only need to remove the user ID and space before it
                    xlRange2.Cells[flightIDRow, 15].Value = flightPassengers.Replace(" " + Identification, ""); //remove the passenger from the flight
                }

                if (paidWith == "Points")
                { //if we paid with a credit card or with in app credit
                    //double price = Double.Parse(xlRange2.Cells[flightIDRow, 17].Value2.ToString()) * 100; //get the price of the flight and the user's points
                    double points = Double.Parse(xlRange1.Cells[userIDRow, 17].Value2.ToString());
                    points = points + price * 100;
                    xlRange1.Cells[userIDRow, 17].Value = points.ToString(); //save the updated points
                    double pointsUsed = Convert.ToDouble(xlRange1.Cells[userIDRow, 30].Value2.ToString());
                    xlRange1.Cells[userIDRow, 30].Value = (pointsUsed - price * 100).ToString(); //increment the points used
                }
                else
                { //if we paid with points
                    //double price = Double.Parse(xlRange2.Cells[flightIDRow, 17].Value2.ToString()); //get the price of the flight and the user's in app credit
                    double credit = Double.Parse(xlRange1.Cells[userIDRow, 16].Value2.ToString());
                    credit = credit + price;
                    xlRange1.Cells[userIDRow, 16].Value = credit.ToString(); //save the updated credit


                }

                double profit = Double.Parse(xlRange2.Cells[flightIDRow, 18].Value2.ToString());
                //profit = profit - Double.Parse(xlRange2.Cells[flightIDRow, 17].Value2.ToString());
                profit = profit - Convert.ToDouble(xlRange2.Cells[flightIDRow, 17].Value2.ToString());
                xlRange2.Cells[flightIDRow, 18].Value = profit.ToString(); //update the profit of the flight

                double attendance = Double.Parse(xlRange2.Cells[flightIDRow, 16].Value2.ToString()) - 1;
                xlRange2.Cells[flightIDRow, 16].Value = attendance.ToString(); //update the attendance as well

                Excel._Worksheet xlWorksheet3 = xlWorkbook.Sheets[4]; //for the airport
                Excel.Range xlRange3 = xlWorksheet3.UsedRange;
                int rowCount3 = functions.getRows(4);
                int airportIDRow = functions.getIDRow(xlRange2.Cells[flightIDRow, 5].Value2.ToString(), 4); //get the row for the origin airport
                xlRange3.Cells[airportIDRow, 6].Value = Double.Parse(xlRange3.Cells[airportIDRow, 6].Value2.ToString()) - price; //decrement the profit for the airport

                //subtract from total profit


                xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU OPEN
                xlWorkbook.Close(); //THIS ONE TOO
                MyFlights myFlights = new MyFlights(Identification); //create a new main menu and go to it
                this.NavigationService.Navigate(myFlights);
                }
                else
                {
                Warning.Text = "Cannot cancel flight, flights can only be cancelled up to an hour before departure time";
                xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU OPEN
                xlWorkbook.Close(); //THIS ONE TOO
                }
            }
            else
            {
                Warning.Text = "You are no longer booked for this flight";
                xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU OPEN
                xlWorkbook.Close(); //THIS ONE TOO
            }




        }
        private void Sign_Out(object sender, RoutedEventArgs e)
        { //to sign out
            SignIn signIn = new SignIn(); //create a new main menu and go to it
            this.NavigationService.Navigate(signIn);
        }


        private void Print_Pass(object sender, RoutedEventArgs e)
        { //Print the boarding pass
            Warning.Text = ""; //clear warning
            if (booked == true)
            { //make sure the user is booked for this flight, in the event they cancel a flight and press the back arrow
                if ((DateTime.Parse(Departure_Date.Text + " " + Departure_Time.Text)).AddHours(-24) <= DateTime.Now)
                { //if we are within 24 hours to departure
                    BoardingPass boardingPass = new BoardingPass(Identification, flightID);
                    boardingPass.Show();
                }
                else
                { //otherwise, print that we can't print yet
                    Warning.Text = "Cannot print boarding pass. Boarding Passes may be printed at 24 hours prior to departure";
                }
            }
            else
            {
                Warning.Text = "You are no longer booked for this flight";
            }
        }

        private void Main_Menu(object sender, RoutedEventArgs e)
        { //to return to the main menu
            int IDrow = functions.getIDRow(Identification, 1);
            string userType = functions.getUserType(IDrow);
            if (userType == "Customer")
            {
                MainMenuCustomer mainMenu = new MainMenuCustomer(Identification); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
            else if (userType == "Load Engineer")
            {
                MainMenuLoadEngineer mainMenu = new MainMenuLoadEngineer(Identification); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
            else if (userType == "Accountant")
            {
                MainMenuAccountant mainMenu = new MainMenuAccountant(Identification); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
            else if (userType == "Marketing Manager")
            {
                MainMenuMarketingManager mainMenu = new MainMenuMarketingManager(Identification); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
            else if (userType == "Flight Manager")
            {
                MainMenuFlightManager mainMenu = new MainMenuFlightManager(Identification); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
        }
    }
}
