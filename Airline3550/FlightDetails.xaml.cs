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
            Price.Text = "$" + xlRange.Cells[IDRow, 17].Value2.ToString();
            Plane.Text = xlRange.Cells[IDRow, 14].Value2.ToString();
            xlWorkbook.Close(true);
        }
        private void Cancel_Flight(object sender, RoutedEventArgs e)
        { //to book the flight
            //cancel the flight and give the customer credit or their points back
            //check if we are less than an hour to departure to see if we can cancel yet

            //define the excel variables to read from both the flight and user databases
            Excel.Workbook xlWorkbook = functions.database_connect();
            Excel._Worksheet xlWorksheet1 = xlWorkbook.Sheets[1];
            Excel.Range xlRange1 = xlWorksheet1.UsedRange;
            int rowCount1 = functions.getRows(1);
            Excel._Worksheet xlWorksheet2 = xlWorkbook.Sheets[2];
            Excel.Range xlRange2 = xlWorksheet2.UsedRange;
            int rowCount2 = functions.getRows(2);
            string userFlights, userFlights2, flightPassengers, paymentMethods, paidWith, paymentMethods2; //create strings to read and modify the user's flights and the flight's passengers
            int userIDRow = functions.getIDRow(Identification, 1); //get the ID Rows for the flight and user IDs
            int flightIDRow = functions.getIDRow(flightID, 2);
            userFlights = xlRange1.Cells[userIDRow, 19].Value2.ToString(); //get the original userFlights, flightPassengers, and payment methods prior to 
            flightPassengers = xlRange2.Cells[flightIDRow, 15].Value2.ToString();
            paymentMethods = xlRange1.Cells[userIDRow, 20].Value2.ToString();
            //paidWith = xlRange1.Cells[userIDRow, 20].Value2.ToString(); //save what the flight was paid with
            paymentMethods2 = ""; //initialize these two strings
            userFlights2 = "";
            paidWith = "";

            if (flightID.Length == userFlights.Length)
            { //if the lengths of the two strings are equal, then this is the user's only flight
                if (functions.isEmpty(1, userIDRow, 22))
                { //if we haven't cancelled any flights yet, just save the info as it is
                    userFlights = userFlights.Replace(flightID, ""); //remove the flight for the user
                    xlRange1.Cells[userIDRow, 19].Value = userFlights; //save the new flights to the database
                    xlRange1.Cells[userIDRow, 20].Value = ""; //clear what the flight was paid with
                    //xlRange2.Cells[flightIDRow, 15].Value = flightPassengers.Replace(Identification, ""); //remove the passenger from the flight
                    xlRange1.Cells[userIDRow, 22].Value = flightID; //save to the cancelled info columns
                    xlRange1.Cells[userIDRow, 23].Value = paymentMethods;
                }
                else
                { //otherwise, add on the info to what's already there
                    userFlights = userFlights.Replace(flightID, ""); //remove the flight for the user
                    xlRange1.Cells[userIDRow, 19].Value = userFlights; //save the new flights to the database
                    xlRange1.Cells[userIDRow, 20].Value = ""; //clear what the flight was paid with
                    //xlRange2.Cells[flightIDRow, 15].Value = flightPassengers.Replace(Identification, ""); //remove the passenger from the flight
                    xlRange1.Cells[userIDRow, 22].Value = xlRange1.Cells[userIDRow, 22].Value2.ToString() + " " + flightID; //save to the cancelled info columns
                    xlRange1.Cells[userIDRow, 23].Value = xlRange1.Cells[userIDRow, 23].Value2.ToString() + " " + paymentMethods;
                }
            }
            else
            { //otherwise, we have to accomodate for other flights
                string[] userFlightsArray = userFlights.Split(' '); //create arrays to find the indeces of the flight and payment method
                string[] paymentMethodsArray = paymentMethods.Split(' ');
                int flightIndex = 0; //save the index of the flight
                int paymentLength = 0; //to help us make the new paymentMethods string
                for (int i = 0; i < userFlightsArray.Length; i++)
                { //go through the array until we find the user's flight index to find and remove the payment method
                    if (userFlightsArray[i] == flightID)
                    {
                        flightIndex = i; //save the index
                        paidWith = paymentMethodsArray[i];
                    }
                    else
                    { //otherwise, we write to and create a new paymentMethods string without the method of the flight we're cancelling
                        if (paymentLength == 0)
                        { //if we haven't written to paymentMethods2 yet, just write the method
                            paymentMethods2 = paymentMethodsArray[i];
                            userFlights2 = userFlightsArray[i];
                            paymentLength++;
                        }
                        else
                        { //otherwise, write a space and a method
                            paymentMethods2 = paymentMethods2 + " " + paymentMethodsArray[i];
                            userFlights2 = userFlights2 + " " + userFlightsArray[i];
                            paymentLength++;
                        }
                    }

                }
                if (functions.isEmpty(1, userIDRow, 22))
                { //if we haven't cancelled any flights yet, just save the info as it is
                    
                    //userFlights = userFlights.Replace(flightID, ""); //remove the flight for the user
                    xlRange1.Cells[userIDRow, 19].Value = userFlights2; //save the new flights to the database
                    xlRange1.Cells[userIDRow, 20].Value = paymentMethods2; //clear what the flight was paid with
                   // xlRange2.Cells[flightIDRow, 15].Value = flightPassengers.Replace(Identification, ""); //remove the passenger from the flight
                    xlRange1.Cells[userIDRow, 22].Value = flightID; //save to the cancelled info columns
                    xlRange1.Cells[userIDRow, 23].Value = paidWith;
                }
                else
                { //otherwise, add on the info to what's already there
                    
                    //userFlights = userFlights.Replace(flightID, ""); //remove the flight for the user
                    xlRange1.Cells[userIDRow, 19].Value = userFlights2; //save the new flights to the database
                    xlRange1.Cells[userIDRow, 20].Value = paymentMethods2; //clear what the flight was paid with
                    //xlRange2.Cells[flightIDRow, 15].Value = flightPassengers.Replace(Identification, ""); //remove the passenger from the flight
                    xlRange1.Cells[userIDRow, 22].Value = xlRange1.Cells[userIDRow, 22].Value2.ToString() + " " + flightID; //save to the cancelled info columns
                    xlRange1.Cells[userIDRow, 23].Value = xlRange1.Cells[userIDRow, 23].Value2.ToString() + " " + paidWith;
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
                double price = Double.Parse(xlRange2.Cells[flightIDRow, 7].Value2.ToString()); //get the price of the flight and the user's points
                double points = Double.Parse(xlRange1.Cells[userIDRow, 17].Value2.ToString()) * 100;
                points = points + price;
                xlRange1.Cells[userIDRow, 17].Value = points.ToString(); //save the updated points
            }
            else
            { //if we paid with points
                double price = Double.Parse(xlRange2.Cells[flightIDRow, 7].Value2.ToString()); //get the price of the flight and the user's in app credit
                double credit = Double.Parse(xlRange1.Cells[userIDRow, 16].Value2.ToString());
                credit = credit + price;
                xlRange1.Cells[userIDRow, 16].Value = credit.ToString(); //save the updated credit
            }

            //subtract from total profit
            

            xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU OPEN
            xlWorkbook.Close(); //THIS ONE TOO
            MyFlights myFlights = new MyFlights(Identification); //create a new main menu and go to it
            this.NavigationService.Navigate(myFlights);
        }
        private void Sign_Out(object sender, RoutedEventArgs e)
        { //to sign out
            SignIn signIn = new SignIn(); //create a new main menu and go to it
            this.NavigationService.Navigate(signIn);
        }


        private void Print_Pass(object sender, RoutedEventArgs e)
        { //Print the boarding pass
            BoardingPass boardingPass = new BoardingPass(Identification, flightID);
            boardingPass.Show();
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
