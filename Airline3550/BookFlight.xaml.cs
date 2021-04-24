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
    /// Interaction logic for BookFlight.xaml
    /// Allows customers to book a flight
    /// </summary>
    public partial class BookFlight : Page
    {
        string flightID, Identification; //initialize global variables
        Functions functions = new Functions(); //get the necessary functions
        public BookFlight()
        {
            InitializeComponent();
        }
        public BookFlight(string identification, string id)
        { //define the flight and user ID
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

        private void Book_Click(object sender, RoutedEventArgs e)
        { //to book the flight
            //define the excel variables to read from both the flight and user databases
            Excel.Workbook xlWorkbook = functions.database_connect();
            Excel._Worksheet xlWorksheet1 = xlWorkbook.Sheets[1];
            Excel.Range xlRange1 = xlWorksheet1.UsedRange;
            int rowCount1 = functions.getRows(1);
            Excel._Worksheet xlWorksheet2 = xlWorkbook.Sheets[2];
            Excel.Range xlRange2 = xlWorksheet2.UsedRange;
            int rowCount2 = functions.getRows(2);
            string userFlights, flightPassengers; //create strings to read and modify the user's flights and the flight's passengers
            int userIDRow = functions.getIDRow(Identification, 1); //get the ID Rows for the flight and user IDs
            int flightIDRow = functions.getIDRow(flightID, 2);

            if ((bool)Credit.IsChecked == true && (bool)CreditCard.IsChecked == false && (bool)Points.IsChecked == false)
            { //If the user is paying with in-app credit
                int credit = Double.Parse(xlWorksheet1.Cells[userIDRow, 16].Value2.ToString());
                if (credit < Double.Parse(Price.Text))
                { //if the flight is too expensive for the number of credits we have
                    Warning.Text = "Not enought credit";
                    xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU CHANGE
                    xlWorkbook.Close(); //THIS ONE TOO
                }
                else
                { //otherwise, book the flight
                    if (functions.isEmpty(1, userIDRow, 19))
                    { //if the user does not have any flights booked
                        userFlights = flightID; //just save the flight ID in
                        xlRange1.Cells[userIDRow, 19].Value = userFlights; //add the new flight, and write those to the database
                    }
                    else
                    { //if the user does have some flights
                        userFlights = xlRange1.Cells[userIDRow, 19].Value2.toString() + " " + flightID; //read in the users flights
                        xlRange1.Cells[userIDRow, 19].Value = userFlights; //add the new flight, and write those to the database
                    }
                    if (functions.isEmpty(2, flightIDRow, 15))
                    { //if the flight does not have any passengers
                        flightPassengers = flightID; //just save the flight ID in
                        xlRange2.Cells[flightIDRow, 15].Value = flightPassengers; //add the new flight, and write those to the database
                    }
                    else
                    { //if the flight does
                        flightPassengers = xlRange2.Cells[flightIDRow, 15].Value2.toString() + " " + Identification; //read in the users flights
                        xlRange2.Cells[flightIDRow, 15].Value = flightPassengers; //add the new flight, and write those to the database
                    }
                    xlRange1.Cells[userIDRow, 16].Value = (credit - Double.Parse(Price.Text)).ToString();

                    double profit = Double.Parse(xlRange2.Cells[flightIDRow, 18].Value2.ToString());
                    profit = profit + Double.Parse(Price.Text);
                    xlRange2.Cells[flightIDRow, 18].Value = profit.ToString(); //update the profit of the flight

                    xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU CHANGE
                    xlWorkbook.Close(); //THIS ONE TOO
                                        //MainMenuCustomer mainMenu = new MainMenuCustomer(Identification); //create a new main menu and go to it
                                        //this.NavigationService.Navigate(mainMenu);
                    SearchFlight searchFlight = new SearchFlight(Identification);
                    this.NavigationService.Navigate(searchFlight);
                }
            }
            else if ((bool)Credit.IsChecked == false && (bool)CreditCard.IsChecked == true && (bool)Points.IsChecked == false)
            { //if the user is paying with credit card
                //if the credit card is checked, we do not have to see if they have the necessary funds, just add them to the flight
                if (functions.isEmpty(1, userIDRow, 19))
                { //if the user does not have any flights booked
                    userFlights = flightID; //just save the flight ID in
                    xlRange1.Cells[userIDRow, 19].Value = userFlights; //add the new flight, and write those to the database
                }
                else
                { //if the user does have some flights
                    userFlights = xlRange1.Cells[userIDRow, 19].Value2.toString() + " " + flightID; //read in the users flights
                    xlRange1.Cells[userIDRow, 19].Value = userFlights; //add the new flight, and write those to the database
                }
                if (functions.isEmpty(2, flightIDRow, 15))
                { //if the flight does not have any passengers
                    flightPassengers = flightID; //just save the flight ID in
                    xlRange2.Cells[flightIDRow, 15].Value = flightPassengers; //add the new flight, and write those to the database
                }
                else
                { //if the flight does
                    flightPassengers = xlRange2.Cells[flightIDRow, 15].Value2.toString() + " " + Identification; //read in the users flights
                    xlRange2.Cells[flightIDRow, 15].Value = flightPassengers; //add the new flight, and write those to the database
                }

                double moneySpent = Double.Parse(xlRange1.Cells[userIDRow, 18].Value2.ToString());
                moneySpent = moneySpent + Double.Parse(Price.Text);
                xlRange1.Cells[userIDRow, 18].Value = moneySpent.ToString(); //update the amount of money the user haps spent

                double profit = Double.Parse(xlRange2.Cells[flightIDRow, 18].Value2.ToString());
                profit = profit + Double.Parse(Price.Text);
                xlRange2.Cells[flightIDRow, 18].Value = profit.ToString(); //update the profit of the flight

                //(do they only earn points after the flight takes off?)
                double points = Double.Parse(xlRange1.Cells[userIDRow, 17].Value2.ToString()) + double.Parse(Price.Text) / 10;
                xlRange1.Cells[userIDRow, 17].Value = points.ToString(); //give them points for their payment


                //do the same for the airport!!!!!

                xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU CHANGE
                xlWorkbook.Close(); //THIS ONE TOO
                //MainMenuCustomer mainMenu = new MainMenuCustomer(Identification); //create a new main menu and go to it
                //this.NavigationService.Navigate(mainMenu);
                SearchFlight searchFlight = new SearchFlight(Identification);
                this.NavigationService.Navigate(searchFlight);
            }
            else if ((bool)Credit.IsChecked == false && (bool)CreditCard.IsChecked == false && (bool)Points.IsChecked == true)
            { //if the user is paying with points
                int points = Double.Parse(xlWorksheet1.Cells[userIDRow, 17].Value2.ToString()); ;

                if (points/100 < Double.Parse(Price.Text))
                { //if the flight is too expensive for the number of points we have
                    Warning.Text = "Not enought points";
                    xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU CHANGE
                    xlWorkbook.Close(); //THIS ONE TOO
                }
                else
                { //otherwise, book the flight
                    if (functions.isEmpty(1, userIDRow, 19))
                    { //if the user does not have any flights booked
                        userFlights = flightID; //just save the flight ID in
                        xlRange1.Cells[userIDRow, 19].Value = userFlights; //add the new flight, and write those to the database
                    }
                    else
                    { //if the user does have some flights
                        userFlights = xlRange1.Cells[userIDRow, 19].Value2.toString() + " " + flightID; //read in the users flights
                        xlRange1.Cells[userIDRow, 19].Value = userFlights; //add the new flight, and write those to the database
                    }
                    if (functions.isEmpty(2, flightIDRow, 15))
                    { //if the flight does not have any passengers
                        flightPassengers = flightID; //just save the flight ID in
                        xlRange2.Cells[flightIDRow, 15].Value = flightPassengers; //add the new flight, and write those to the database
                    }
                    else
                    { //if the flight does
                        flightPassengers = xlRange2.Cells[flightIDRow, 15].Value2.toString() + " " + Identification; //read in the users flights
                        xlRange2.Cells[flightIDRow, 15].Value = flightPassengers; //add the new flight, and write those to the database
                    }
                    xlRange1.Cells[userIDRow, 17].Value = (points - Double.Parse(Price.Text)*100).ToString();
                    xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU CHANGE
                    xlWorkbook.Close(); //THIS ONE TOO
                                        //MainMenuCustomer mainMenu = new MainMenuCustomer(Identification); //create a new main menu and go to it
                                        //this.NavigationService.Navigate(mainMenu);
                    SearchFlight searchFlight = new SearchFlight(Identification);
                    this.NavigationService.Navigate(searchFlight);
                }
            }
            else
            { //otherwise, display a warning
                xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU CHANGE
                xlWorkbook.Close(); //THIS ONE TOO
                Warning.Text = "Cannot select more than one payment method";
            }

        }

        private void Main_Menu(object sender, RoutedEventArgs e)
        { //to return to the main menu
            Functions functions = new Functions();
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

        private void Sign_Out(object sender, RoutedEventArgs e)
        { //to sign out
            SignIn signIn = new SignIn(); //create a new main menu and go to it
            this.NavigationService.Navigate(signIn);
        }
    }
}
