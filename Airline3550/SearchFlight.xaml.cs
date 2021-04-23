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
    /// Interaction logic for SearchFlight.xaml
    /// </summary>
    public partial class SearchFlight : Page
    {
        //AdventureWorksLT2008Entities dataEntities = new AdventureWorksLT2008Entities();
        string origin, destination, flightID, identification; //initialize global variables
        public SearchFlight()
        {
            InitializeComponent();
        }
        public SearchFlight(string id) : base()
        { //define the user ID
            InitializeComponent();
            identification = id;
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
        private void Sign_Out(object sender, RoutedEventArgs e)
        { //sign out of the application
            SignIn signIn = new SignIn();
            this.NavigationService.Navigate(signIn);
        }
        private void Search_Click(object sender, RoutedEventArgs e)
        { //search for flights
            Flights.Items.Clear();
            //get access to functions
            Functions functions = new Functions();

            //get the origin, destination and dates
            origin = functions.getAirportCode(Start.Text);
            destination = functions.getAirportCode(End.Text);
            
            if (!(origin == "" || destination == "" || Departure.Text == "" || Arrival.Text == "" || Departure.Text == "Select a date" || Arrival.Text == "Select a date"))
            { //if origin, destination, startDate, and endDate have all been assigned values
                DateTime startDate = Convert.ToDateTime(Departure.Text);
                DateTime endDate = Convert.ToDateTime(Arrival.Text);
                //string checker;

                //define the excel variables
                //Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = functions.database_connect();
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[2];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                int rowCount = functions.getRows(2);
                int colCount = xlRange.Columns.Count;
                int[,] flight = new int[rowCount, 2]; //two dimensional array of potential flights
                int[] twoD = new int[rowCount]; //use to record which flight returned is a round trip

                int numOfFlights = 0; //the number of flights in the array
                int attendance, attendance2; //the attendance of the flight
                string plane, plane2; //the plane ID
                DateTime foundDate, foundDate2, foundEndDate, earliestEndDate;
                double foundTime, foundTime2, earliestEndTime;
                earliestEndDate = DateTime.Today;
                earliestEndTime = 0;
                string flights = "empty"; //string to save the flights we currently have


                //need to adjust for two leg flight and round trip

                for (int i = 2; i <= rowCount; i++)
                { //Find the flights going to and from the origin and destination
                    //string userType = xlRange.Cells[IDcolumn, 2].Value2.ToString(); //get the user type from the database
                    //origin is 5, destination is 6, date is 7th row
                    foundDate = DateTime.FromOADate(xlRange.Cells[i, 7].Value2); //get the date of the flight
                    
                    if (foundDate >= startDate && foundDate <= endDate)
                    { //if the flight takes place between the start and end date
                        if (xlRange.Cells[i, 5].Value2.ToString() == origin && xlRange.Cells[i, 6].Value2.ToString() == destination)
                        { //if the origin and destination match
                            foundTime = Convert.ToDouble(xlRange.Cells[i, 11].Value2.ToString()); //find the end date and time
                            foundEndDate = DateTime.FromOADate(xlRange.Cells[i, 10].Value2);
                            attendance = Int32.Parse(xlRange.Cells[i, 16].Value2.ToString()); //along with the attendance and plane
                            plane = xlRange.Cells[i, 14].Value2.ToString();
                            if ((functions.fullFlight(attendance, plane) == false) && ((flights == "empty" || !(flights.Contains(xlRange.Cells[i, 1].Value2.ToString())))))
                            { //if the flight isn't full and the flight isn't already in the array
                                flight[numOfFlights, 0] = i; //save the index of the flight
                                numOfFlights++; //increment the number of flights
                                if (numOfFlights == 1)
                                { //set the initial value of earliestEndTime and earliestEndDate
                                    earliestEndTime = Convert.ToDouble(xlRange.Cells[i, 11].Value2.ToString());
                                    earliestEndDate = DateTime.FromOADate(xlRange.Cells[i, 10].Value2);
                                    flights = xlRange.Cells[i, 1].Value2.ToString();
                                }
                                else if (numOfFlights > 1 && (foundEndDate < earliestEndDate) && (foundTime < earliestEndTime))
                                { //set the new earliest end time and date
                                    earliestEndTime = Convert.ToDouble(xlRange.Cells[i, 11].Value2.ToString());
                                    earliestEndDate = DateTime.FromOADate(xlRange.Cells[i, 10].Value2);
                                    flights = flights + " " + xlRange.Cells[i, 1].Value2.ToString();
                                }
                            }
                        }
                        else if (xlRange.Cells[i, 5].Value2.ToString() == origin && xlRange.Cells[i, 6].Value2.ToString() != destination)
                        { //if the origin matches but the destination doesn't
                            string leg = xlRange.Cells[i, 6].Value2.ToString(); //save the value of the leg
                            int legs = 0; //count the number of second locations we have
                            foundTime = Convert.ToDouble(xlRange.Cells[i, 11].Value2.ToString()); //save the end time and date of the first flight
                            foundEndDate = DateTime.FromOADate(xlRange.Cells[i, 10].Value2);
                            attendance = Int32.Parse(xlRange.Cells[i, 16].Value2.ToString()); //along with the plane and attendance
                            plane = xlRange.Cells[i, 14].Value2.ToString();
                            if ((functions.fullFlight(attendance, plane) == false) && ((flights == "empty" || !(flights.Contains(xlRange.Cells[i, 1].Value2.ToString())))))
                            { //if the flight isn't full
                                for (int j = 2; j < rowCount; j++)
                                { //go through the list again looking for 2nd legs
                                    foundDate2 = DateTime.FromOADate(xlRange.Cells[j, 7].Value2); //get the date of the 2nd flight
                                    foundTime2 = Convert.ToDouble(xlRange.Cells[j, 8].Value2.ToString()); //get the start time as well
                                    if (foundDate2 >= startDate && foundDate2 <= endDate && (foundDate2 > foundDate || (foundDate2 == foundDate && foundTime2 > foundTime)))
                                    { //if the flight takes place between the start and end date and is after the first leg flight
                                        if (xlRange.Cells[j, 5].Value2.ToString() == leg && xlRange.Cells[j, 6].Value2.ToString() == destination)
                                        { //if the origin and destination match
                                            attendance2 = Int32.Parse(xlRange.Cells[j, 16].Value2.ToString());
                                            plane2 = xlRange.Cells[j, 14].Value2.ToString(); //save the second flight's plane and attendance
                                            if ((functions.fullFlight(attendance2, plane2) == false) && ((flights == "empty" || !(flights.Contains(xlRange.Cells[j, 1].Value2.ToString())))))
                                            { //if the flight isn't full
                                                if (legs == 0)
                                                { //if this is the first leg
                                                    flight[numOfFlights, 0] = i; //save the index of the flight
                                                    numOfFlights++;
                                                    flight[numOfFlights, 0] = j; //save the index of the second flight
                                                    numOfFlights++;
                                                    legs++; //increment the number of legs

                                                    if (numOfFlights == 1)
                                                    { //set the initial value of earliestEndTime and earliestEndDate
                                                        earliestEndTime = Convert.ToDouble(xlRange.Cells[i, 11].Value2.ToString());
                                                        earliestEndDate = DateTime.FromOADate(xlRange.Cells[i, 10].Value2);
                                                        flights = xlRange.Cells[i, 1].Value2.ToString() + " " + xlRange.Cells[j, 1].Value2.ToString();
                                                    }
                                                    else if (numOfFlights > 1 && (foundEndDate < earliestEndDate) && (foundTime < earliestEndTime))
                                                    { //set the new earliest end time and date
                                                        earliestEndTime = Convert.ToDouble(xlRange.Cells[i, 11].Value2.ToString());
                                                        earliestEndDate = DateTime.FromOADate(xlRange.Cells[i, 10].Value2);
                                                        flights = flights + " " +  xlRange.Cells[i, 1].Value2.ToString() + " " + xlRange.Cells[j, 1].Value2.ToString();
                                                    }
                                                }
                                                else
                                                {
                                                    flight[numOfFlights, 0] = j; //save the index of the flight
                                                    numOfFlights++;
                                                    legs++;
                                                    flights = flights + " " + xlRange.Cells[j, 1].Value2.ToString();
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                if (((bool)RoundTrip.IsChecked == true) && (numOfFlights >= 1))
                { //if round trip is checked, loop back through and flights back that take place after the earlies flight to
                    for (int i = 2; i <= rowCount; i++)
                    { //Find the flights going to and from the origin to the destination
                      //origin is 5, destination is 6, date is 7th row
                        foundDate = DateTime.FromOADate(xlRange.Cells[i, 7].Value2); //get the date of the flight

                        if (foundDate >= startDate && foundDate <= endDate)
                        { //if the flight takes place between the start and end date
                            if (xlRange.Cells[i, 5].Value2.ToString() == destination && xlRange.Cells[i, 6].Value2.ToString() == origin)
                            { //if the origin and destination match
                                foundTime = Convert.ToDouble(xlRange.Cells[i, 8].Value2.ToString());
                                foundEndDate = DateTime.FromOADate(xlRange.Cells[i, 10].Value2);
                                attendance = Int32.Parse(xlRange.Cells[i, 16].Value2.ToString());
                                if ((foundEndDate > earliestEndDate) || (foundEndDate == earliestEndDate) && (earliestEndTime < foundTime))
                                {
                                    plane = xlRange.Cells[i, 14].Value2.ToString();
                                    if (functions.fullFlight(attendance, plane) == false)
                                    { //if the flight isn't full
                                        flight[numOfFlights, 0] = i; //save the index of the flight
                                        numOfFlights++;
                                        flights = flights + " " + xlRange.Cells[i, 1].Value2.ToString();
                                        
                                    }
                                }
                                
                            }

                        }
                    }
                }


                if (endDate < DateTime.Today)
                {
                    Warning.Text = "Please search for flights that are in the future";
                }
                if (numOfFlights == 0)
                { //if num of flights = 0, then we didn't find any flights
                    Warning.Text = "No flights found";
                }
                else
                { //otherwise
                    for (int i = 0; i < numOfFlights; i++)
                    { //for each flight we found
                        var item = new flightItem
                        {
                            ID = xlRange.Cells[flight[i, 0], 1].Value2.ToString(),
                            Origin = functions.getAirport(xlRange.Cells[flight[i, 0], 5].Value2.ToString()),
                            Destination = functions.getAirport(xlRange.Cells[flight[i, 0], 6].Value2.ToString()),
                            Departure = DateTime.FromOADate(xlRange.Cells[flight[i, 0], 7].Value2).ToString("MM/dd/yyyy") + " " + DateTime.FromOADate(xlRange.Cells[flight[i, 0], 8].Value2).ToString("h:mm tt"),
                            Arrival = DateTime.FromOADate(xlRange.Cells[flight[i, 0], 10].Value2).ToString("MM/dd/yyyy") + " " + DateTime.FromOADate(xlRange.Cells[flight[i, 0], 11].Value2).ToString("h:mm tt"),
                            Price = "$" + xlRange.Cells[flight[i, 0], 17].Value2.ToString()
                        }; //create a new flight item to insert into the data grid
                        Flights.Items.Add(item);
                    }
                }
                xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU CHANGE
                xlWorkbook.Close();
            }
            else
            { //if one of them has not been set
                Warning.Text = "Assign values for origin, destination, and departure and arrival dates";
            }
        }
        private void Submit_Click(object sender, RoutedEventArgs e)
        { //Go back to the main menu
            flightID = FlightID.Text;

            //check if the flightID exists
            Functions functions = new Functions();
            if((functions.isNum(flightID) == true) && (functions.isFlight(flightID) == true))
            { //if the flight ID exists, go to the flight
                if (functions.flightBooked(identification, flightID) == false)
                { //if the user is not already booked for the flight, take them to the booking page
                    BookFlight bookFlight = new BookFlight(identification, flightID);
                    this.NavigationService.Navigate(bookFlight);
                }
                else
                { //otherwise, display a warning telling them they are already booked for it
                    Warning.Text = "You are already booked for this flight";
                }
                
            }
            else
            { //otherwise, display an error
                Warning.Text = "Invalid Flight ID";
            }
            
        }

        private void Main_Menu(object sender, RoutedEventArgs e)
        { //to return to the main menu
            Functions functions = new Functions();
            int IDrow = functions.getIDRow(identification, 1);
            string userType = functions.getUserType(IDrow);
            if (userType == "Customer")
            {
                MainMenuCustomer mainMenu = new MainMenuCustomer(identification); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
            else if (userType == "Load Engineer")
            {
                MainMenuLoadEngineer mainMenu = new MainMenuLoadEngineer(identification); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
            else if (userType == "Accountant")
            {
                MainMenuAccountant mainMenu = new MainMenuAccountant(identification); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
            else if (userType == "Marketing Manager")
            {
                MainMenuMarketingManager mainMenu = new MainMenuMarketingManager(identification); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
            else if (userType == "Flight Manager")
            {
                MainMenuFlightManager mainMenu = new MainMenuFlightManager(identification); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);
            }
        }
    }
}
