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
        string origin, destination, identification; //initialize global variables
        int flightID;
        Functions functions = new Functions();
        int numOfFlights = 0; //the number of flights in the array

        int[,] flight; //two dimensional array of potential flights
        int[] flightLegs; //use to record how many lets each flight has
        string[] flightType; //hold records as to whether a flight is to or return

        public SearchFlight()
        {
            InitializeComponent();
        }
        public SearchFlight(string id) : base()
        { //define the user ID
            InitializeComponent();
            identification = id;
            int rowCount = functions.getRows(2);
            int power = rowCount * rowCount * rowCount; //because we can have up to three legs for a flight, we must set the array to the number of flights we can have
            flight = new int[power, 3]; //two dimensional array of potential flights
            flightLegs = new int[power]; //use to record how many lets each flight has
            flightType = new string[power]; //hold records as to whether a flight is to or return
        }

        public class flightItem
        { //class used to insert flights into the data grid
            public string Type { get; set; }
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
            Warning.Text = ""; //clear warning
            //clear the grid and the array
            Flights.Items.Refresh();
            Flights.Items.Clear();
            numOfFlights = 0;
            //get access to functions

            //get the origin, destination and dates
            origin = functions.getAirportCode(Start.Text);
            destination = functions.getAirportCode(End.Text);

            if ((!(origin == " " || destination == " " || Departure.Text == " " || Arrival.Text == " " || Departure.Text == "Select a date" || Arrival.Text == "Select a date")) && Arrival.SelectedDate >= DateTime.Today && Departure.SelectedDate >= DateTime.Today)
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

                int attendance, attendance2, attendance3; //the attendance of the flight
                string plane, plane2, plane3; //the plane ID
                DateTime foundDate, foundDate2, foundDate3, foundEndDate, foundEndDate2, foundEndDate3, earliestEndDate;
                double foundTime, foundTime2, foundTime3, earliestEndTime;
                earliestEndDate = DateTime.Today;
                earliestEndTime = 0;


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
                            if (functions.fullFlight(attendance, plane) == false)
                            { //if the flight isn't full and the flight isn't already in the array
                                flight[numOfFlights, 0] = i; //save the index of the flight
                                flight[numOfFlights, 1] = 0; //we don't have a second leg, so save that as zero
                                flight[numOfFlights, 2] = 0; //we don't have a third leg, so save that as zero
                                flightLegs[numOfFlights] = 1;
                                flightType[numOfFlights] = "To";
                                numOfFlights++; //increment the number of flights

                                if (numOfFlights == 1)
                                { //set the initial value of earliestEndTime and earliestEndDate
                                    earliestEndTime = Convert.ToDouble(xlRange.Cells[i, 11].Value2.ToString());
                                    earliestEndDate = DateTime.FromOADate(xlRange.Cells[i, 10].Value2);
                                }
                                else if (numOfFlights > 1 && (foundEndDate < earliestEndDate) && (foundTime < earliestEndTime))
                                { //set the new earliest end time and date
                                    earliestEndTime = Convert.ToDouble(xlRange.Cells[i, 11].Value2.ToString());
                                    earliestEndDate = DateTime.FromOADate(xlRange.Cells[i, 10].Value2);
                                }
                            }
                        }
                        else if (xlRange.Cells[i, 5].Value2.ToString() == origin && xlRange.Cells[i, 6].Value2.ToString() != destination)
                        {
                            string leg = xlRange.Cells[i, 6].Value2.ToString(); //save the value of the leg
                            foundTime = Convert.ToDouble(xlRange.Cells[i, 11].Value2.ToString()); //save the end time and date of the first flight
                            foundEndDate = DateTime.FromOADate(xlRange.Cells[i, 10].Value2);
                            attendance = Int32.Parse(xlRange.Cells[i, 16].Value2.ToString()); //along with the plane and attendance
                            plane = xlRange.Cells[i, 14].Value2.ToString();
                            if (functions.fullFlight(attendance, plane) == false)
                            { //if the plane isn't full, begin looking for a second leg
                                for (int j = 2; j <= rowCount; j++)
                                { //go through the list again looking for 2nd legs
                                    foundDate2 = DateTime.FromOADate(xlRange.Cells[j, 7].Value2); //get the date of the 2nd flight
                                    foundTime2 = Convert.ToDouble(xlRange.Cells[j, 8].Value2.ToString()); //get the start time as well
                                    if (foundDate2 >= startDate && foundDate2 <= endDate && (foundDate2 > foundEndDate || (foundDate2 == foundEndDate && (foundTime2 - 2.0 / 72) >= foundTime)))
                                    { //if the flight takes place between the start and end date and is at least forty minutes after the first leg flight
                                        if (xlRange.Cells[j, 5].Value2.ToString() == leg && xlRange.Cells[j, 6].Value2.ToString() == destination)
                                        { //if the leg and destination match
                                            foundTime2 = Convert.ToDouble(xlRange.Cells[j, 11].Value2.ToString()); //find the end date and time
                                            foundEndDate2 = DateTime.FromOADate(xlRange.Cells[j, 10].Value2);
                                            attendance2 = Int32.Parse(xlRange.Cells[j, 16].Value2.ToString()); //along with the attendance and plane
                                            plane2 = xlRange.Cells[j, 14].Value2.ToString();
                                            if (functions.fullFlight(attendance2, plane2) == false)
                                            { //if the second plane isn't full
                                                flight[numOfFlights, 0] = i; //save the index of the flight
                                                flight[numOfFlights, 1] = j;
                                                flight[numOfFlights, 2] = 0; //we don't have a third leg, so save that as zero
                                                flightLegs[numOfFlights] = 2;
                                                flightType[numOfFlights] = "To";
                                                numOfFlights++; //increment the number of flights
                                                if (numOfFlights == 1)
                                                { //set the initial value of earliestEndTime and earliestEndDate
                                                    earliestEndTime = Convert.ToDouble(xlRange.Cells[j, 11].Value2.ToString());
                                                    earliestEndDate = DateTime.FromOADate(xlRange.Cells[j, 10].Value2);
                                                }
                                                else if (numOfFlights > 1 && (foundEndDate2 < earliestEndDate) && (foundTime2 < earliestEndTime))
                                                { //set the new earliest end time and date
                                                    earliestEndTime = Convert.ToDouble(xlRange.Cells[j, 11].Value2.ToString());
                                                    earliestEndDate = DateTime.FromOADate(xlRange.Cells[j, 10].Value2);
                                                }
                                            }
                                        }
                                        else if (xlRange.Cells[j, 5].Value2.ToString() == leg && xlRange.Cells[j, 6].Value2.ToString() != destination)
                                        {
                                            string leg2 = xlRange.Cells[j, 6].Value2.ToString(); //save the value of the leg
                                            foundTime2 = Convert.ToDouble(xlRange.Cells[j, 11].Value2.ToString()); //save the end time and date of the first flight
                                            foundEndDate2 = DateTime.FromOADate(xlRange.Cells[j, 10].Value2);
                                            attendance2 = Int32.Parse(xlRange.Cells[j, 16].Value2.ToString()); //along with the plane and attendance
                                            plane2 = xlRange.Cells[j, 14].Value2.ToString();
                                            if (functions.fullFlight(attendance2, plane2) == false)
                                            { //if the plane isn't full, begin looking for a second leg
                                                for (int k = 2; k <= rowCount; k++)
                                                { //go through the list again looking for 2nd legs
                                                    foundDate3 = DateTime.FromOADate(xlRange.Cells[k, 7].Value2); //get the date of the 2nd flight
                                                    foundTime3 = Convert.ToDouble(xlRange.Cells[k, 8].Value2.ToString()); //get the start time as well
                                                    if (foundDate3 >= startDate && foundDate3 <= endDate && (foundDate3 > foundEndDate2 || (foundDate3 == foundEndDate2 && (foundTime3 - 2.0 / 72) >= foundTime2)))
                                                    { //if the flight takes place between the start and end date and is at least forty minutes after the first leg flight
                                                        if (xlRange.Cells[k, 5].Value2.ToString() == leg2 && xlRange.Cells[k, 6].Value2.ToString() == destination && leg2 != origin)
                                                        { //if the leg and destination match and the leg is not the origin
                                                            foundTime3 = Convert.ToDouble(xlRange.Cells[k, 11].Value2.ToString()); //find the end date and time
                                                            foundEndDate3 = DateTime.FromOADate(xlRange.Cells[k, 10].Value2);
                                                            attendance3 = Int32.Parse(xlRange.Cells[k, 16].Value2.ToString()); //along with the attendance and plane
                                                            plane3 = xlRange.Cells[k, 14].Value2.ToString();
                                                            if (functions.fullFlight(attendance3, plane3) == false)
                                                            { //if the second plane isn't full
                                                                flight[numOfFlights, 0] = i; //save the index of the flight
                                                                flight[numOfFlights, 1] = j;
                                                                flight[numOfFlights, 2] = k;
                                                                flightLegs[numOfFlights] = 3;
                                                                flightType[numOfFlights] = "To";
                                                                numOfFlights++; //increment the number of flights
                                                                if (numOfFlights == 1)
                                                                { //set the initial value of earliestEndTime and earliestEndDate
                                                                    earliestEndTime = Convert.ToDouble(xlRange.Cells[k, 11].Value2.ToString());
                                                                    earliestEndDate = DateTime.FromOADate(xlRange.Cells[k, 10].Value2);
                                                                }
                                                                else if (numOfFlights > 1 && (foundEndDate3 < earliestEndDate) && (foundTime3 < earliestEndTime))
                                                                { //set the new earliest end time and date
                                                                    earliestEndTime = Convert.ToDouble(xlRange.Cells[k, 11].Value2.ToString());
                                                                    earliestEndDate = DateTime.FromOADate(xlRange.Cells[k, 10].Value2);
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }


                int toFlights = numOfFlights;
                if ((bool)RoundTrip.IsChecked == true)
                { //if round trip is checked, loop back through and flights back that take place after the earlies flight to
                    if ((!(Round.Text == " " || Return.Text == " " || Round.Text == "Select a date" || Return.Text == "Select a date")) && Round.SelectedDate >= DateTime.Today && Return.SelectedDate >= DateTime.Today)
                    { //make sure that the round trip dates had values and are set to after the current date
                        DateTime roundDate = Convert.ToDateTime(Round.Text);
                        DateTime returnDate = Convert.ToDateTime(Return.Text);
                        for (int i = 2; i <= rowCount; i++)
                        { //Find the flights going to and from the origin and destination
                            //string userType = xlRange.Cells[IDcolumn, 2].Value2.ToString(); //get the user type from the database
                            //origin is 5, destination is 6, date is 7th row
                            foundDate = DateTime.FromOADate(xlRange.Cells[i, 7].Value2); //get the date of the flight
                            foundTime = Convert.ToDouble(xlRange.Cells[i, 8].Value2.ToString());
                            if (foundDate >= roundDate && foundDate <= returnDate && (foundDate > earliestEndDate || (foundDate == earliestEndDate && foundTime - 0.0278 > earliestEndTime)))
                            { //if the flight takes place between the start and end date
                                if (xlRange.Cells[i, 5].Value2.ToString() == destination && xlRange.Cells[i, 6].Value2.ToString() == origin)
                                { //if the origin and destination match
                                    foundTime = Convert.ToDouble(xlRange.Cells[i, 11].Value2.ToString()); //find the end date and time
                                    foundEndDate = DateTime.FromOADate(xlRange.Cells[i, 10].Value2);
                                    attendance = Int32.Parse(xlRange.Cells[i, 16].Value2.ToString()); //along with the attendance and plane
                                    plane = xlRange.Cells[i, 14].Value2.ToString();
                                    if (functions.fullFlight(attendance, plane) == false)
                                    { //if the flight isn't full and the flight isn't already in the array
                                        flight[numOfFlights, 0] = i; //save the index of the flight
                                        flight[numOfFlights, 1] = 0; //we don't have a second leg, so save that as zero
                                        flight[numOfFlights, 2] = 0; //we don't have a third leg, so save that as zero
                                        flightLegs[numOfFlights] = 1;
                                        flightType[numOfFlights] = "Round";
                                        numOfFlights++; //increment the number of flights
                                        /*
                                        if (numOfFlights == 1)
                                        { //set the initial value of earliestEndTime and earliestEndDate
                                            earliestEndTime = Convert.ToDouble(xlRange.Cells[i, 11].Value2.ToString());
                                            earliestEndDate = DateTime.FromOADate(xlRange.Cells[i, 10].Value2);
                                        }
                                        else if (numOfFlights > 1 && (foundEndDate < earliestEndDate) && (foundTime < earliestEndTime))
                                        { //set the new earliest end time and date
                                            earliestEndTime = Convert.ToDouble(xlRange.Cells[i, 11].Value2.ToString());
                                            earliestEndDate = DateTime.FromOADate(xlRange.Cells[i, 10].Value2);
                                        }
                                        */
                                    }
                                }
                                else if (xlRange.Cells[i, 5].Value2.ToString() == destination && xlRange.Cells[i, 6].Value2.ToString() != origin)
                                {
                                    string leg = xlRange.Cells[i, 6].Value2.ToString(); //save the value of the leg
                                    foundTime = Convert.ToDouble(xlRange.Cells[i, 11].Value2.ToString()); //save the end time and date of the first flight
                                    foundEndDate = DateTime.FromOADate(xlRange.Cells[i, 10].Value2);
                                    attendance = Int32.Parse(xlRange.Cells[i, 16].Value2.ToString()); //along with the plane and attendance
                                    plane = xlRange.Cells[i, 14].Value2.ToString();
                                    if (functions.fullFlight(attendance, plane) == false)
                                    { //if the plane isn't full, begin looking for a second leg
                                        for (int j = 2; j <= rowCount; j++)
                                        { //go through the list again looking for 2nd legs
                                            foundDate2 = DateTime.FromOADate(xlRange.Cells[j, 7].Value2); //get the date of the 2nd flight
                                            foundTime2 = Convert.ToDouble(xlRange.Cells[j, 8].Value2.ToString()); //get the start time as well
                                            if (foundDate2 >= roundDate && foundDate2 <= returnDate && (foundDate2 > foundEndDate || (foundDate2 == foundEndDate && (foundTime2 - 2.0/72) >= foundTime)))
                                            { //if the flight takes place between the start and end date and is at least forty minutes after the first leg flight
                                                if (xlRange.Cells[j, 5].Value2.ToString() == leg && xlRange.Cells[j, 6].Value2.ToString() == origin)
                                                { //if the leg and destination match
                                                    foundTime2 = Convert.ToDouble(xlRange.Cells[j, 11].Value2.ToString()); //find the end date and time
                                                    foundEndDate2 = DateTime.FromOADate(xlRange.Cells[j, 10].Value2);
                                                    attendance2 = Int32.Parse(xlRange.Cells[j, 16].Value2.ToString()); //along with the attendance and plane
                                                    plane2 = xlRange.Cells[j, 14].Value2.ToString();
                                                    if (functions.fullFlight(attendance2, plane2) == false)
                                                    { //if the second plane isn't full
                                                        flight[numOfFlights, 0] = i; //save the index of the flight
                                                        flight[numOfFlights, 1] = j;
                                                        flight[numOfFlights, 2] = 0; //we don't have a third leg, so save that as zero
                                                        flightLegs[numOfFlights] = 2;
                                                        flightType[numOfFlights] = "Round";
                                                        numOfFlights++; //increment the number of flights
                                                        /*
                                                        if (numOfFlights == 1)
                                                        { //set the initial value of earliestEndTime and earliestEndDate
                                                            earliestEndTime = Convert.ToDouble(xlRange.Cells[j, 11].Value2.ToString());
                                                            earliestEndDate = DateTime.FromOADate(xlRange.Cells[j, 10].Value2);
                                                        }
                                                        else if (numOfFlights > 1 && (foundEndDate2 < earliestEndDate) && (foundTime2 < earliestEndTime))
                                                        { //set the new earliest end time and date
                                                            earliestEndTime = Convert.ToDouble(xlRange.Cells[j, 11].Value2.ToString());
                                                            earliestEndDate = DateTime.FromOADate(xlRange.Cells[j, 10].Value2);
                                                        }
                                                        */
                                                    }
                                                }
                                                else if (xlRange.Cells[j, 5].Value2.ToString() == leg && xlRange.Cells[j, 6].Value2.ToString() != origin)
                                                {
                                                    string leg2 = xlRange.Cells[j, 6].Value2.ToString(); //save the value of the leg
                                                    foundTime2 = Convert.ToDouble(xlRange.Cells[j, 11].Value2.ToString()); //save the end time and date of the first flight
                                                    foundEndDate2 = DateTime.FromOADate(xlRange.Cells[j, 10].Value2);
                                                    attendance2 = Int32.Parse(xlRange.Cells[j, 16].Value2.ToString()); //along with the plane and attendance
                                                    plane2 = xlRange.Cells[j, 14].Value2.ToString();
                                                    if (functions.fullFlight(attendance2, plane2) == false)
                                                    { //if the plane isn't full, begin looking for a second leg
                                                        for (int k = 2; k <= rowCount; k++)
                                                        { //go through the list again looking for 2nd legs
                                                            foundDate3 = DateTime.FromOADate(xlRange.Cells[k, 7].Value2); //get the date of the 2nd flight
                                                            foundTime3 = Convert.ToDouble(xlRange.Cells[k, 8].Value2.ToString()); //get the start time as well
                                                            if (foundDate3 >= roundDate && foundDate3 <= returnDate && (foundDate3 > foundEndDate2 || (foundDate3 == foundEndDate2 && (foundTime3 - 2.0/72) >= foundTime2)))
                                                            { //if the flight takes place between the start and end date and is at least forty minutes after the first leg flight
                                                                if (xlRange.Cells[k, 5].Value2.ToString() == leg2 && xlRange.Cells[k, 6].Value2.ToString() == origin && leg2 != destination)
                                                                { //if the leg and destination match and the leg is not the origin
                                                                    foundTime3 = Convert.ToDouble(xlRange.Cells[k, 11].Value2.ToString()); //find the end date and time
                                                                    foundEndDate3 = DateTime.FromOADate(xlRange.Cells[k, 10].Value2);
                                                                    attendance3 = Int32.Parse(xlRange.Cells[k, 16].Value2.ToString()); //along with the attendance and plane
                                                                    plane3 = xlRange.Cells[k, 14].Value2.ToString();
                                                                    if (functions.fullFlight(attendance3, plane3) == false)
                                                                    { //if the second plane isn't full
                                                                        flight[numOfFlights, 0] = i; //save the index of the flight
                                                                        flight[numOfFlights, 1] = j;
                                                                        flight[numOfFlights, 2] = k;
                                                                        flightLegs[numOfFlights] = 3;
                                                                        flightType[numOfFlights] = "Round";
                                                                        numOfFlights++; //increment the number of flights
                                                                        /*
                                                                        if (numOfFlights == 1)
                                                                        { //set the initial value of earliestEndTime and earliestEndDate
                                                                            earliestEndTime = Convert.ToDouble(xlRange.Cells[k, 11].Value2.ToString());
                                                                            earliestEndDate = DateTime.FromOADate(xlRange.Cells[k, 10].Value2);
                                                                        }
                                                                        else if (numOfFlights > 1 && (foundEndDate3 < earliestEndDate) && (foundTime3 < earliestEndTime))
                                                                        { //set the new earliest end time and date
                                                                            earliestEndTime = Convert.ToDouble(xlRange.Cells[k, 11].Value2.ToString());
                                                                            earliestEndDate = DateTime.FromOADate(xlRange.Cells[k, 10].Value2);
                                                                        }
                                                                        */
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
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
                        if (flightLegs[i] == 1)
                        {
                            var item = new flightItem
                            {
                                Type = flightType[i],
                                ID = i.ToString(),
                                Origin = functions.getAirport(xlRange.Cells[flight[i, 0], 5].Value2.ToString()),
                                Destination = functions.getAirport(xlRange.Cells[flight[i, 0], 6].Value2.ToString()),
                                Departure = DateTime.FromOADate(xlRange.Cells[flight[i, 0], 7].Value2).ToString("MM/dd/yyyy") + " " + DateTime.FromOADate(xlRange.Cells[flight[i, 0], 8].Value2).ToString("h:mm tt"),
                                Arrival = DateTime.FromOADate(xlRange.Cells[flight[i, 0], 10].Value2).ToString("MM/dd/yyyy") + " " + DateTime.FromOADate(xlRange.Cells[flight[i, 0], 11].Value2).ToString("h:mm tt"),
                                Price = "$" + xlRange.Cells[flight[i, 0], 17].Value2.ToString()
                            }; //create a new flight item to insert into the data grid
                            Flights.Items.Add(item);
                        }
                        else if (flightLegs[i] == 2)
                        {
                            var item = new flightItem
                            {
                                Type = flightType[i],
                                ID = i.ToString(),
                                Origin = functions.getAirport(xlRange.Cells[flight[i, 0], 5].Value2.ToString()),
                                Destination = functions.getAirport(xlRange.Cells[flight[i, 1], 6].Value2.ToString()),
                                Departure = DateTime.FromOADate(xlRange.Cells[flight[i, 0], 7].Value2).ToString("MM/dd/yyyy") + " " + DateTime.FromOADate(xlRange.Cells[flight[i, 0], 8].Value2).ToString("h:mm tt"),
                                Arrival = DateTime.FromOADate(xlRange.Cells[flight[i, 1], 10].Value2).ToString("MM/dd/yyyy") + " " + DateTime.FromOADate(xlRange.Cells[flight[i, 1], 11].Value2).ToString("h:mm tt"),
                                Price = "$" + (Double.Parse(xlRange.Cells[flight[i, 0], 17].Value2.ToString()) + Double.Parse(xlRange.Cells[flight[i, 1], 17].Value2.ToString()) + 8).ToString()
                            }; //create a new flight item to insert into the data grid
                            Flights.Items.Add(item);
                        }
                        else if (flightLegs[i] == 3)
                        {
                            var item = new flightItem
                            {
                                Type = flightType[i],
                                ID = i.ToString(),
                                Origin = functions.getAirport(xlRange.Cells[flight[i, 0], 5].Value2.ToString()),
                                Destination = functions.getAirport(xlRange.Cells[flight[i, 2], 6].Value2.ToString()),
                                Departure = DateTime.FromOADate(xlRange.Cells[flight[i, 0], 7].Value2).ToString("MM/dd/yyyy") + " " + DateTime.FromOADate(xlRange.Cells[flight[i, 0], 8].Value2).ToString("h:mm tt"),
                                Arrival = DateTime.FromOADate(xlRange.Cells[flight[i, 2], 10].Value2).ToString("MM/dd/yyyy") + " " + DateTime.FromOADate(xlRange.Cells[flight[i, 2], 11].Value2).ToString("h:mm tt"),
                                Price = "$" + (Double.Parse(xlRange.Cells[flight[i, 0], 17].Value2.ToString()) + Double.Parse(xlRange.Cells[flight[i, 1], 17].Value2.ToString()) + Double.Parse(xlRange.Cells[flight[i, 2], 17].Value2.ToString()) + 16).ToString()
                            }; //create a new flight item to insert into the data grid
                            Flights.Items.Add(item);
                        }

                    }
                }
                xlWorkbook.Close();
                if (((bool)RoundTrip.IsChecked == true) && numOfFlights == toFlights)
                { //if we were searching for round trips and only found flights to our destination
                    Warning.Text = "Did not find any return flights"; //print that we didn't find any return flights
                }

            }
            else
            { //if one of them has not been set
                Warning.Text = "Assign values for origin, destination, and departure and arrival dates that are after today";
            }
        }
        private void Submit_Click(object sender, RoutedEventArgs e)
        { //Go back to the main menu
            bool isNumber = int.TryParse(FlightID.Text, out flightID);
            if (isNumber == true)
            { //if the entered ID is a number
                if (flightID < 0 || flightID >= numOfFlights)
                { //if the ID entered is one of the ID's search returned
                    Warning.Text = "Please enter one of the numbers listed for the search";

                }
                else
                {
                    BookFlight bookFlight = new BookFlight(identification, flight[flightID, 0], flight[flightID, 1], flight[flightID, 2]);
                    this.NavigationService.Navigate(bookFlight);
                }
            }
            else
            { //if ID is not a number, print a warning
                Warning.Text = "Please enter one of the numbers listed for the search";
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
