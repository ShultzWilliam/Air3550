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
        int flightIDRow1, flightIDRow2, flightIDRow3, numberOfLegs;
        double fullPrice;
        public BookFlight()
        {
            InitializeComponent();
        }
        public BookFlight(string identification, int row1, int row2, int row3)
        { //define the flight and user IDs
            InitializeComponent();
            flightIDRow1 = row1;
            flightIDRow2 = row2;
            flightIDRow3 = row3;
            Identification = identification;
        }
        private void Window_Loaded(object sender, RoutedEventArgs e)
        { //when the window is loaded, load in the flight info

            //get the necessary excel variables
            Excel.Workbook xlWorkbook = functions.database_connect();
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[2];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            //int rowCount = functions.getRows(2);
            //int IDRow = functions.getIDRow(flightID, 2);
            //Excel.Workbook xlWorkbook2 = functions.database_connect();
            Excel._Worksheet xlWorksheet2 = xlWorkbook.Sheets[1];
            Excel.Range xlRange2 = xlWorksheet2.UsedRange;
            //int rowCount2 = functions.getRows(1);
            int userIDRow = functions.getIDRow(Identification, 1);
            //set the multi leg values to none initially
            SecondLeg.Text = "None";
            ThirdLeg.Text = "None";
            Arrival_Date2.Text = "None";
            Arrival_Time2.Text = "None";
            Arrival_Terminal2.Text = "None";
            Departure_Date2.Text = "None";
            Departure_Time2.Text = "None";
            Departure_Terminal2.Text = "None";
            Arrival_Date3.Text = "None";
            Arrival_Time3.Text = "None";
            Arrival_Terminal3.Text = "None";
            Departure_Date3.Text = "None";
            Departure_Time3.Text = "None";
            Departure_Terminal3.Text = "None";
            Plane2.Text = "None";
            Plane3.Text = "None";

            if (flightIDRow2 == 0 && flightIDRow3 == 0)
            {
                numberOfLegs = 1;
                FlightID.Text = xlRange.Cells[flightIDRow1, 1].Value2.ToString();
                Origin.Text = functions.getAirport(xlRange.Cells[flightIDRow1, 5].Value2.ToString());
                Destination.Text = functions.getAirport(xlRange.Cells[flightIDRow1, 6].Value2.ToString());
                Departure_Date.Text = (DateTime.FromOADate(xlRange.Cells[flightIDRow1, 7].Value2)).ToString("MM/dd/yyyy");
                Departure_Time.Text = (DateTime.FromOADate(xlRange.Cells[flightIDRow1, 8].Value2)).ToString("h:mm tt");
                Departure_Terminal.Text = xlRange.Cells[flightIDRow1, 9].Value2.ToString();
                Arrival_Date.Text = (DateTime.FromOADate(xlRange.Cells[flightIDRow1, 10].Value2)).ToString("MM/dd/yyyy");
                Arrival_Time.Text = (DateTime.FromOADate(xlRange.Cells[flightIDRow1, 11].Value2)).ToString("h:mm tt");
                Arrival_Terminal.Text = xlRange.Cells[flightIDRow1, 12].Value2.ToString();
                Price.Text = "$" + xlRange.Cells[flightIDRow1, 17].Value2.ToString();
                Plane.Text = xlRange.Cells[flightIDRow1, 14].Value2.ToString();
                Credits.Text = xlRange2.Cells[userIDRow, 16].Value2.ToString();
                Point.Text = xlRange2.Cells[userIDRow, 17].Value2.ToString();
                fullPrice = Convert.ToDouble(xlRange.Cells[flightIDRow1, 17].Value2.ToString());
            }
            else if (flightIDRow2 != 0 && flightIDRow3 == 0)
            {
                numberOfLegs = 2;
                FlightID.Text = xlRange.Cells[flightIDRow1, 1].Value2.ToString() + ", " + xlRange.Cells[flightIDRow2, 1].Value2.ToString();
                Origin.Text = functions.getAirport(xlRange.Cells[flightIDRow1, 5].Value2.ToString());
                Destination.Text = functions.getAirport(xlRange.Cells[flightIDRow2, 6].Value2.ToString());
                SecondLeg.Text = functions.getAirport(xlRange.Cells[flightIDRow1, 6].Value2.ToString());
                Departure_Date.Text = (DateTime.FromOADate(xlRange.Cells[flightIDRow1, 7].Value2)).ToString("MM/dd/yyyy");
                Departure_Time.Text = (DateTime.FromOADate(xlRange.Cells[flightIDRow1, 8].Value2)).ToString("h:mm tt");
                Departure_Terminal.Text = xlRange.Cells[flightIDRow1, 9].Value2.ToString();
                Arrival_Date2.Text = (DateTime.FromOADate(xlRange.Cells[flightIDRow1, 10].Value2)).ToString("MM/dd/yyyy");
                Arrival_Time2.Text = (DateTime.FromOADate(xlRange.Cells[flightIDRow1, 11].Value2)).ToString("h:mm tt");
                Arrival_Terminal2.Text = xlRange.Cells[flightIDRow1, 12].Value2.ToString();
                Departure_Date2.Text = (DateTime.FromOADate(xlRange.Cells[flightIDRow2, 7].Value2)).ToString("MM/dd/yyyy");
                Departure_Time2.Text = (DateTime.FromOADate(xlRange.Cells[flightIDRow2, 8].Value2)).ToString("h:mm tt");
                Departure_Terminal2.Text = xlRange.Cells[flightIDRow2, 9].Value2.ToString();
                Arrival_Date.Text = (DateTime.FromOADate(xlRange.Cells[flightIDRow2, 10].Value2)).ToString("MM/dd/yyyy");
                Arrival_Time.Text = (DateTime.FromOADate(xlRange.Cells[flightIDRow2, 11].Value2)).ToString("h:mm tt");
                Arrival_Terminal.Text = xlRange.Cells[flightIDRow2, 12].Value2.ToString();
                Price.Text = "$" + (Convert.ToDouble(xlRange.Cells[flightIDRow1, 17].Value2.ToString()) + 8 + Convert.ToDouble(xlRange.Cells[flightIDRow2, 17].Value2.ToString())).ToString();
                Plane.Text = xlRange.Cells[flightIDRow1, 14].Value2.ToString();
                Plane2.Text = xlRange.Cells[flightIDRow2, 14].Value2.ToString();
                Credits.Text = xlRange2.Cells[userIDRow, 16].Value2.ToString();
                Point.Text = xlRange2.Cells[userIDRow, 17].Value2.ToString();
                fullPrice = Convert.ToDouble(xlRange.Cells[flightIDRow1, 17].Value2.ToString()) + 8 + Convert.ToDouble(xlRange.Cells[flightIDRow2, 17].Value2.ToString());
            }
            else if (flightIDRow2 != 0 && flightIDRow3 != 0)
            {
                numberOfLegs = 3;
                FlightID.Text = xlRange.Cells[flightIDRow1, 1].Value2.ToString() + ", " + xlRange.Cells[flightIDRow2, 1].Value2.ToString() + ", " + xlRange.Cells[flightIDRow3, 1].Value2.ToString();
                Origin.Text = functions.getAirport(xlRange.Cells[flightIDRow1, 5].Value2.ToString());
                Destination.Text = functions.getAirport(xlRange.Cells[flightIDRow3, 6].Value2.ToString());
                SecondLeg.Text = functions.getAirport(xlRange.Cells[flightIDRow1, 6].Value2.ToString());
                ThirdLeg.Text = functions.getAirport(xlRange.Cells[flightIDRow2, 6].Value2.ToString());
                Departure_Date.Text = (DateTime.FromOADate(xlRange.Cells[flightIDRow1, 7].Value2)).ToString("MM/dd/yyyy");
                Departure_Time.Text = (DateTime.FromOADate(xlRange.Cells[flightIDRow1, 8].Value2)).ToString("h:mm tt");
                Departure_Terminal.Text = xlRange.Cells[flightIDRow1, 9].Value2.ToString();
                Arrival_Date2.Text = (DateTime.FromOADate(xlRange.Cells[flightIDRow1, 10].Value2)).ToString("MM/dd/yyyy");
                Arrival_Time2.Text = (DateTime.FromOADate(xlRange.Cells[flightIDRow1, 11].Value2)).ToString("h:mm tt");
                Arrival_Terminal2.Text = xlRange.Cells[flightIDRow1, 12].Value2.ToString();
                Departure_Date2.Text = (DateTime.FromOADate(xlRange.Cells[flightIDRow2, 7].Value2)).ToString("MM/dd/yyyy");
                Departure_Time2.Text = (DateTime.FromOADate(xlRange.Cells[flightIDRow2, 8].Value2)).ToString("h:mm tt");
                Departure_Terminal2.Text = xlRange.Cells[flightIDRow2, 9].Value2.ToString();
                Arrival_Date3.Text = (DateTime.FromOADate(xlRange.Cells[flightIDRow2, 10].Value2)).ToString("MM/dd/yyyy");
                Arrival_Time3.Text = (DateTime.FromOADate(xlRange.Cells[flightIDRow2, 11].Value2)).ToString("h:mm tt");
                Arrival_Terminal3.Text = xlRange.Cells[flightIDRow2, 12].Value2.ToString();
                Departure_Date3.Text = (DateTime.FromOADate(xlRange.Cells[flightIDRow3, 7].Value2)).ToString("MM/dd/yyyy");
                Departure_Time3.Text = (DateTime.FromOADate(xlRange.Cells[flightIDRow3, 8].Value2)).ToString("h:mm tt");
                Departure_Terminal3.Text = xlRange.Cells[flightIDRow3, 9].Value2.ToString();
                Arrival_Date.Text = (DateTime.FromOADate(xlRange.Cells[flightIDRow3, 10].Value2)).ToString("MM/dd/yyyy");
                Arrival_Time.Text = (DateTime.FromOADate(xlRange.Cells[flightIDRow3, 11].Value2)).ToString("h:mm tt");
                Arrival_Terminal.Text = xlRange.Cells[flightIDRow3, 12].Value2.ToString();
                Price.Text = "$" + (Convert.ToDouble(xlRange.Cells[flightIDRow1, 17].Value2.ToString()) + 8 + Convert.ToDouble(xlRange.Cells[flightIDRow2, 17].Value2.ToString()) + 8 + Convert.ToDouble(xlRange.Cells[flightIDRow3, 17].Value2.ToString())).ToString();
                Plane.Text = xlRange.Cells[flightIDRow1, 14].Value2.ToString();
                Plane2.Text = xlRange.Cells[flightIDRow2, 14].Value2.ToString();
                Plane3.Text = xlRange.Cells[flightIDRow3, 14].Value2.ToString();
                Credits.Text = xlRange2.Cells[userIDRow, 16].Value2.ToString();
                Point.Text = xlRange2.Cells[userIDRow, 17].Value2.ToString();
                fullPrice = Convert.ToDouble(xlRange.Cells[flightIDRow1, 17].Value2.ToString()) + 8 + Convert.ToDouble(xlRange.Cells[flightIDRow2, 17].Value2.ToString()) + 8 + Convert.ToDouble(xlRange.Cells[flightIDRow3, 17].Value2.ToString());
            }
            xlWorkbook.Close(true);
        }

        private void Book_Click(object sender, RoutedEventArgs e)
        { //to book the flight
            //define the excel variables to read from both the flight and user databases
            Excel.Workbook xlWorkbook = functions.database_connect();
            Excel._Worksheet xlWorksheet1 = xlWorkbook.Sheets[1]; //for the user
            Excel.Range xlRange1 = xlWorksheet1.UsedRange;
            int rowCount1 = functions.getRows(1);
            Excel._Worksheet xlWorksheet2 = xlWorkbook.Sheets[2]; //for the flight
            Excel.Range xlRange2 = xlWorksheet2.UsedRange;
            int rowCount2 = functions.getRows(2);
            string userFlights, flightPassengers, paymentMethods; //create strings to read and modify the user's flights and the flight's passengers
            int userIDRow = functions.getIDRow(Identification, 1); //get the ID Rows for the flight and user IDs
            int flightIDRow;
            double price;
            double credit = Double.Parse(Credits.Text);
            double points = Double.Parse(Point.Text);
            int[] flightIDRows = new int[3];
            flightIDRows[0] = flightIDRow1;
            flightIDRows[1] = flightIDRow2;
            flightIDRows[2] = flightIDRow3;
            string name = xlRange1.Cells[userIDRow, 3].Value2.ToString() + " " + xlRange1.Cells[userIDRow, 5].Value2.ToString();  //get the full name of the user

            if (functions.flightBooked(Identification, xlRange2.Cells[flightIDRow1, 1].Value2.ToString()) || (flightIDRow2 != 0 && functions.flightBooked(Identification, xlRange2.Cells[flightIDRow2, 1].Value2.ToString())) || (flightIDRow3 != 0 && functions.flightBooked(Identification, xlRange2.Cells[flightIDRow3, 1].Value2.ToString())))
            { //if the user is already booked for one of these flights, tell them
                Warning.Text = "You are already booked for one of these flights";
                xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU CHANGE
                xlWorkbook.Close(); //THIS ONE TOO
                //if ((bool)Credit.IsChecked == true && (bool)CreditCard.IsChecked == false && (bool)Points.IsChecked == false)
            }
            else if ((bool)Credit.IsChecked == true && credit < fullPrice)
            {
                Warning.Text = "You do not have enough credit";
                xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU CHANGE
                xlWorkbook.Close(); //THIS ONE TOO
            }
            else if ((bool)Points.IsChecked == true && points / 100 < fullPrice)
            {
                Warning.Text = "You do not have enough points";
                xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU CHANGE
                xlWorkbook.Close(); //THIS ONE TOO
            }
            else
            {
                if ((DateTime.Parse(Departure_Date.Text + " " + Departure_Time.Text)).AddMonths(-6) <= DateTime.Now && (DateTime.Parse(Arrival_Date.Text + " " + Arrival_Time.Text)).AddMonths(-6) <= DateTime.Now)
                { //if we are at or within six months of the flight
                    for (int i = 0; i < numberOfLegs; i++)
                    {
                        flightIDRow = flightIDRows[i];
                        flightID = xlRange2.Cells[flightIDRow, 1].Value2.ToString();
                        credit = Convert.ToDouble(xlRange1.Cells[userIDRow, 16].Value2.ToString()); //reset the flightIDRow, flightID, credit, and points
                        points = Convert.ToDouble(xlRange1.Cells[userIDRow, 17].Value2.ToString());

                        price = Double.Parse(xlRange2.Cells[flightIDRow, 17].Value2.ToString()); //save the price of the flight
                        if (i > 0)
                        { //if this is the second or third leg, increment the price for it
                            price = price + 8;
                        }

                        if ((bool)Credit.IsChecked == true && (bool)CreditCard.IsChecked == false && (bool)Points.IsChecked == false)
                        { //If the user is paying with in-app credit
                            if (credit < price)
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
                                    paymentMethods = "Credit";
                                    xlRange1.Cells[userIDRow, 20].Value = paymentMethods;
                                    xlRange1.Cells[userIDRow, 21].Value = price.ToString();
                                }
                                else
                                { //if the user does have some flights
                                    userFlights = xlRange1.Cells[userIDRow, 19].Value2.ToString() + " " + flightID; //read in the users flights
                                    xlRange1.Cells[userIDRow, 19].Value = userFlights; //add the new flight, and write those to the database
                                    paymentMethods = xlRange1.Cells[userIDRow, 20].Value2.ToString() + " " + "Credit";
                                    xlRange1.Cells[userIDRow, 20].Value = paymentMethods;
                                    xlRange1.Cells[userIDRow, 21].Value = xlRange1.Cells[userIDRow, 21].Value2.ToString() + " " + price.ToString();

                                }
                                if (functions.isEmpty(2, flightIDRow, 15))
                                { //if the flight does not have any passengers
                                    flightPassengers = Identification; //just save the flight ID in
                                    xlRange2.Cells[flightIDRow, 15].Value = flightPassengers; //add the new flight, and write those to the database
                                }
                                else
                                { //if the flight does
                                    flightPassengers = xlRange2.Cells[flightIDRow, 15].Value2.ToString() + " " + Identification; //read in the users flights
                                    xlRange2.Cells[flightIDRow, 15].Value = flightPassengers; //add the new flight, and write those to the database
                                }
                                xlRange1.Cells[userIDRow, 16].Value = (credit - price).ToString();

                                double profit = Double.Parse(xlRange2.Cells[flightIDRow, 18].Value2.ToString());
                                profit = profit + Double.Parse(xlRange2.Cells[flightIDRow, 17].Value2.ToString());
                                xlRange2.Cells[flightIDRow, 18].Value = profit.ToString(); //update the profit of the flight (should I do this for this one?)
                                double attendance = Double.Parse(xlRange2.Cells[flightIDRow, 16].Value2.ToString()) + 1;
                                xlRange2.Cells[flightIDRow, 16].Value = attendance.ToString(); //update the attendance as well

                                Excel._Worksheet xlWorksheet3 = xlWorkbook.Sheets[4]; //for the airport
                                Excel.Range xlRange3 = xlWorksheet3.UsedRange;
                                int rowCount3 = functions.getRows(4);
                                int airportIDRow = functions.getIDRow(xlRange2.Cells[flightIDRow, 5].Value2.ToString(), 4); //get the row for the origin airport
                                xlRange3.Cells[airportIDRow, 6].Value = (Double.Parse(xlRange3.Cells[airportIDRow, 6].Value2.ToString()) + price).ToString(); //increment the profit for the airport



                            }
                        }
                        else if ((bool)Credit.IsChecked == false && (bool)CreditCard.IsChecked == true && (bool)Points.IsChecked == false)
                        { //if the user is paying with credit card
                          //if the credit card is checked, we do not have to see if they have the necessary funds, just add them to the flight
                            if (functions.isEmpty(1, userIDRow, 19))
                            { //if the user does not have any flights booked
                                userFlights = flightID; //just save the flight ID in
                                xlRange1.Cells[userIDRow, 19].Value = userFlights; //add the new flight, and write those to the database
                                paymentMethods = "CreditCard";
                                xlRange1.Cells[userIDRow, 20].Value = paymentMethods;
                                xlRange1.Cells[userIDRow, 21].Value = price.ToString();

                            }
                            else
                            { //if the user does have some flights
                                userFlights = xlRange1.Cells[userIDRow, 19].Value2.ToString() + " " + flightID; //read in the users flights
                                xlRange1.Cells[userIDRow, 19].Value = userFlights; //add the new flight, and write those to the database
                                paymentMethods = xlRange1.Cells[userIDRow, 20].Value2.ToString() + " " + "CreditCard";
                                xlRange1.Cells[userIDRow, 20].Value = paymentMethods;
                                xlRange1.Cells[userIDRow, 21].Value = xlRange1.Cells[userIDRow, 21].Value2.ToString() + " " + price.ToString();

                            }
                            if (functions.isEmpty(2, flightIDRow, 15))
                            { //if the flight does not have any passengers
                                flightPassengers = Identification; //just save the flight ID in
                                xlRange2.Cells[flightIDRow, 15].Value = flightPassengers; //add the new flight, and write those to the database
                            }
                            else
                            { //if the flight does
                                flightPassengers = xlRange2.Cells[flightIDRow, 15].Value2.ToString() + " " + Identification; //read in the users flights
                                xlRange2.Cells[flightIDRow, 15].Value = flightPassengers; //add the new flight, and write those to the database
                            }

                            double moneySpent = Double.Parse(xlRange1.Cells[userIDRow, 18].Value2.ToString());
                            moneySpent = moneySpent + price;
                            xlRange1.Cells[userIDRow, 18].Value = moneySpent.ToString(); //update the amount of money the user has spent

                            double profit = Double.Parse(xlRange2.Cells[flightIDRow, 18].Value2.ToString());
                            profit = profit + Double.Parse(xlRange2.Cells[flightIDRow, 17].Value2.ToString());
                            xlRange2.Cells[flightIDRow, 18].Value = profit.ToString(); //update the profit of the flight
                            double attendance = Double.Parse(xlRange2.Cells[flightIDRow, 16].Value2.ToString()) + 1;
                            xlRange2.Cells[flightIDRow, 16].Value = attendance.ToString(); //update the attendance as well

                            //(do they only earn points after the flight takes off?)
                            //double points = Double.Parse(xlRange1.Cells[userIDRow, 17].Value2.ToString()) + double.Parse(Price.Text) / 10;
                            //xlRange1.Cells[userIDRow, 17].Value = points.ToString(); //give them points for their payment

                            Excel._Worksheet xlWorksheet4 = xlWorkbook.Sheets[5]; //save a record of the financial transaction
                            Excel.Range xlRange4 = xlWorksheet4.UsedRange;
                            int rowCount4 = functions.getRows(5);
                            xlRange4.Cells[rowCount4 + 1, 1].Value = xlRange1.Cells[userIDRow, 3].Value2.ToString() + " " + xlRange1.Cells[userIDRow, 4].Value2.ToString() + " " + xlRange1.Cells[userIDRow, 5].Value2.ToString();
                            xlRange4.Cells[rowCount4 + 1, 2].Value = xlRange1.Cells[userIDRow, 13].Value2.ToString();
                            xlRange4.Cells[rowCount4 + 1, 3].Value = price;


                            Excel._Worksheet xlWorksheet3 = xlWorkbook.Sheets[4]; //for the airport
                            Excel.Range xlRange3 = xlWorksheet3.UsedRange;
                            int rowCount3 = functions.getRows(4);
                            int airportIDRow = functions.getIDRow(xlRange2.Cells[flightIDRow, 5].Value2.ToString(), 4); //get the row for the origin airport
                            xlRange3.Cells[airportIDRow, 6].Value = (Double.Parse(xlRange3.Cells[airportIDRow, 6].Value2.ToString()) + price).ToString(); //increment the profit for the airport



                        }
                        else if ((bool)Credit.IsChecked == false && (bool)CreditCard.IsChecked == false && (bool)Points.IsChecked == true)
                        { //if the user is paying with points

                            if (points / 100 < price)
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
                                    paymentMethods = "Points";
                                    xlRange1.Cells[userIDRow, 20].Value = paymentMethods;
                                    xlRange1.Cells[userIDRow, 21].Value = price.ToString();

                                }
                                else
                                { //if the user does have some flights
                                    userFlights = xlRange1.Cells[userIDRow, 19].Value2.ToString() + " " + flightID; //read in the users flights
                                    xlRange1.Cells[userIDRow, 19].Value = userFlights; //add the new flight, and write those to the database
                                    paymentMethods = xlRange1.Cells[userIDRow, 20].Value2.ToString() + " " + "Points";
                                    xlRange1.Cells[userIDRow, 20].Value = paymentMethods;
                                    xlRange1.Cells[userIDRow, 21].Value = xlRange1.Cells[userIDRow, 21].Value2.ToString() + " " + price.ToString();
                                }
                                if (functions.isEmpty(2, flightIDRow, 15))
                                { //if the flight does not have any passengers
                                    flightPassengers = Identification; //just save the flight ID in
                                    xlRange2.Cells[flightIDRow, 15].Value = flightPassengers; //add the new flight, and write those to the database
                                }
                                else
                                { //if the flight does
                                    flightPassengers = xlRange2.Cells[flightIDRow, 15].Value2.ToString() + " " + Identification; //read in the users flights
                                    xlRange2.Cells[flightIDRow, 15].Value = flightPassengers; //add the new flight, and write those to the database
                                }

                                xlRange1.Cells[userIDRow, 17].Value = ((points / 100 - price) * 100).ToString();
                                double pointsUsed = Convert.ToDouble(xlRange1.Cells[userIDRow, 30].Value2.ToString());
                                xlRange1.Cells[userIDRow, 30].Value = (pointsUsed + price * 100).ToString(); //increment the points used

                                double profit = Double.Parse(xlRange2.Cells[flightIDRow, 18].Value2.ToString());
                                profit = profit + Double.Parse(xlRange2.Cells[flightIDRow, 17].Value2.ToString());
                                xlRange2.Cells[flightIDRow, 18].Value = profit.ToString(); //update the profit of the flight
                                double attendance = Double.Parse(xlRange2.Cells[flightIDRow, 16].Value2.ToString()) + 1;
                                xlRange2.Cells[flightIDRow, 16].Value = attendance.ToString(); //update the attendance as well

                                //(do they only earn points after the flight takes off?)
                                //double points = Double.Parse(xlRange1.Cells[userIDRow, 17].Value2.ToString()) + double.Parse(Price.Text) / 10;
                                //xlRange1.Cells[userIDRow, 17].Value = points.ToString(); //give them points for their payment



                                Excel._Worksheet xlWorksheet3 = xlWorkbook.Sheets[4]; //for the airport
                                Excel.Range xlRange3 = xlWorksheet3.UsedRange;
                                int rowCount3 = functions.getRows(4);
                                int airportIDRow = functions.getIDRow(xlRange2.Cells[flightIDRow, 5].Value2.ToString(), 4); //get the row for the origin airport
                                xlRange3.Cells[airportIDRow, 6].Value = (Double.Parse(xlRange3.Cells[airportIDRow, 6].Value2.ToString()) + price).ToString(); //increment the profit for the airport

                            }

                        }

                        else
                        { //otherwise, display a warning
                            xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU CHANGE
                            xlWorkbook.Close(); //THIS ONE TOO
                            Warning.Text = "Please select a payment method. Cannot select more than one payment method";
                            break; //break out of the loop
                        }
                        xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU CHANGE
                    }
                    xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU CHANGE
                    xlWorkbook.Close(); //THIS ONE TOO
                                        //MainMenuCustomer mainMenu = new MainMenuCustomer(Identification); //create a new main menu and go to it
                                        //this.NavigationService.Navigate(mainMenu);
                    SearchFlight searchFlight = new SearchFlight(Identification);
                    this.NavigationService.Navigate(searchFlight);

                }
                else
                { //otherwise, display a warning that the flight cannot be booked yet
                    xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU CHANGE
                    xlWorkbook.Close(); //THIS ONE TOO
                    Warning.Text = "Flights can only be booked within six months of the flight.";
                }
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

        private void Sign_Out(object sender, RoutedEventArgs e)
        { //to sign out
            SignIn signIn = new SignIn(); //create a new main menu and go to it
            this.NavigationService.Navigate(signIn);
        }
    }
}
