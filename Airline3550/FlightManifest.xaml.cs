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
    /// Interaction logic for FlightManifest.xaml
    /// Place where a flight manager can print a flight manifest or search for customer records
    /// </summary>
    public partial class FlightManifest : Page
    {
        string identification, IDCHECK, moreFlightID; //initialize global variables

        public FlightManifest()
        {
            InitializeComponent();
        }
        public FlightManifest(string id) : base()
        { //define the user and flight IDs
            InitializeComponent();
            identification = id;
        }

        private void Sign_Out(object sender, RoutedEventArgs e)
        { //sign out of the application
            SignIn signIn = new SignIn();
            this.NavigationService.Navigate(signIn);
        }

        private void Main_Menu(object sender, RoutedEventArgs e)
        { //to return to the main menu
            MainMenuFlightManager mainMenu = new MainMenuFlightManager(identification); //create a new main menu and go to it
            this.NavigationService.Navigate(mainMenu);
        }
        private void Print(object sender, RoutedEventArgs e)
        {
            IDCHECK = FlightID.Text;
            bool TakenOrNotTaken = true;

            //check if the flightID exists

            Functions functions = new Functions();
            if (functions.isNum(IDCHECK) == true)
            {
                // finding all the crew members and attendents
                Excel.Workbook xlWorkbook = functions.database_connect();
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[2];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                int numRows = functions.getRows(2);
                Crew.Text = "Crew members: ";
                for (int i = 2; i <= numRows; i++)
                {//find the flight ID row
                    string temp = xlRange.Cells[i, 1].Value2.ToString();

                    if (temp == FlightID.Text)
                    {
                        TakenOrNotTaken = xlRange.Cells[i, 2].Value2;
                        moreFlightID = temp;
                        string crew, attendants;
                        crew = xlRange.Cells[i, 3].Value2.ToString();
                        attendants = xlRange.Cells[i, 4].Value2.ToString();
                        Crew.Text += crew + ", " + attendants;
                    }
                }
                xlWorkbook.Close(true);


                if (!TakenOrNotTaken)
                {
                    //Jacob Added
                    //Reset flight variables in database
                    string Destination = functions.getLocation(functions.getDestination(FlightID.Text));
                    string PlaneID = functions.getPlaneID(FlightID.Text);
                    functions.setInFlight(FlightID.Text);
                    functions.setPlaneLocation(Destination, PlaneID);
                    functions.addPlaneToInventory(Destination, functions.getPlaneModel(PlaneID));
                }


                // finding all the customers
                Excel.Workbook xlWorkbook2 = functions.database_connect();
                Excel._Worksheet xlWorksheet2 = xlWorkbook2.Sheets[1];
                Excel.Range xlRange2 = xlWorksheet2.UsedRange;
                Passengers.Text = "Passengers: ";
                int numRows2 = functions.getRows(1);
                // Passengers.Text += numRows2;
                for (int i = 2; i < numRows2; i++)
                {// find who is attending the flight
                    if (functions.isEmpty(1, i, 19) == false)
                    { // if the user has any purchased flights
                        string flights = xlRange2.Cells[i, 19].Value2.ToString();
                        if (flights.Contains(FlightID.Text))
                        {   //if the user is scheduled for the flight

                            //double dollarAmount;
                            //int pointsEarned;
                            // get rid of the flight in the user.flights portion
                            // get rid of users.paid with portion for that flight
                            //put that flight into flight history on their profile
                            //string userFlights = xlRange2.Cells[i, 19].Value2.ToString();
                            //string paymentMethods = xlRange2.Cells[i, 20].Value2.ToString();
                            //string userPayment = xlRange2.Cells[i, 21].Value2.ToString();
                            //int userPoints = xlRange2.Cells[i, 17].Value2;
                            //string writtenToHistory = "";
                            //List<string> userflightsArray = userFlights.Split(' ').ToList();
                            //List<string> paymentMethodsArray = paymentMethods.Split(' ').ToList();
                            //List<string> userPaymentArray = userPayment.Split(' ').ToList();
                            //for (int a = 0; a < userflightsArray.Count; a++)
                            //{ //find the slots in user.flights and user.paidwith for that info
                            //    if (userflightsArray[a] == FlightID.Text)
                            //    { // we found the index, which should be the same for user.flights and user.paidwith
                            //       writtenToHistory = userflightsArray[a];
                            //        userflightsArray.RemoveAt(a);               // remove the flight from users.flights
                            //        paymentMethodsArray.RemoveAt(a);            // remove the payment method from users.paid With

                            //        dollarAmount = Convert.ToDouble(userPaymentArray.ElementAt(a));
                            //        dollarAmount = dollarAmount / 10;
                            //        pointsEarned = Convert.ToInt32(dollarAmount);   // find the amount of points to give to the customer
                            //        userPoints += pointsEarned;
                            //        xlRange2.Cells[i, 17].Value2 = userPoints;       //update the user.points

                            //        userPaymentArray.RemoveAt(a);               // remove the price from user.price
                            //        string updatedUserFlightInfo = string.Join(" ", userflightsArray);
                            //        string updatedUserPaidWithInfo = string.Join(" ", paymentMethodsArray);
                            //        string userFlightHistory = xlRange2.Cells[i, 22].Value2.toString(); // find the user's flight history
                            //       userFlightHistory += " " + writtenToHistory;    // add the newly taken flight to the history string
                            //        xlRange2.Cells[i, 19] = updatedUserFlightInfo;      // update user.Flights
                            //        xlRange2.Cells[i, 20] = updatedUserPaidWithInfo;    //update user.Paid_With
                            //        xlRange2.Cells[i, 22] = userFlightHistory;           //update user.History
                            //        xlWorkbook2.Application.ActiveWorkbook.Save();
                            //    }
                            //}
                            Passengers.Text += xlRange2.Cells[i, 3].Value2.ToString() + " " + xlRange2.Cells[i, 5].Value2.ToString() + ", ";
                        }
                    }
                    else if (functions.isEmpty(1, i, 22) == false)
                    { // if the user has any purchased flights
                        string flights = xlRange2.Cells[i, 22].Value2.ToString();
                        if (flights.Contains(FlightID.Text))
                        {   //if the user was scheduled for the flight
                            Passengers.Text += xlRange2.Cells[i, 3].Value2.ToString() + " " + xlRange2.Cells[i, 5].Value2.ToString() + ", ";

                        }

                    }
                }
                Passengers.Text = Passengers.Text.Substring(0, Passengers.Text.Length - 2);
                xlWorkbook2.Close(true);
            }
            else
            { //otherwise, display an error
                Warning.Text = "Invalid Flight ID";
            }
        }
    }
}
