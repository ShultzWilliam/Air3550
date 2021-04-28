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
    /// Interaction logic for SchedulePlane.xaml
    /// Place where a Marketing Manager can schedule planes for flights
    /// </summary>
    public partial class SchedulePlane : Page
    {
        string flightID, Identification; //initialize global variables
        int flightRow = 2;

        public SchedulePlane()
        {
            InitializeComponent();
        }
        public SchedulePlane(string identification, string id)
        { //define the flight and user ID
            InitializeComponent();
            flightID = id;
            Identification = identification;
        }
        private void Window_Loaded(object sender, RoutedEventArgs e)
        { //when the window is loaded, load in the flight info

            //load in the data from the database

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

                    flightRow = i;
                    string Departure, Arrival, DepartureTime, DepartureDate, ArrivalTime, ArrivalDate;

                    Departure = DateTime.FromOADate(xlRange.Cells[i, 7].Value2).ToString("MM/dd/yyyy") + " " + DateTime.FromOADate(xlRange.Cells[i, 8].Value2).ToString("h:mm tt");
                    Arrival = DateTime.FromOADate(xlRange.Cells[i, 10].Value2).ToString("MM/dd/yyyy") + " " + DateTime.FromOADate(xlRange.Cells[i, 11].Value2).ToString("h:mm tt");

                    int dCutOff = Departure.IndexOf(" ");
                    DepartureTime = Departure.Substring(dCutOff + 1);

                    dCutOff = Arrival.IndexOf(" ");
                    ArrivalTime = Arrival.Substring(dCutOff + 1);

                    dCutOff = Departure.IndexOf(" ");
                    DepartureDate = Departure.Substring(0, dCutOff);
                    //DateTime date = DateTime.Parse(DepartureDate);

                    dCutOff = Arrival.IndexOf(" ");
                    ArrivalDate = Arrival.Substring(0, dCutOff);


                    Origin.Text = functions.getLocation(xlRange.Cells[i, 5].Value2.ToString());
                    Destination.Text = functions.getLocation(xlRange.Cells[i, 6].Value2.ToString());
                    Departure_Date.Text = DepartureDate;
                    Departure_Time.Text = DepartureTime;
                    Departure_Terminal.Text = xlRange.Cells[i, 9].Value2.ToString();
                    Arrival_Date.Text = ArrivalDate;
                    Arrival_Time.Text = ArrivalTime;                   
                    Arrival_Terminal.Text = xlRange.Cells[i, 12].Value2.ToString();
                    Price.Text = xlRange.Cells[i, 17].Value2.ToString();


                    //xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU OPEN
                    xlWorkbook.Close(); //THIS ONE TOO
                    return;
                }

            }

            Warning.Text = "Database Error";
            return;
        }
        private void Schedule_Click(object sender, RoutedEventArgs e)
        { //to book the flight

            //schedule the flight
            Functions functions = new Functions();

            if (Plane.Text != "737" && Plane.Text != "747" && Plane.Text != "767")
            {
                Warning.Text = "Please enter a valid plane model (737, 747, 767)";
                return;
            }

            string PlaneStatus = functions.isPlaneAvailableAndRemove(Origin.Text, Plane.Text);

            if (PlaneStatus == "FOUND")
            {//Add Plane to flightID

                string sPilots, sAttendant, sPlaneID;
                string sCrewPlaneID = functions.getCrewMeHartiesAndPlaneID(Origin.Text, Plane.Text);

                if (String.IsNullOrEmpty(sCrewPlaneID))
                {
                    Warning.Text = "Database Error";
                    return;
                }

                int CrewDivider = sCrewPlaneID.LastIndexOf(",");
                int PlaneDivider = sCrewPlaneID.LastIndexOf(" ");
                sPilots = sCrewPlaneID.Substring(0, CrewDivider);
                sAttendant = sCrewPlaneID.Substring(CrewDivider + 2, PlaneDivider - (CrewDivider + 2));
                sPlaneID = sCrewPlaneID.Substring(PlaneDivider + 1);
                functions.setPlaneBooked(sPlaneID);


                //Add flight to excel doc
                //create the excel variables
                Excel.Workbook xlWorkbook = functions.database_connect();
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[2];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int numRows = functions.getRows(2);



                xlRange.Cells[flightRow, 3].value = sPilots;
                xlRange.Cells[flightRow, 4].value = sAttendant;
                xlRange.Cells[flightRow, 14].value = sPlaneID;
                xlRange.Cells[flightRow, 21].value = Identification;

                xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU OPEN
                xlWorkbook.Close(); //THIS ONE TOO

                MainMenuMarketingManager mainMenu = new MainMenuMarketingManager(Identification); //create a new main menu and go to it
                this.NavigationService.Navigate(mainMenu);

            }
            else if (PlaneStatus == "ZERO")
            {
                Warning.Text = "Airport does not have selected model";
            }
            else if (PlaneStatus == "EMPTY")
            {
                Warning.Text = "Airport is empty";
            }

        }
        private void Sign_Out(object sender, RoutedEventArgs e)
        { //to sign out
            SignIn signIn = new SignIn(); //create a new main menu and go to it
            this.NavigationService.Navigate(signIn);
        }

        private void Main_Menu(object sender, RoutedEventArgs e)
        { //to return to the main menu

            MainMenuMarketingManager mainMenu = new MainMenuMarketingManager(Identification); //create a new main menu and go to it
            this.NavigationService.Navigate(mainMenu);
            
        }
    }
}
