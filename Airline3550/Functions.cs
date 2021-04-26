using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
namespace Air3550
{
    class Functions
    {

        public int NUM_AIRPORTS = 14;

        public string CEprofile(string firstName, string middleName, string lastName, string address, string city, string ZIP, string phone, string email, string credit, string csv, string password, string birth, string expiration)
        { //check that the profile is formatted correctly
            if (birth == "")
            { //if the birth date wasn't entered
                return "No birth date entered";
            }

            if (expiration == "")
            { //if the expiration wasn't entered
                return "No expiration date entered";
            }

            if (firstName == "")
            { //if the first name wasn't entered
                return "No first name entered";
            }
            for (int i = 0; i < firstName.Length; i++)
            { //check if the first name is formatted correctly
                if ((i == 0) && (firstName[i] > 90 || firstName[i] < 65))
                { //if the first letter isn't capitalized
                    return "First name formatting wrong. Start with a capital letter, the rest are lower case";
                }
                if ((i > 0) && (firstName[i] > 122 || firstName[i] < 97))
                { //if any other letters are capitalized
                    return "First name formatting wrong. Start with a capital letter, the rest are lower case";
                }
            }

            if (middleName == "")
            { //if the middle name wasn't entered
                return "No middle name entered";
            }
            for (int i = 0; i < middleName.Length; i++)
            { //check if the middle name is formatted correctly
                if ((i == 0) && (middleName[i] > 90 || middleName[i] < 65))
                { //if the middle letter isn't capitalized
                    return "Middle name formatting wrong. Start with a capital letter, the rest are lower case";
                }
                if ((i > 0) && (middleName[i] > 122 || middleName[i] < 97))
                { //if any other letters are capitalized
                    return "Middle name formatting wrong. Start with a capital letter, the rest are lower case";
                }
            }

            if (lastName == "")
            { //if the last name wasn't entered
                return "No last name entered";
            }
            for (int i = 0; i < lastName.Length; i++)
            { //check if the last name is formatted correctly
                if ((i == 0) && (lastName[i] > 90 || lastName[i] < 65))
                { //if the last letter isn't capitalized
                    return "Last name formatting wrong. Start with a capital letter, the rest are lower case";
                }
                if ((i > 0) && (lastName[i] > 122 || lastName[i] < 97))
                { //if any other letters are capitalized
                    return "Last name formatting wrong. Start with a capital letter, the rest are lower case";
                }
            }

            bool containsAt = false; //check if the email has @
            if (email == "")
            { //if the email wasn't entered
                return "No email entered";
            }
            for (int i = 0; i < email.Length - 4; i++)
            { //check if the email is formatted correctly
                if (email[i] == 64)
                { //if it contains @, set bool to true
                    containsAt = true;
                }
            }

            if (!((containsAt == true) && ((email.Substring(Math.Max(0, email.Length - 4)) == ".com") ||
                (email.Substring(Math.Max(0, email.Length - 4)) == ".org") ||
                (email.Substring(Math.Max(0, email.Length - 4)) == ".edu"))))
            { //if the email formatting is wrong
                return "Email formatting wrong. Must contain @ and end in .com, .org, or .edu";
            }

            if (address == "")
            { //if the address wasn't entered
                return "No address entered";
            }

            if (city == "")
            { //if the city wasn't entered
                return "No city entered";
            }

            for (int i = 0; i < city.Length; i++)
            { //check that the city contains the correct characters
                if (!((city[i] > 64 && city[i] < 91) || (city[i] > 96 && city[i] < 123) || city[i] == ' '))
                {
                    return "City Formatting Wrong. Words start with a capital letter, the other letters are lower case";
                }
            }

            if (ZIP == "")
            { //if the zip code wasn't entered
                return "No zip code entered";
            }
            if (!(ZIP.All(char.IsDigit)) || ZIP.Length != 5)
            { //check that the zip code is correct
                return "ZIP Code Wrong. Should be 5 numbers";
            }

            //var pattern1 = @"\((?<AreaCode>\d{3})\)\s*(?<Number>\d{3}(?:-|\s*)\d{4})"; //create patterns for the phone number
            // var regexp1 = new System.Text.RegularExpressions.Regex(pattern1);
            //var pattern2 = @"\(?<AreaCode>\d{3}(?:-|\s*)?<Number>\d{3}(?:-|\s*)\d{4})";
            //var regexp2 = new System.Text.RegularExpressions.Regex(pattern2);

            if (phone == "")
            { //if the phone number wasn't entered
                return "No phone number entered";
            }

            //if (!(regexp1.IsMatch(phone) || regexp2.IsMatch(phone)))
            if (phone.Length == 12 || phone.Length == 14)
            { //check the formatting of the phone number
                if (!((Char.IsDigit(phone[0]) && Char.IsDigit(phone[1]) && Char.IsDigit(phone[2]) && phone[3] == 45 &&
                    Char.IsDigit(phone[4]) && Char.IsDigit(phone[5]) && Char.IsDigit(phone[6]) && phone[7] == 45 &&
                    Char.IsDigit(phone[8]) && Char.IsDigit(phone[9]) && Char.IsDigit(phone[10]) && Char.IsDigit(phone[11])) ||
                    (phone[0] == 40 && Char.IsDigit(phone[1]) && Char.IsDigit(phone[2]) && Char.IsDigit(phone[3]) &&
                    phone[4] == 41 && phone[5] == 32 && Char.IsDigit(phone[6]) && Char.IsDigit(phone[7]) && Char.IsDigit(phone[8]) &&
                    phone[9] == 45 && Char.IsDigit(phone[10]) && Char.IsDigit(phone[11]) && Char.IsDigit(phone[12]) && Char.IsDigit(phone[13]))))
                { //if wrong format for the phone number
                    return "Phone Formatting Wrong. Must be of the form ###-###-#### or (###) ###-####";
                }
            }
            else
            {
                return "Phone Formatting Wrong. Must be of the form ###-###-#### or (###) ###-####";
            }

            //var pattern4 = @"\(<Number>\d{4}(?:-|\s*)\d{4}(?:-|\s*)\d{4}(?:-|\s*)\d{4})";
            //var regexp4 = new System.Text.RegularExpressions.Regex(pattern4); //create a pattern for the birth date

            if (credit == "")
            { //if the credit card number wasn't entered
                return "No credit card number entered";
            }

            //if (!(regexp4.IsMatch(credit)))
            if (credit.Length != 19 && !(Char.IsDigit(credit[0]) && Char.IsDigit(credit[1]) && Char.IsDigit(credit[2]) && Char.IsDigit(credit[3])
                && credit[4] == 45 && Char.IsDigit(credit[5]) && Char.IsDigit(credit[6]) && Char.IsDigit(credit[7]) && Char.IsDigit(credit[8])
                && credit[9] == 45 && Char.IsDigit(credit[10]) && Char.IsDigit(credit[11]) && Char.IsDigit(credit[12]) && Char.IsDigit(credit[13])
                && credit[14] == 45 && Char.IsDigit(credit[15]) && Char.IsDigit(credit[16]) && Char.IsDigit(credit[17]) && Char.IsDigit(credit[18])))
            { //return that the credit card format was wrong
                return "Credit Card Formatting Wrong. Must be of the form ####-####-####-####";
            }

            //var pattern6 = @"\(<Number>\d{3})";
            //var regexp6 = new System.Text.RegularExpressions.Regex(pattern6); //create a pattern for the birth date
            if (csv == "")
            { //if the csv wasn't entered
                return "No csv entered";
            }

            //if (!(regexp6.IsMatch(csv)))
            if (!(csv.All(char.IsDigit)) || csv.Length != 3)
            { //return that the birth date format was wrong
                return "Credit CSV Formatting Wrong. Must be three digits";
            }

            //if (password == "")
            //{ //if the password wasn't entered
            //    return "No password entered";
            //}

            return "Correct";
        }

        public bool isNum(string input)
        { //function to check if an input is a number

            for (int i = 0; i < input.Length; i++)
            {
                if ((input[i] == '.') || Char.IsDigit(input[i]))
                {

                }
                else
                {
                    return false;
                }
            }
            return true;
        }
        public bool isTime(string input)
        { //function to check if an input is a time

            if ((input.Length > 5 && input.Length < 9) && ((Char.IsDigit(input[0]) && input[1] == ':' &&
                Char.IsDigit(input[2]) && Char.IsDigit(input[3])) || (Char.IsDigit(input[0]) && Char.IsDigit(input[1])
                && input[2] == ':' && Char.IsDigit(input[3]) && Char.IsDigit(input[4]))))
            { //check that the input is a time
                if (input.Substring(input.Length - 2) == "PM" || input.Substring(input.Length - 2) == "AM")
                { //check that AM or PM is at the back
                    return true;
                }
                else
                {
                    return false;
                }
            }
            return false;
        }

        public Excel.Workbook database_connect()
        { //easy way to connect to a database so that, when a user needs to change the file path, they only do so in one location
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\Nathan Burns\Desktop\Classes\Software Engineering\Air3550_Database\Air3550Database.xlsx");
            return xlWorkbook;
        }
        public int getIDRow(string ID, int sheet)
        { //get the ID column
            //connect to the excel database
            //Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = database_connect();
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[sheet];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = getRows(sheet);
            int colCount = xlRange.Columns.Count;
            int IDrow = 0; //initialize the ID column

            for (int i = 1; i <= rowCount; i++)
            { //get the column of the ID
                if (xlRange.Cells[i, 1].Value2.ToString() == ID)
                { //if we found the ID, set the column
                    IDrow = i;
                }
            }
            xlWorkbook.Close();
            return IDrow; //return the ID column
        }

        public string getUserType(int IDcolumn)
        { //get the user type
            //connect to the database
            //Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = database_connect();
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = getRows(1);
            int colCount = xlRange.Columns.Count;
            string userType = xlRange.Cells[IDcolumn, 2].Value2.ToString(); //get the user type from the database
            xlWorkbook.Close();
            return userType; //return the user type
        }
        public string getName(int IDcolumn)
        { //get the user's name
            //connect to the database
            //Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = database_connect();
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = getRows(1);
            int colCount = xlRange.Columns.Count;
            string name = xlRange.Cells[IDcolumn, 3].Value2.ToString(); //get the user's name from the database
            xlWorkbook.Close();
            return name; //return the name
        }

        public string getAirportCode(string airport)
        {  //function to get the airport code based on the inputted airport
            string code = " ";
            //go through the airports and return the code
            if (airport == "Cleveland, Ohio")
            {
                code = "44135";
            }
            else if (airport == "Nashville, Tennessee")
            {
                code = "37214";
            }
            else if (airport == "Miami, Florida")
            {
                code = "33122";
            }
            else if (airport == "Houston, Texas")
            {
                code = "77032";
            }
            else if (airport == "Queens, New York")
            {
                code = "11340";
            }
            else if (airport == "Billings, Montana")
            {
                code = "59105";
            }
            else if (airport == "Los Angeles, California")
            {
                code = "90045";
            }
            else if (airport == "Ketchikan, Alaska")
            {
                code = "99901";
            }
            else if (airport == "Hilo, Hawaii")
            {
                code = "96720";
            }
            else if (airport == "Salt Lake City, Utah")
            {
                code = "84122";
            }
            else if (airport == "San Diego, California")
            {
                code = "92101";
            }
            else if (airport == "Abuquerque, New Mexico")
            {
                code = "87106";
            }
            else if (airport == "Birmingham, Alabama")
            {
                code = "35212";
            }
            else if (airport == "Kansas City, Missouri")
            {
                code = "64153";
            }
            return code;
        }

        public string getAirport(string airport)
        {  //function to get the airport code based on the inputted airport
            string code = " ";
            //go through the airport codes and return the airport
            if (airport == "44135")
            {
                code = "Cleveland, Ohio";
            }
            else if (airport == "37214")
            {
                code = "Nashville, Tennessee";
            }
            else if (airport == "33122")
            {
                code = "Miami, Florida";
            }
            else if (airport == "77032")
            {
                code = "Houston, Texas";
            }
            else if (airport == "11340")
            {
                code = "Queens, New York";
            }
            else if (airport == "59105")
            {
                code = "Billings, Montana";
            }
            else if (airport == "90045")
            {
                code = "Los Angeles, California";
            }
            else if (airport == "99901")
            {
                code = "Ketchikan, Alaska";
            }
            else if (airport == "96720")
            {
                code = "Hilo, Hawaii";
            }
            else if (airport == "84122")
            {
                code = "Salt Lake City, Utah";
            }
            else if (airport == "92101")
            {
                code = "San Diego, California";
            }
            else if (airport == "87106")
            {
                code = "Abuquerque, New Mexico";
            }
            else if (airport == "35212")
            {
                code = "Birmingham, Alabama";
            }
            else if (airport == "64143")
            {
                code = "Kansas City, Missouri";
            }
            return code;
        }

        public bool fullFlight(int attendance, string plane)
        { //check if a flight is completely booked
            bool full = true; //value to say if we're full or not
            //define the excel variables
            //Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = database_connect();
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[3];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = getRows(3);
            int colCount = xlRange.Columns.Count;
            for (int i = 1; i <= rowCount; i++)
            { //get the column of the ID
                if (xlRange.Cells[i, 1].Value2.ToString() == plane)
                { //if we found the ID, set the column
                    if (attendance < Int32.Parse(xlRange.Cells[i, 3].Value2.ToString()))
                    { //if the flight is not completely booked, set full to false
                        full = false;
                    }
                }
            }
            xlWorkbook.Close();
            return full;
        }
        public int getRows(int sheet)
        { //used to get the number of rows in a workbook
            int rowCount = 1; //initialize the row count
            Excel.Workbook xlWorkbook = database_connect();
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[sheet];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            //while the rows aren't null, loop through and increment the counter
            while (true)
            {
                if (((xlRange.Cells[rowCount + 1, 1].Value2) == null) || (xlRange.Cells[rowCount + 1, 1].Value2.ToString()==""))
                { //if the cell is null or empty, break
                    break;
                }
                rowCount++;
            }
            xlWorkbook.Close();
            return rowCount;
        }
        public bool isEmpty(int sheet, int row, int column)
        { //check if a cell is empty
            bool empty = false; //boolean value to return
            Excel.Workbook xlWorkbook = database_connect(); //define the excel values
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[sheet];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            if (((xlRange.Cells[row, column].Value2) == null) || (xlRange.Cells[row, column].Value2.ToString() == ""))
            { //if the cell is null or empty, set to true
                empty = true;
            }
            xlWorkbook.Close();
            return empty;
        }

        public bool isFlight(string flightID)
        { //check if the flight exists
            bool exists = false; //boolean value to return
            //create the excel variables
            Excel.Workbook xlWorkbook = database_connect();
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[2];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = getRows(2);
            for (int i = 1; i <= rowCount; i++)
            {
                if (xlRange.Cells[i, 1].Value2.ToString() == flightID)
                { //if we found the ID, set exists to true
                    exists = true;
                }
            }
            xlWorkbook.Close(); //close the workbook
            return exists;
        }

        public bool flightBooked (string userID, string flightID)
        { //check if a user has booked a flight
            bool booked = false; //value to return if the user is one the flight
            int userRow = getIDRow(userID, 1); //get the userID's row in the database
            int flightRow = getIDRow(flightID, 2);
            //define the excel variables
            Excel.Workbook xlWorkbook = database_connect();
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = getRows(1);
            if (!isEmpty(1, userRow, 19))
            {
                string flights = xlRange.Cells[userRow, 19].Value2.ToString(); //get the string for the flights the user is booked for
                if (flights != "" && flights.Contains(flightID))
                { //if the flights contains the flight, the user is already booked for it
                    booked = true; //set booked to true
                }
            }
            
            xlWorkbook.Close(); //close the workbook
            return booked;
        }
        public void createUserRecord(string userID, string flightID, string method, string name, string price, string CCN)
        { //function to create a new user record

            string fileName = @"C:\Users\Nathan Burns\Desktop\Classes\Software Engineering\Air3550_Database\UserRecords\" + userID + @"\" + flightID + @".txt"; //create the path for the file
                                                                                                                                                                // Check if file already exists. If yes, delete it.     
            FileStream fs = File.Create(fileName, 5000);
            //Pass the filepath and filename to the StreamWriter Constructor
            StreamWriter sw = new StreamWriter(@"C:\Users\Nathan Burns\Desktop\Classes\Software Engineering\Air3550_Database\UserRecords\" + userID + @"\" + flightID + @".txt");
            //Write the name
            sw.WriteLine("Name :" + name);
            //Write the price paid
            sw.WriteLine("Price: " + price);
            //Write the method of payment
            sw.WriteLine("Method: " + method);
            if (method == "Credit Card")
            { //if a credit card was used to pay, write the credit card number
                //Write the credit card number
                sw.WriteLine("Credit Card Number: " + CCN);
            }
            
            //Close the file
            sw.Close();

        }
        //Uses string of location name NOT ZIP
        public string getDistance(string Origin, string Destination)
        {
            //Add flight to excel doc
            //create the excel variables
            Excel.Workbook xlWorkbook = database_connect();
            Excel._Worksheet xlWorksheetAir = xlWorkbook.Sheets[4];
            Excel.Range xlRangeAir = xlWorksheetAir.UsedRange;

            //Get distance from Airport table
            string sDistance = null;
            int rowCount = getRows(2);
            int DistanceRow = 2;
            string Location = Origin;
            //string Test;

            for (int i = 2; i <= NUM_AIRPORTS + 1; i++)
            {//Once origin row is found
                //Test = xlRangeAir.Cells[i, 4].Value2.ToString();
                if (xlRangeAir.Cells[i, 4].Value2.ToString() == Location)
                {//Look for ending location
                    DistanceRow = i;
                    break;
                }
            }

            Location = Destination;
            for (int j = 10; j < j + NUM_AIRPORTS + 1; j++)
            {//Once ending location is found
                //Test = xlRangeAir.Cells[1, j].Value2.ToString();
                if (xlRangeAir.Cells[1, j].Value2.ToString() == Location)
                {//Get distance
                    sDistance = xlRangeAir.Cells[DistanceRow, j].Value2.ToString();
                    break;
                }
            }

            //            xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU OPEN
            xlWorkbook.Close(); //THIS ONE TOO
            return sDistance;
        }

        public string getZip(string sLocation)
        {
            //Open Airports tab on excel
            Excel.Workbook xlWorkbook = database_connect();
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[4];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            string Temp;
            //Match Location
            for (int i = 2; i <= NUM_AIRPORTS + 1; i++)
            {
                if (sLocation == xlRange.Cells[i, 4].Value2.ToString())
                {//Get Zip Code
                    Temp = xlRange.Cells[i, 5].Value2.ToString();
                    //                    xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU OPEN
                    xlWorkbook.Close(); //THIS ONE TOO
                    return Temp;
                }
            }
            //Should be impossible
            //            xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU OPEN
            xlWorkbook.Close(); //THIS ONE TOO
            return null;
        }

        public string getLocation(string sZip)
        {
            //Open Airports tab on excel
            Excel.Workbook xlWorkbook = database_connect();
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[4];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            string Temp;
            //Match Location
            for (int i = 2; i <= NUM_AIRPORTS + 1; i++)
            {
                if (sZip == xlRange.Cells[i, 5].Value2.ToString())
                {//Get Zip Code
                    Temp = xlRange.Cells[i, 4].Value2.ToString();
                    //                    xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU OPEN
                    xlWorkbook.Close(); //THIS ONE TOO
                    return Temp;
                }
            }
            //Impossible if Database is correctly formatted
            xlWorkbook.Application.ActiveWorkbook.Save(); //MAKE SURE TO USE THESE TO SAVE AND CLOSE EVERY WORKBOOK YOU OPEN
            xlWorkbook.Close(); //THIS ONE TOO
            return null;
        }

        //Hi professor and or grader
        //I may have cut a few corners with this one
        public string getArrival(string sDistance, string sStartDate, string sStartTime)
        {
            float iDistance = float.Parse(sDistance);
            int Hours, Minutes;
            int iStartHour, iStartMin, iStartDay, iStartMonth, iStartYear;
            string Temp, sAMorPM = sStartTime.Substring(sStartTime.Length - 2, 2);

            //Seperate the start time componets
            if (Char.IsDigit(sStartTime[1]))
            {//Hour has 2 digits
                Temp = sStartTime.Substring(0, 2);
                iStartHour = int.Parse(Temp);
                Temp = sStartTime.Substring(3, 2);
                iStartMin = int.Parse(Temp);
            }
            else
            {//Hour has 1 digit
                Temp = sStartTime.Substring(0, 1);
                iStartHour = int.Parse(Temp);
                Temp = sStartTime.Substring(3, 2);
                iStartMin = int.Parse(Temp);
            }

            //Seperate the Date components
            if (Char.IsDigit(sStartDate[1]))
            {//Month has 2 digits
                Temp = sStartDate.Substring(0, 2);
                iStartMonth = int.Parse(Temp);

                if (Char.IsDigit(sStartDate[4]))
                {//Day has 2 digits
                    Temp = sStartDate.Substring(3, 2);
                    iStartDay = int.Parse(Temp);
                    Temp = sStartDate.Substring(6, 4);
                    iStartYear = int.Parse(Temp);
                    //Full Date Set
                }
                else
                {//Day has 1 digit
                    Temp = sStartDate.Substring(3, 1);
                    iStartDay = int.Parse(Temp);
                    Temp = sStartDate.Substring(5, 4);
                    iStartYear = int.Parse(Temp);
                    //Full Date Set
                }

            }
            else
            {//Month has 1 digit
                Temp = sStartDate.Substring(0, 1);
                iStartMonth = int.Parse(Temp);

                if (Char.IsDigit(sStartDate[3]))
                {//Day has 2 digits
                    Temp = sStartDate.Substring(2, 2);
                    iStartDay = int.Parse(Temp);
                    Temp = sStartDate.Substring(5, 4);
                    iStartYear = int.Parse(Temp);
                    //Full Date Set
                }
                else
                {//Day has 1 digit
                    Temp = sStartDate.Substring(2, 1);
                    iStartDay = int.Parse(Temp);
                    Temp = sStartDate.Substring(4, 4);
                    iStartYear = int.Parse(Temp);
                    //Full Date Set
                }
            }

            //Total Min
            int TripTime = ((int)iDistance / 500) + 30;

            //Convert to hours and minutes
            if (TripTime > 60)
            {
                Hours = TripTime / 60;
                Minutes = TripTime - (60 * Hours);
            }
            else
            {
                Hours = 0;
                Minutes = TripTime;
            }


            //Add trip time to departure time
            iStartMin += Minutes;
            if (iStartMin >= 60)
            {
                iStartHour++;
                iStartMin = iStartMin - 60;
            }

            iStartHour += Hours;
            if (iStartHour > 12 && sAMorPM == "PM")
            {//Overnight... Switch Day
                iStartDay++;
                sAMorPM = "AM";
            }
            else if (iStartHour > 12)
            {//Flip to PM
                iStartHour -= 12;
                sAMorPM = "PM";
            }

            //ACorrect date if necessary

            //The february special
            if (iStartMonth == 2 && (iStartYear % 4 == 0) && (iStartDay > 29))
            {//Leap Year
                iStartMonth++;
                iStartDay = 1;
            }
            else if (iStartMonth == 2 && iStartDay > 28)
            {//Non Leap year
                iStartMonth++;
                iStartDay = 1;
            }
            else if (iStartMonth == 12 && iStartDay > 31)
            {//New Year
                iStartMonth = 1;
                iStartDay = 1;
                iStartYear++;
            }
            else if (iStartMonth == (1 | 3 | 5 | 7 | 8 | 10) && iStartDay > 31)
            {//Next Month (31 days)
                iStartMonth++;
                iStartDay = 1;
            }
            else if (iStartMonth == (2 | 4 | 6 | 9 | 11) && iStartDay > 30)
            {//Next Month (30 days)
                iStartMonth++;
                iStartDay = 1;
            }

            return iStartHour.ToString() + ":" + iStartMin + " " + sAMorPM + " " + iStartMonth.ToString() + "/" + iStartDay.ToString() + "/" + iStartYear.ToString();
        }
    }


}

