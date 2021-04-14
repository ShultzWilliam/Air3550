using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace Air3550
{
    class Functions
    {
        public string CEprofile(string name, string address, string city, string state, string ZIP, string phone, string email, string birth, string credit, string exp, string csv)
        { //check that the profile is formatted correctly
            int names = 0;

            for (int i = 0; i < name.Length; i++)
            { //check that the name is two or three words and that no chars are digits
                if(name[i] == ' ')
                {
                    names++;
                }
                if (((i == 0) || (i > 0 && name[i-1] == ' ')) && (name[i] > 90 || name[i] < 65))
                { //if the beginning of the name isn't capitalized
                    return "name formatting wrong";
                }
                if ((i > 0 && name[i - 1] != ' ') && (name[i] > 122 || name[i] < 97))
                { //if the other letters in the name aren't lowercased (or letters)
                    return "name formatting wrong";
                }
            }

            if (names < 2 || names > 3)
            { //keep checking the name
                return "Name Wrong";
            }

            //check the address and email

            for(int i = 0; i < city.Length; i++)
            { //check that the city contains the correct characters
                if (!((city[i] > 64 && city[i] < 91) || (city[i] > 96 && city[i] < 123) || city[i] == ' '))
                {
                    return "City Wrong";
                }
            }

            for (int i = 0; i < state.Length; i++)
            { //check that the state contains the correct characters
                if (!((state[i] > 64 && state[i] < 91) || (state[i] > 96 && state[i] < 123) || state[i] == ' '))
                {
                    return "State Wrong";
                }
            }

            if (!(ZIP.All(char.IsDigit)) || ZIP.Length != 5)
            { //check that the zip code is correct
                return "ZIP Wrong";
            }

            var pattern1 = @"\((?<AreaCode>\d{3})\)\s*(?<Number>\d{3}(?:-|\s*)\d{4})"; //create patterns for the phone number
            var regexp1 = new System.Text.RegularExpressions.Regex(pattern1);
            var pattern2 = @"\(?<AreaCode>\d{3}(?:-|\s*)?<Number>\d{3}(?:-|\s*)\d{4})";
            var regexp2 = new System.Text.RegularExpressions.Regex(pattern2);

            if(!(regexp1.IsMatch(phone) || regexp2.IsMatch(phone)))
            { //if wrong format for the phone number
                return "Phone Wrong";
            }

            var pattern3 = @"\(<Number>\d{3}(?:/|\s*)\d{3}(?:/|\s*)\d{4})";
            var regexp3 = new System.Text.RegularExpressions.Regex(pattern3); //create a pattern for the birth date

            if (!(regexp3.IsMatch(birth)))
            { //return that the birth date format was wrong
                return "Birth Wrong";
            }

            var pattern4 = @"\(<Number>\d{4}(?:-|\s*)\d{4}(?:-|\s*)\d{4}(?:-|\s*)\d{4})";
            var regexp4 = new System.Text.RegularExpressions.Regex(pattern4); //create a pattern for the birth date

            if (!(regexp4.IsMatch(credit)))
            { //return that the birth date format was wrong
                return "Credit Wrong";
            }

            var pattern5 = @"\(<Number>\d{2}(?:-|\s*)\d{2})";
            var regexp5 = new System.Text.RegularExpressions.Regex(pattern5); //create a pattern for the birth date

            if (!(regexp5.IsMatch(exp)))
            { //return that the birth date format was wrong
                return "Credit EXP Wrong";
            }

            var pattern6 = @"\(<Number>\d{3})";
            var regexp6 = new System.Text.RegularExpressions.Regex(pattern6); //create a pattern for the birth date

            if (!(regexp6.IsMatch(csv)))
            { //return that the birth date format was wrong
                return "Credit CSV Wrong";
            }
            return "Correct";
        }

        public bool isNum(string input)
        { //function to check if an input is a number

            if (input.All(char.IsDigit))
            { //check that the input only contains numbers
                return true;
            }
            return false;
        }
        public bool isTime(string input)
        { //function to check if an input is a time

            if ((input.Length > 3 && input.Length < 6) && ((Char.IsDigit(input[0]) && input[1] == ':' &&
                Char.IsDigit(input[2]) && Char.IsDigit(input[3])) || (Char.IsDigit(input[0]) && Char.IsDigit(input[1])
                && input[2] == ':' && Char.IsDigit(input[3]) && Char.IsDigit(input[4]))))
            { //check that the input is a time
                return true;
            }
            return false;
        }
    }

    
}
