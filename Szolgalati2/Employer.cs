using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Szolgalati2
{
    class Employer
    {
        public string FirstName;
        public string LastName;
        public string Phone { get; set; }
        public Dictionary<string, string> WorkDays;

        public Employer(string FirstName, string LastName, string Phone, Dictionary<string, string> WorkDays)
        {
            this.FirstName = FirstName;
            this.LastName = LastName;
            this.Phone = Phone;
            this.WorkDays = WorkDays;
        }

        public string GetFullName()
        {
            return FirstName + " " + LastName;
        }
    }
}
