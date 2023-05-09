using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using NanoXLSX;

namespace Szolgalati2
{
    class Program
    {

        public static int EmployersCount = 25;
        public static List<Employer> Employers = new List<Employer>();
        public static List<string> ServicePhones = new List<string>();



        static void Main(string[] args)
        {

            SetEmployers();

            GetEmployers();

            Console.ReadKey();

        }



        public static void GetEmployers()
        {
            foreach(Employer emp in Employers)
            {
                Console.WriteLine(emp.GetFullName());
                Console.WriteLine(emp.Phone);
                for(int i = 0; i < emp.WorkDays.Count; i++)
                {
                    Console.Write(emp.WorkDays.ElementAt(i).Key + ": " + emp.WorkDays.ElementAt(i).Value);
                }
            }
        }



        public static string GetEmployerPhone(string firstName, string lastName)
        {
            StreamReader reader = new StreamReader("telefonok.txt");

            try
            {
                string line;
                string name = (firstName + " " + lastName).Trim().ToLower();
                while ((line = reader.ReadLine()) != null)
                {
                    string[] row = line.Split(';');

                    if(row[0].Trim().ToLower() == name)
                    {
                        return row[1];
                    }
                }

            } catch(IOException e)
            {
                Console.WriteLine(e);
            }

            return "-";
        }




        public static void SetEmployers()
        {
            StreamReader reader = new StreamReader("betegkiserok.csv");

            try
            {
                string line;
                int i = 0;

                while ((line = reader.ReadLine()) != null)
                {
                    string[] row = line.Split(',');

                    if (i > 5 && EmployersCount + 6 > i)
                    {
                        string[] name = row[0].Split(' ');
                        string firstName = name[0];
                        string lastName = name[1];
                        string phone = GetEmployerPhone(firstName, lastName);
                        Dictionary<string, string> workDays = new Dictionary<string, string>();

                        // System.Console.WriteLine(firstName + " " + lastName);
                        for (int j = 3; j < row.Length; j++)
                        {
                            if(row[j] != "")
                            {
                                workDays.Add((j - 2).ToString(), row[j]);
                                System.Console.WriteLine((j - 2) + ": " + row[j]);
                            }
                        }
                        // System.Console.WriteLine("");

                        Employer employer = new Employer(firstName, lastName, phone, workDays);
                        Employers.Add(employer);

                    }
                    i++;
                }

            } catch(IOException e)
            {
                System.Console.WriteLine(e);
            }

            reader.Close();
        }
    }
}
