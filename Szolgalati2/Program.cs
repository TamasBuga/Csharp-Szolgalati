using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using NanoXLSX;
using NanoXLSX.Styles;

namespace Szolgalati2
{
    class Program
    {




        public static int EmployersCount = 25;
        public static int Month = 5;
        public static int Day = 8;
        public static List<Employer> Employers = GetEmployers();
        public static List<string> ServicePhones = GetServicePhones();
        public static List<string> EmployersOfTheDay = GetEmployersOfTheDay("5", "19-7");





        static void Main(string[] args)
        {

            // DisplayEmployers();

            // DisplayEmployersOfTheDay();

            ReadXLSX();


            Console.ReadKey();

        }




        // =====================================================================

        // Read and set values on xlsx file

        // =====================================================================

        public static void ReadXLSX()
        {
            Workbook wb = Workbook.Load("szolgalatilap.xlsx");
            // Console.WriteLine(wb.CurrentWorksheet.SheetName);

            Style s = new Style();
            s.CurrentFont.Italic = true;
            // s.CurrentFont.Bold = true;
            s.CurrentFont.Size = 14;
            s.CurrentBorder.BottomStyle = Border.StyleValue.thin;
            s.CurrentBorder.LeftStyle = Border.StyleValue.thin;
            s.CurrentBorder.RightStyle = Border.StyleValue.thin;
            s.CurrentBorder.TopStyle = Border.StyleValue.thin;

            string date = "Dátum: " + FormatDate(Month) + "." + FormatDate(Day) + ".";
            wb.CurrentWorksheet.AddCell(date, 1, 0, s);
            wb.CurrentWorksheet.AddCell(date, 4, 0, s);
            wb.CurrentWorksheet.AddCell(date, 1, 17, s);
            wb.CurrentWorksheet.AddCell(date, 4, 17, s);

            int ambIndex = 0;
            int fekIndex = 0;

            for (int i = 0; i < EmployersOfTheDay.Count; i++)
            {
                
                string[] emp = EmployersOfTheDay.ElementAt(i).Split(';');
                string empName = FormatName(emp[0]);
                string empPhone = emp[1];
                string empShift = emp[2];

                // Szolgálatvezető
                if (empShift.Contains('*'))
                {
                    // Név
                    wb.CurrentWorksheet.AddCell(empName, 0, 3, s);
                    wb.CurrentWorksheet.AddCell(empName, 3, 3, s);
                    wb.CurrentWorksheet.AddCell(empName, 0, 20, s);
                    wb.CurrentWorksheet.AddCell(empName, 3, 20, s);

                    // Telefon
                    wb.CurrentWorksheet.AddCell(empPhone, 1, 3, s);
                    wb.CurrentWorksheet.AddCell(empPhone, 4, 3, s);
                    wb.CurrentWorksheet.AddCell(empPhone, 1, 20, s);
                    wb.CurrentWorksheet.AddCell(empPhone, 4, 20, s);
                }

                // Fektetős
                if (empShift.Contains('F'))
                {
                    // Név
                    wb.CurrentWorksheet.AddCell(empName, 0, 14 + fekIndex, s);
                    wb.CurrentWorksheet.AddCell(empName, 3, 14 + fekIndex, s);
                    wb.CurrentWorksheet.AddCell(empName, 0, 31 + fekIndex, s);
                    wb.CurrentWorksheet.AddCell(empName, 3, 31 + fekIndex, s);

                    // Telefon
                    wb.CurrentWorksheet.AddCell(empPhone, 1, 14 + fekIndex, s);
                    wb.CurrentWorksheet.AddCell(empPhone, 4, 14 + fekIndex, s);
                    wb.CurrentWorksheet.AddCell(empPhone, 1, 31 + fekIndex, s);
                    wb.CurrentWorksheet.AddCell(empPhone, 4, 31 + fekIndex, s);

                    fekIndex++;
                }

                // Ambulancia
                if(!empShift.Contains('F'))
                {
                    // Név
                    wb.CurrentWorksheet.AddCell(empName, 0, 7 + ambIndex, s);
                    wb.CurrentWorksheet.AddCell(empName, 3, 7 + ambIndex, s);
                    wb.CurrentWorksheet.AddCell(empName, 0, 24 + ambIndex, s);
                    wb.CurrentWorksheet.AddCell(empName, 3, 24 + ambIndex, s);

                    // Telefon
                    wb.CurrentWorksheet.AddCell(empPhone, 1, 7 + ambIndex, s);
                    wb.CurrentWorksheet.AddCell(empPhone, 4, 7 + ambIndex, s);
                    wb.CurrentWorksheet.AddCell(empPhone, 1, 24 + ambIndex, s);
                    wb.CurrentWorksheet.AddCell(empPhone, 4, 24 + ambIndex, s);

                    ambIndex++;
                }

                // Console.WriteLine(empName + " " + empPhone + " " + empShift
            }


            // datum: (1, 0), (4, 0), (1, 17), (4, 17)


            // nevek:
            // szolgalatvezeto: (0, 3) (3, 3) (0, 20) (3, 20)
            // Ambulancia: (0, 7-12) (3, 7-12) (0, 24-29) (3, 24-29)
            // fekteto: (0, 14-15) (3, 14-15) (0, 31-32) (3, 31-32)

            // telefonok:
            // szolgalatvezeto: (1, 3) (4, 3) (1, 20) (4, 20)
            // Ambulancia: (1, 7-12) (4, 7-12) (1, 24-29) (4, 24-29)
            // fekteto: (1, 14-15) (4, 14-15) (1, 31-32) (4, 31-32)

            wb.SaveAs("demo3.xlsx");
        }


        public static string FormatDate(int date)
        {
            if (date < 10)
            {
                return "0" + date;
            }
            else
            {
                return Convert.ToString(date);
            }
        }



        public static string FormatName(string name)
        {
            string[] empName = name.Split(' ');
            string result = "";
            for(int i = 0; i < empName.Length; i++)
            {
                result += char.ToUpper(empName[i][0]) + empName[i].Substring(1).ToLower();
                if(i < empName.Length - 1)
                {
                    result += " ";
                }
            }
            return result;
        }



        // =====================================================================

        // Read and get Data from files

        // =====================================================================

        public static void DisplayEmployersOfTheDay()
        {
            for (int i = 0; i < EmployersOfTheDay.Count; i++)
            {
                string[] splitter = EmployersOfTheDay.ElementAt(i).Split(';');
                Console.WriteLine(splitter[0] + " : " + splitter[1]);

            }
        }




        public static List<string> GetEmployersOfTheDay(string day, string shift)
        {
            List<string> emps = new List<string>();

            for (int i = 0; i < Employers.Count; i++)
            {
                for (int j = 0; j < Employers.ElementAt(i).WorkDays.Count; j++)
                {
                    string empShift = Employers.ElementAt(i).WorkDays.ElementAt(j).Value;
                    if (Employers.ElementAt(i).WorkDays.ElementAt(j).Key == day && empShift.Contains(shift))
                    {
                        if (Employers.ElementAt(i).Phone == "-" && ServicePhones.Count > 0)
                        {
                            emps.Add(Employers.ElementAt(i).GetFullName() + ";" + ServicePhones.ElementAt(0) + ";" + empShift);
                            ServicePhones.RemoveAt(0);
                        }
                        else
                        {
                            emps.Add(Employers.ElementAt(i).GetFullName() + ";" + Employers.ElementAt(i).Phone + ";" + empShift);
                        }
                    }
                }
            }

            return emps;
        }




        public static void DisplayEmployers()
        {
            foreach (Employer emp in Employers)
            {
                Console.WriteLine(emp.GetFullName());
                Console.WriteLine(emp.Phone);
                for (int i = 0; i < emp.WorkDays.Count; i++)
                {
                    Console.WriteLine(emp.WorkDays.ElementAt(i).Key + ": " + emp.WorkDays.ElementAt(i).Value);
                }
                Console.WriteLine("");
            }
            Console.WriteLine("");
        }




        public static List<string> GetServicePhones()
        {
            StreamReader reader = new StreamReader("telefonok.txt.txt");
            List<string> phones = new List<string>();

            try
            {
                string line;

                while ((line = reader.ReadLine()) != null)
                {
                    string[] row = line.Split(';');
                    if (row[0].Trim().ToLower() == "szolg")
                    {
                        phones.Add(row[1]);
                    }
                }

            }
            catch (IOException e)
            {
                Console.WriteLine(e);
            }

            return phones;
        }





        public static string GetEmployerPhone(string firstName, string lastName)
        {
            StreamReader reader = new StreamReader("telefonok.txt.txt");

            try
            {
                string line;
                string name = (firstName + " " + lastName).Trim().ToLower();

                while ((line = reader.ReadLine()) != null)
                {
                    string[] row = line.Split(';');

                    if (row[0].Trim().ToLower() == name)
                    {
                        return row[1];
                    }
                }

            }
            catch (IOException e)
            {
                Console.WriteLine(e);
            }

            return "-";
        }





        public static List<Employer> GetEmployers()
        {
            StreamReader reader = new StreamReader("betegkiserok.csv");
            List<Employer> emps = new List<Employer>();

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

                        for (int j = 3; j < row.Length; j++)
                        {
                            if (row[j] != "")
                            {
                                workDays.Add((j - 2).ToString(), row[j]);
                            }
                        }

                        Employer employer = new Employer(firstName, lastName, phone, workDays);
                        emps.Add(employer);

                    }
                    i++;
                }

            }
            catch (IOException e)
            {
                Console.WriteLine(e);
            }

            reader.Close();

            return emps;
        }
    }
}
