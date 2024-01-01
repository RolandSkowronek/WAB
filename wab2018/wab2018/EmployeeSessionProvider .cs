using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.SessionState;

namespace wab2018
{
    public class EmployeeEntry
    {
        public EmployeeEntry()
        {
        }
        public EmployeeEntry(string id, string firstName, string lastName, DateTime hireDate, string employerId)
        {
            ID = id;
            FirstName = firstName;
            LastName = lastName;
            HireDate = hireDate;
            EmployerID = employerId;
        }
        public string ID { get; private set; }
        public string EmployerID { get; private set; }
        public string FirstName { get; private set; }
        public string LastName { get; private set; }
        public DateTime HireDate { get; private set; }
        public void Assign(EmployeeEntry source)
        {
            FirstName = source.FirstName;
            LastName = source.LastName;
            HireDate = source.HireDate;
            EmployerID = source.EmployerID;
        }
    }

    public static class EmployeeSessionProvider
    {
        const string Key = "DxEmployeeSessionProvider";

        static List<EmployeeEntry> CreateData()
        {
            List<EmployeeEntry> result = new List<EmployeeEntry>();
            result.Add(new EmployeeEntry(GenerateNewID(), "Nancy", "Davolio", new DateTime(1992, 5, 1), GetEmployerId(result)));
            result.Add(new EmployeeEntry(GenerateNewID(), "Andrew", "Fuller", new DateTime(1992, 8, 14), GetEmployerId(result)));
            result.Add(new EmployeeEntry(GenerateNewID(), "Janet", "Leverling", new DateTime(1992, 4, 1), GetEmployerId(result)));
            result.Add(new EmployeeEntry(GenerateNewID(), "Margaret", "Peacock", new DateTime(1993, 5, 3), GetEmployerId(result)));
            result.Add(new EmployeeEntry(GenerateNewID(), "Steven", "Buchanan", new DateTime(1993, 10, 17), GetEmployerId(result)));
            result.Add(new EmployeeEntry(GenerateNewID(), "Michael", "Suyama", new DateTime(1993, 10, 17), GetEmployerId(result)));
            result.Add(new EmployeeEntry(GenerateNewID(), "Robert", "King", new DateTime(1994, 1, 2), GetEmployerId(result)));
            result.Add(new EmployeeEntry(GenerateNewID(), "Laura", "Callahan", new DateTime(1994, 3, 5), GetEmployerId(result)));
            result.Add(new EmployeeEntry(GenerateNewID(), "Anne", "Dodsworth", new DateTime(1994, 11, 15), GetEmployerId(result)));
            return result;
        }
        static string GetEmployerId(List<EmployeeEntry> existingEmployees)
        {
            if (!existingEmployees.Any())
                return "";
            return existingEmployees[(int)(existingEmployees.Count / 3)].ID;
        }
        static string GenerateNewID()
        {
            return Guid.NewGuid().ToString();
        }

        public static IEnumerable<EmployeeEntry> Select()
        {
            HttpSessionState session = HttpContext.Current.Session;
            if (session[Key] == null)
                session[Key] = CreateData();
            return (List<EmployeeEntry>)session[Key];
        }
    }
}