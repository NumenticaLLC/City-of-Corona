using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            string location = "https://coronaca.sharepoint.com/:f:/r/sites/CityClerk_Subpoena_Sandbox/incident/ABCDE_2808d0f7aac5eb11bacd001dd804b33e";
            if (!string.IsNullOrEmpty(location))
            {
                string[] abc = location.Split('/');
                string foldername = abc[abc.Length - 1];

            }
        }
    }
}
