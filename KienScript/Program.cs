using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net.Cache;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace KienScript
{
    class Program
    {
        static void Main(string[] args)
        {
            string amountOfTables;

            do
            {
                Console.Clear();
                Console.Write("Geef het aantal kienkaarten: ");
                amountOfTables = Console.ReadLine();
            } while (!int.TryParse(amountOfTables, out int result));


            try
            {
                Logic logic = new Logic();

                logic.GenerateDocument(Convert.ToInt32(amountOfTables));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }
        }
    }
}
