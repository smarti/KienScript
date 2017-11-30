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
                Console.Write("Geef het aantal kienkaarten: ");
                amountOfTables = Console.ReadLine();
            } while (!int.TryParse(amountOfTables, out int result));


            //try
            //{
            Logic logic = new Logic();

            logic.GenerateDocument(Convert.ToInt32(amountOfTables));
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine(ex.Message);
            //}
        }


        //private static void CreateFile(int requestedAmount)
        //{
        //    //TODO: standaard 15 cijfers laten printen
        //    try
        //    {
        //        //create instance of word app
        //        Application winword = new Application
        //        {

        //            //set animation status for word application
        //            ShowAnimation = false,
        //            //set status for word application is to be visible or not
        //            Visible = false
        //        };

        //        //create a missing variable for missing value
        //        object missing = System.Reflection.Missing.Value;

        //        //create new word document
        //        Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);

        //        //Ensure randomness per card
        //        Random random = new Random();
        //        List<string> usedValues = new List<string>();


        //        for (int currentAmount = 0; currentAmount < requestedAmount; currentAmount++)
        //        {

        //            //initiate table
        //            Range tableLocation = document.Range(document.Content.End - 1, ref missing);
        //            document.Tables.Add(tableLocation, 3, 9, ref missing, ref missing);

        //            //create table
        //            Table table = document.Tables[document.Tables.Count];
        //            foreach (Row row in table.Rows)
        //            {
        //                foreach (Cell cell in row.Cells)
        //                {
        //                    //create random value
        //                    int secondInt = random.Next(0, 13);
        //                    if (secondInt == 0 | secondInt >= 11)
        //                        cell.Range.Text = "";

        //                    else if (secondInt == 10)
        //                        cell.Range.Text = (cell.ColumnIndex).ToString() + "0";

        //                    else if (cell.ColumnIndex == 1)
        //                    {
        //                        cell.Range.Text = secondInt.ToString();
        //                    }
        //                    else
        //                    {
        //                        cell.Range.Text = ((cell.ColumnIndex - 1).ToString()) + (secondInt.ToString());
        //                    }

        //                    if (usedValues.Contains(cell.Range.Text))
        //                        cell.Range.Text = "";

        //                    usedValues.Add(cell.Range.Text);
        //                    cell.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
        //                }
        //            }

        //            usedValues.Clear();
        //            document.Paragraphs.Add(ref missing).Range.InsertParagraphAfter();
        //        }

        //        winword.Visible = true;
        //    }

        //    catch (Exception ex)
        //    {
        //        Console.WriteLine(ex.Message);
        //        Console.ReadKey();
        //    }
        //}

    }
}
