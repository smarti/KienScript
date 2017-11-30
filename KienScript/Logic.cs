using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace KienScript
{
    class Logic
    {
        private Application OWord;
        private Document document;

        private object missing = System.Reflection.Missing.Value;

        public Logic()
            => CreateDocument();

        private void CreateDocument()
        {
            //create instance of word
            OWord = new Application
            {
                ShowAnimation = false,
                Visible = false
            };

            //create new document
            document = OWord.Documents.Add(ref missing, ref missing, ref missing, ref missing);
        }

        public void GenerateDocument(int amountOfTables)
        {
            for (int i = 0; i < amountOfTables; i++)
            {
                //create new table
                int[,] dataSet = CreateTable();

                //add table to document
                AddTableToDocument(dataSet);
            }

            //make result visible
            OWord.Visible = true;
            OWord.ShowAnimation = true;
        }

        private int[,] CreateTable()
        {
            Random random = new Random();

            int[,] dataSet = new int[3, 9];

            List<int> value = new List<int> { 0, 0, 0 };

            for (int index1 = 0; index1 < 9; index1++)
            {
                do
                {
                    value[0] = random.Next(1, 11);
                    value[1] = random.Next(1, 11);
                    value[2] = random.Next(1, 11);
                } while (value.Count != value.Distinct().Count());

                value.Sort();

                for (int index0 = 0; index0 < 3; index0++)
                {
                    dataSet[index0, index1] = index1 * 10 + value[index0];
                }
            }

            //add blank cells
            for (int row = 0; row < 3; row++)
            {
                int targetAmount = 23 - row * 4;
                int currentAmount;
                do
                {
                    dataSet[row, random.Next(0, 9)] = 0;

                    currentAmount = dataSet.Cast<int>().Count(VARIABLE => VARIABLE > 0);
                } while (currentAmount > targetAmount);
            }

            return dataSet;
        }

        private void AddTableToDocument(int[,] dataSet)
        {
            //initiate table
            Range tableLocation = document.Range(document.Content.End - 1, ref missing);
            document.Tables.Add(tableLocation, 3, 9, ref missing, ref missing);

            //create table
            Table table = document.Tables[document.Tables.Count];
            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    cell.Range.Text = dataSet[cell.RowIndex - 1, cell.ColumnIndex - 1] != 0 ?
                        dataSet[cell.RowIndex - 1, cell.ColumnIndex - 1].ToString() : "";

                    cell.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                }
            }

            //add paragraph on end of each table
            document.Paragraphs.Add(ref missing).Range.InsertParagraphAfter();
        }
    }
}
