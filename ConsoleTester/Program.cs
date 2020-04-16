using Mastermind.MsOffice;
using Microsoft.Office.Interop.Excel;
using System;

namespace ConsoleApp2
{
    class Program
    {
        static void Main(string[] args)
        {
            const int Rows = 50;
            const int Columns = 100;

            object[,] array = new object[Rows,Columns];

            for (int row = 0; row < array.GetLength(0); row++)
            {
                for (int col = 0; col < array.GetLength(1); col++)
                {
                    array[row, col] = row * array.GetLength(1) + col;
                }
            }

            var excel = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = true
            };
            Workbook book = excel.Workbooks.Add();
            Worksheet sheet = book.Sheets[1];

            int Row = 1;
            int Col = 5;

            sheet.SetRange(Row, Col, array);

            object[,] obj = sheet.GetRange(Row,Col+Columns-1,Rows,1);
            Console.WriteLine(obj[Rows, 1]); //returns 4999

            obj = sheet.GetRange(1, "CZ", 1, 1); 
            Console.WriteLine(obj[0,0]); //returns 99


            //get value of a single cell
            object v = sheet.GetValue(Rows, Columns+Col-1);
            Console.WriteLine(v); //returns 4999

            //get value of a single cell by column name
            v = sheet.GetValue(Rows, "cz");
            Console.WriteLine(v); //returns 4999

            book.Close(false);
            excel.Quit();
            excel = null;
        }
    }
}