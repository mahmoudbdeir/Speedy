// Author: Mohammed Ezzedine

using System;
using Mastermind.MsOffice;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SetMethodsUnitTester
{
    [TestClass]
    public class AlphaValuesUnitTesting
    {
        private static Microsoft.Office.Interop.Excel.Application excel;

        private static Workbook book;
        private static Worksheet sheet;

        private static char[] letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();
        private static string[] words;

        [ClassInitialize]
        public static void InitializeContent(TestContext context)
        {
            excel = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = true
            };

            book = excel.Workbooks.Add();
            sheet = book.Sheets[1];

            words = new string[100];
            var random = new Random();
            for (int i = 0; i < 100; i++)
            {
                // Make a word of 5 letters.
                string word = "";
                for (int j = 1; j <= 5; j++)
                {
                    int letter_num = random.Next(0, letters.Length - 1);

                    // Append the letter.
                    word += letters[letter_num];
                }
                words[i] = word;
            }
        }

        [ClassCleanup]
        public static void CloseExcel()
        {
            book.Close(false);
            excel.Quit();
            excel = null;
        }

        [TestMethod]
        public void SetAndGetFirstValue_Row1Col1Valabc_ReturnAreEqual()
        {
            string value = "abc";
            int row = 1;
            int col = 1;

            object[,] block = new object[1, 1];
            block[0, 0] = value;

            sheet.SetRange(row, col, block);

            Assert.AreEqual(sheet.GetValue(row, col), value);
        }

        [TestMethod]
        public void SetColAndGet_Col857ValRandomWords_ReturnAreEqual()
        {
            int colRank = 857;

            object[,] column = new object[50, 1];

            for (int _row = 0; _row < 50; _row++)
            {
                column[_row, 0] = words[_row];
            }

            sheet.SetRange(1, colRank, column);

            Assert.AreEqual(sheet.GetValue(1, colRank), words[0]);
            Assert.AreEqual(sheet.GetValue(13, colRank), words[12]);
            Assert.AreEqual(sheet.GetValue(33, colRank), words[32]);
            Assert.AreEqual(sheet.GetValue(47, colRank), words[46]);
            Assert.AreEqual(sheet.GetValue(50, colRank), words[49]);
        }

        [TestMethod]
        public void SetRowAndGet_Row28ValRandomWords_ReturnAreEqual()
        {
            int rowRank = 28;

            object[,] row = new object[1, 100];

            for (int _col = 0; _col < 100; _col++)
            {
                row[0, _col] = words[_col];
            }

            sheet.SetRange(rowRank, 1, row);

            Assert.AreEqual(sheet.GetValue(rowRank, 1), words[0]);
            Assert.AreEqual(sheet.GetValue(rowRank, 8), words[7]);
            Assert.AreEqual(sheet.GetValue(rowRank, 15), words[14]);
            Assert.AreEqual(sheet.GetValue(rowRank, 29), words[28]);
            Assert.AreEqual(sheet.GetValue(rowRank, 46), words[45]);
            Assert.AreEqual(sheet.GetValue(rowRank, 68), words[67]);
            Assert.AreEqual(sheet.GetValue(rowRank, 99), words[98]);
            Assert.AreEqual(sheet.GetValue(rowRank, 100), words[99]);
        }

        [TestMethod]
        public void SetAndGetRandomValues_RowRandomColRandomValRandomWords_ReturnAreEqual()
        {
            for (int i = 0; i < 10; i++)
            {
                var random = new Random();
                int row = random.Next(1, 1000);
                int col = random.Next(1, 1000);

                int index = random.Next(0, 100);

                object[,] block = new object[1, 1];
                block[0, 0] = words[index];

                sheet.SetRange(row, col, block);

                Assert.AreEqual(sheet.GetValue(row, col), words[index]);
            }
        }

        [TestMethod]
        public void SetAndGetRandomCol_ColRandomValRandomWords_ReturnAreEqual()
        {
            var random = new Random();
            int colRank = random.Next(1, 1000);

            object[,] column = new object[50, 1];

            for (int _row = 0; _row < 50; _row++)
            {
                column[_row, 0] = words[_row];
            }

            sheet.SetRange(1, colRank, column);

            Assert.AreEqual(sheet.GetValue(1, colRank), words[0]);
            Assert.AreEqual(sheet.GetValue(24, colRank), words[23]);
            Assert.AreEqual(sheet.GetValue(33, colRank), words[32]);
            Assert.AreEqual(sheet.GetValue(50, colRank), words[49]);
        }

        [TestMethod]
        public void SetAndGetRandomRow_RowRandomValRandomWords_ReturnAreEqual()
        {
            var random = new Random();
            int rowRank = random.Next(1, 1000);

            object[,] row = new object[1, 100];

            for (int _col = 0; _col < 100; _col++)
            {
                row[0, _col] = words[_col];
            }

            sheet.SetRange(rowRank, 1, row);

            Assert.AreEqual(sheet.GetValue(rowRank, 1), words[0]);
            Assert.AreEqual(sheet.GetValue(rowRank, 23), words[22]);
            Assert.AreEqual(sheet.GetValue(rowRank, 57), words[56]);
            Assert.AreEqual(sheet.GetValue(rowRank, 100), words[99]);
        }

        [TestMethod]
        public void SetAndGetRandomGrid_RowRandomColRandomValRandomWords_ReturnAreEqual()
        {
            var random = new Random();

            object[,] array = new object[12, 8];

            for (int row = 0; row < array.GetLength(0); row++)
            {
                for (int col = 0; col < array.GetLength(1); col++)
                {
                    array[row, col] = words[row * array.GetLength(1) + col];
                }
            }

            sheet.SetRange(1, 1, array);

            Assert.AreEqual(sheet.GetValue(1, 1), words[0]);
            Assert.AreEqual(sheet.GetValue(6, 3), words[42]);
            Assert.AreEqual(sheet.GetValue(10, 5), words[76]);
            Assert.AreEqual(sheet.GetValue(2, 7), words[14]);
        }
    }
}
