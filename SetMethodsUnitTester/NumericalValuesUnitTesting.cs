// Author: Mohammed

using System;
using Mastermind.MsOffice;
using Microsoft.Office.Interop.Excel;
using Microsoft.CSharp.RuntimeBinder;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SetMethodsUnitTester
{
    [TestClass]
    public class NumericalValuesUnitTesting
    {
        private static Microsoft.Office.Interop.Excel.Application excel;

        private static Workbook book;
        private static Worksheet sheet;

        private static int Rows;
        private static int Columns;

        [ClassInitialize]
        public static void InitializeContent(TestContext context)
        {
            excel = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = true
            };

            book = excel.Workbooks.Add();
            sheet = book.Sheets[1];
        }

        [TestInitialize]
        public void FillData()
        {
            Rows = 50;
            Columns = 100;

            object[,] array = new object[Rows, Columns];

            for (int row = 0; row < array.GetLength(0); row++)
            {
                for (int col = 0; col < array.GetLength(1); col++)
                {
                    array[row, col] = row * array.GetLength(1) + col;
                }
            }

            sheet.SetRange(1, 1, array);
        }

        [ClassCleanup]
        public static void CloseExcel()
        {
            book.Close(false);
            excel.Quit();
            excel = null;
        }

        [TestMethod]
        public void SetAndGetFirstValue_Row1Col1ValNeg1_ReturnAreEqual()
        {
            double value = -1;
            int row = 1;
            int col = 1;

            object[,] block = new object[1, 1];
            block[0, 0] = value;

            sheet.SetRange(row, col, block);

            Assert.AreEqual(sheet.GetValue(row, col), value);
        }

        [TestMethod]
        public void SetAndGetLastValue_Row50Col100ValNeg86582_ReturnAreEqual()
        {
            double value = -86582;
            int row = 50;
            int col = 100;

            object[,] block = new object[1, 1];
            block[0, 0] = value;

            sheet.SetRange(row, col, block);

            Assert.AreEqual(sheet.GetValue(row, col), value);
        }

        [TestMethod]
        public void SetAndGetFirstValueInRow_Row6Col1Val10000_ReturnAreEqual()
        {
            double value = 10000;
            int row = 6;
            int col = 1;

            object[,] block = new object[1, 1];
            block[0, 0] = value;

            sheet.SetRange(row, col, block);

            Assert.AreEqual(sheet.GetValue(row, col), value);
        }

        [TestMethod]
        public void SetAndGetLastValueInRow_Row31Col100Val75_ReturnAreEqual()
        {
            double value = 75;
            int row = 31;
            int col = 100;

            object[,] block = new object[1, 1];
            block[0, 0] = value;

            sheet.SetRange(row, col, block);

            Assert.AreEqual(sheet.GetValue(row, col), value);
        }

        [TestMethod]
        public void SetAndGetFirstValueInCol_Row1Col45Val0_ReturnAreEqual()
        {
            double value = 0;
            int row = 1;
            int col = 45;

            object[,] block = new object[1, 1];
            block[0, 0] = value;

            sheet.SetRange(row, col, block);

            Assert.AreEqual(sheet.GetValue(row, col), value);
        }

        [TestMethod]
        public void SetAndGetLastValueInCol_Row50Col73Val9572_ReturnAreEqual()
        {
            double value = 9572;
            int row = 50;
            int col = 73;

            object[,] block = new object[1, 1];
            block[0, 0] = value;

            sheet.SetRange(row, col, block);

            Assert.AreEqual(sheet.GetValue(row, col), value);
        }

        [TestMethod]
        [ExpectedException(typeof(NullReferenceException))]
        public void SetAndGetInvalidValue_Row0Col0Val10_ReturnException()
        {
            double value = 10;
            int row = 0;
            int col = 0;

            object[,] block = new object[1, 1];
            block[0, 0] = value;

            sheet.SetRange(row, col, block);

            Assert.AreEqual(sheet.GetValue(row, col), value);
        }

        [TestMethod]
        [ExpectedException(typeof(NullReferenceException))]
        public void SetAndGetInvalidValue_RowNeg1Col0Val7_ReturnException()
        {
            double value = 7;
            int row = -1;
            int col = 0;

            object[,] block = new object[1, 1];
            block[0, 0] = value;

            sheet.SetRange(row, col, block);

            Assert.AreEqual(sheet.GetValue(row, col), value);
        }

        [TestMethod]
        [ExpectedException(typeof(NullReferenceException))]
        public void SetAndGetInvalidValue_Row5ColNeg3Val64_ReturnException()
        {
            double value = 64;
            int row = 5;
            int col = -3;

            object[,] block = new object[1, 1];
            block[0, 0] = value;

            sheet.SetRange(row, col, block);

            Assert.AreEqual(sheet.GetValue(row, col), value);
        }

        [TestMethod]
        [ExpectedException(typeof(NullReferenceException))]
        public void SetAndGetInvalidValue_RowNeg900ColNeg7ValNeg55_ReturnException()
        {
            double value = -55;
            int row = -900;
            int col = -7;

            object[,] block = new object[1, 1];
            block[0, 0] = value;

            sheet.SetRange(row, col, block);

            Assert.AreEqual(sheet.GetValue(row, col), value);
        }

        [TestMethod]
        public void SetAndGetValueOutsideRange_Row50Col101Val13_ReturnAreEqual()
        {
            double value = 13;
            int row = 50;
            int col = 101;

            object[,] block = new object[1, 1];
            block[0, 0] = value;

            sheet.SetRange(row, col, block);

            Assert.AreEqual(sheet.GetValue(row, col), value);
        }

        [TestMethod]
        public void SetAndGetValueOutsideRange_Row51Col100Val46_ReturnAreEqual()
        {
            double value = 46;
            int row = 51;
            int col = 100;

            object[,] block = new object[1, 1];
            block[0, 0] = value;

            sheet.SetRange(row, col, block);

            Assert.AreEqual(sheet.GetValue(row, col), value);
        }

        [TestMethod]
        public void SetAndGetValueOutsideRange_Row1005Col999Val154_ReturnAreEqual()
        {
            double value = 154;
            int row = 1005;
            int col = 999;

            object[,] block = new object[1, 1];
            block[0, 0] = value;

            sheet.SetRange(row, col, block);

            Assert.AreEqual(sheet.GetValue(row, col), value);
        }

        [TestMethod]
        public void SetAndGetValueOutsideRange_Row89ColfhVal12345_ReturnAreEqual()
        {
            double value = 12345;
            int row = 89;
            string col = "fh";

            object[,] block = new object[1, 1];
            block[0, 0] = value;

            sheet.SetRange(row, col, block);

            Assert.AreEqual(sheet.GetValue(row, col), value);
        }

        [TestMethod]
        [ExpectedException(typeof(RuntimeBinderException))]
        public void SetNullAndGet_Row30ColgValnull_ReturnException()
        {
            string value = null;
            int row = 30;
            string col = "g";

            object[,] block = new object[1, 1];
            block[0, 0] = value;

            sheet.SetRange(row, col, block);

            Assert.AreEqual(sheet.GetValue(row, col), value);
        }

        [TestMethod]
        public void SetNullAndGetBoundaries_Row30Col15Valnull_ReturnAreEqual()
        {
            string value = null;
            int row = 30;
            int col = 15;

            object[,] block = new object[1, 1];
            block[0, 0] = value;

            sheet.SetRange(row, col, block);

            Assert.AreEqual(sheet.GetValue(row, col - 1), (double)2913);
            Assert.AreEqual(sheet.GetValue(row, col + 1), (double)2915);
            Assert.AreEqual(sheet.GetValue(row - 1, col), (double)2814);
            Assert.AreEqual(sheet.GetValue(row + 1, col), (double)3014);
        }

        [TestMethod]
        [ExpectedException(typeof(RuntimeBinderException))]
        public void SetColNullAndGet_Col17Valnull_ReturnException()
        {
            string value = null;
            int colRank = 17;

            object[,] column = new object[50, 1];

            for (int _row = 0; _row < 50; _row++)
            {
                column[_row, 0] = value;
            }

            sheet.SetRange(1, colRank, column);

            Assert.AreEqual(sheet.GetValue(5, colRank), value);
            Assert.AreEqual(sheet.GetValue(1, colRank), value);
            Assert.AreEqual(sheet.GetValue(50, colRank), value);
            Assert.AreEqual(sheet.GetValue(51, colRank), value);
        }

        [TestMethod]
        public void SetColNullAndGetBoundaries_Col6Valnull_ReturnAreEqual()
        {
            string value = null;
            int colRank = 6;

            object[,] column = new object[50, 1];

            for (int _row = 0; _row < 50; _row++)
            {
                column[_row, 0] = value;
            }

            sheet.SetRange(1, colRank, column);

            Assert.AreEqual(sheet.GetValue(5, colRank - 1), (double)404);
            Assert.AreEqual(sheet.GetValue(5, colRank + 1), (double)406);
            Assert.AreEqual(sheet.GetValue(1, colRank - 1), (double)4);
            Assert.AreEqual(sheet.GetValue(1, colRank + 1), (double)6);
            Assert.AreEqual(sheet.GetValue(50, colRank - 1), (double)4904);
            Assert.AreEqual(sheet.GetValue(50, colRank + 1), (double)4906);
        }

        [TestMethod]
        [ExpectedException(typeof(RuntimeBinderException))]
        public void SetRowNullAndGet_Row32Valnull_ReturnException()
        {
            string value = null;
            int rowRank = 32;

            object[,] row = new object[1, 100];

            for (int _col = 0; _col < 100; _col++)
            {
                row[0, _col] = value;
            }

            sheet.SetRange(rowRank, 1, row);

            Assert.AreEqual(sheet.GetValue(rowRank, 5), value);
            Assert.AreEqual(sheet.GetValue(rowRank, 1), value);
            Assert.AreEqual(sheet.GetValue(rowRank, 50), value);
            Assert.AreEqual(sheet.GetValue(rowRank, 51), value);
        }

        [TestMethod]
        public void SetRowNullAndGetBoundaries_Row20Valnull_ReturnAreEqual()
        {
            string value = null;
            int rowRank = 20;

            object[,] row = new object[1, 100];

            for (int _col = 0; _col < 100; _col++)
            {
                row[0, _col] = value;
            }

            sheet.SetRange(rowRank, 1, row);

            Assert.AreEqual(sheet.GetValue(rowRank - 1, 5), (double)1804);
            Assert.AreEqual(sheet.GetValue(rowRank + 1, 5), (double)2004);
            Assert.AreEqual(sheet.GetValue(rowRank - 1, 1), (double)1800);
            Assert.AreEqual(sheet.GetValue(rowRank + 1, 1), (double)2000);
            Assert.AreEqual(sheet.GetValue(rowRank - 1, 50), (double)1849);
            Assert.AreEqual(sheet.GetValue(rowRank + 1, 50), (double)2049);
        }

        [TestMethod]
        public void SetColAndGet_Col67Val1To5_ReturnAreEqual()
        {
            int colRank = 67;

            object[,] column = new object[5, 1];

            for (int _row = 0; _row < 5; _row++)
            {
                column[_row, 0] = _row;
            }

            sheet.SetRange(1, colRank, column);

            Assert.AreEqual(sheet.GetValue(1, colRank), (double)0);
            Assert.AreEqual(sheet.GetValue(3, colRank), (double)2);
            Assert.AreEqual(sheet.GetValue(5, colRank), (double)4);
        }

        [TestMethod]
        public void SetRowAndGet_Row42ValNeg10toNeg19_ReturnAreEqual()
        {
            int rowRank = 20;

            object[,] row = new object[1, 10];

            for (int _col = 0; _col < 10; _col++)
            {
                row[0, _col] = -10 + _col * -1;
            }

            sheet.SetRange(rowRank, 7, row);

            Assert.AreEqual(sheet.GetValue(rowRank, 7), (double)-10);
            Assert.AreEqual(sheet.GetValue(rowRank, 10), (double)-13);
            Assert.AreEqual(sheet.GetValue(rowRank, 15), (double)-18);
            Assert.AreEqual(sheet.GetValue(rowRank, 16), (double)-19);
        }

        [TestMethod]
        public void SetAndGetRandomValues_RowRandomColRandomValRandom_ReturnAreEqual()
        {
            for (int i = 0; i < 10; i++)
            {
                var random = new Random();
                double value = (double)random.Next(-1000, 1000);
                int row = random.Next(1, 1000);
                int col = random.Next(1, 1000);

                object[,] block = new object[1, 1];
                block[0, 0] = value;

                sheet.SetRange(row, col, block);

                Assert.AreEqual(sheet.GetValue(row, col), value);
            }
        }

        [TestMethod]
        public void SetAndGetRandomCol_ColRandomValRandom_ReturnAreEqual()
        {
            var random = new Random();
            double value = (double)random.Next(-1000, 1000);
            int colRank = random.Next(1, 1000);

            object[,] column = new object[50, 1];

            for (int _row = 0; _row < 50; _row++)
            {
                column[_row, 0] = value;
            }

            sheet.SetRange(1, colRank, column);

            Assert.AreEqual(sheet.GetValue(1, colRank), value);
            Assert.AreEqual(sheet.GetValue(24, colRank), value);
            Assert.AreEqual(sheet.GetValue(33, colRank), value);
            Assert.AreEqual(sheet.GetValue(50, colRank), value);
        }

        [TestMethod]
        public void SetAndGetRandomRow_RowRandomValRandom_ReturnAreEqual()
        {
            var random = new Random();
            double value = (double)random.Next(-1000, 1000);
            int rowRank = random.Next(1, 1000);

            object[,] row = new object[1, 100];

            for (int _col = 0; _col < 100; _col++)
            {
                row[0, _col] = value;
            }

            sheet.SetRange(rowRank, 1, row);

            Assert.AreEqual(sheet.GetValue(rowRank, 1), value);
            Assert.AreEqual(sheet.GetValue(rowRank, 23), value);
            Assert.AreEqual(sheet.GetValue(rowRank, 57), value);
            Assert.AreEqual(sheet.GetValue(rowRank, 100), value);
        }
    }
}
