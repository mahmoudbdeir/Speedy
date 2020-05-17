// Author: Pardi Bedirian

using System;
using Mastermind.MsOffice;
using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace GetMethodsUnitTester
{
    [TestClass]
    public class UnitTest
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
            int Row = 1;
            int Col = 1;

            sheet.SetRange(Row, Col, array);

        }
        
        [ClassCleanup]
        public static void CloseExcel()
        {
            book.Close(false);
            excel.Quit();
            excel = null;
        }

        [TestMethod]
        public void GetFirstValue_Row1Col1Val0_ReturnAreEqual()
        {
            double value = 0;
            int row = 1;
            int col = 1;
            Assert.AreEqual(sheet.GetValue(row, col), value);
        }

        [TestMethod]
        public void GetLastValue_Row50Col100Val4999_ReturnAreEqual()
        {
            double value = 4999;
            int row = 50;
            int col = 100;
            Assert.AreEqual(sheet.GetValue(row, col), value);
        }

        [TestMethod]
        public void GetFirstValueInRow_Row6Col1Val500_ReturnAreEqual()
        {
            double value = 500;
            int row = 6;
            int col = 1;
            Assert.AreEqual(sheet.GetValue(row, col), value);
        }
        [TestMethod]
        public void GetFirstValueInCol_Row1Col33Val32_ReturnAreEqual()
        {
            double value = 32;
            int row = 1;
            int col = 33;
            Assert.AreEqual(sheet.GetValue(row, col), value);
        }

        [TestMethod]
        public void GetValueInFirstRow_ColCRVal95_ReturnAreEqual()
        {
            double value = 95;
            int row = 1;
            string col = "CR";
            Assert.AreEqual(sheet.GetValue(row, col), value);
        }

        [TestMethod]
        public void GetValueByColName_ColCvVal4999_ReturnAreEqual()
        {
            double value = 4999;
            int row = 50;
            var v = sheet.GetValue(row, "CV");
            Assert.AreEqual(v, value);
        }

        [TestMethod]
        [ExpectedException(typeof(NullReferenceException))]
        public void GetInvalidValue_RowNeg3Col0Val2_ReturnException()
        {
            double value = 2;
            int row = -3;
            int col = 0;
            Assert.AreEqual(sheet.GetValue(row, col), value);
        }

        [TestMethod]
        [ExpectedException(typeof(NullReferenceException))]
        public void GetInvalidValue_Row2ColNeg3Val12_ReturnException()
        {
            double value = 12;
            int row = 2;
            int col = -3;
            Assert.AreEqual(sheet.GetValue(row, col), value);
        }

        [TestMethod]
        [ExpectedException(typeof(NullReferenceException))]
        public void GetInvalidValue_RowNeg300ColNeg8ValNeg100_ReturnException()
        {
            double value = -100;
            int row = -300;
            int col = -8;
            Assert.AreEqual(sheet.GetValue(row, col), value);
        }

        [TestMethod]
        [ExpectedException(typeof(RuntimeBinderException))]
        public void GetValueOutsideRange_Row50Col101Val22_ReturnException()
        {
            double value = 22;
            int row = 50;
            int col = 101;
            Assert.AreEqual(sheet.GetValue(row, col), value);
        }

        [TestMethod]
        [ExpectedException(typeof(RuntimeBinderException))]
        public void GetValueOutsideRange_Row51Col100Val44_ReturnException()
        {
            double value = 44;
            int row = 51;
            int col = 100;
            Assert.AreEqual(sheet.GetValue(row, col), value);
        }

        [TestMethod]
        [ExpectedException(typeof(RuntimeBinderException))]
        public void GetValueOutsideRange_Row1011Col789Val122_ReturnException()
        {
            double value = 122;
            int row = 1011;
            int col = 789;
            Assert.AreEqual(sheet.GetValue(row, col), value);
        }

        [TestMethod]
        [ExpectedException(typeof(RuntimeBinderException))]
        public void GetValueOutsideRange_Row90ColczVal115_ReturnException()
        {
            double value = 115;
            int row = 90;
            string col = "CZ";
            Assert.AreEqual(sheet.GetValue(row, col), value);
        }

        [TestMethod]
        public void Get_Row25ColaVal2400_ReturnAreEqual()
        {
            double value = 2400;
            int row = 25;
            string col = "A";
            Assert.AreEqual(sheet.GetValue(row, col), value);
        }

        [TestMethod]
        public void GetBoundaries_Row10Col5Vals_ReturnAreEqual()
        {
            double valN1 = 903; //values of the nearests (boundaries)
            double valN2 = 905;
            double valN3 = 804;
            double valN4 = 1004;

            int row = 10;
            int col = 5;
            Assert.AreEqual(sheet.GetValue(row, col - 1), valN1);
            Assert.AreEqual(sheet.GetValue(row, col + 1), valN2);
            Assert.AreEqual(sheet.GetValue(row - 1, col), valN3);
            Assert.AreEqual(sheet.GetValue(row + 1, col), valN4);
        }

        [TestMethod]
        public void GetBoundaries_Col6Vals_ReturnAreEqual()
        {
            int colRank = 6;
            Assert.AreEqual(sheet.GetValue(5, colRank - 1), 404.0);
            Assert.AreEqual(sheet.GetValue(5, colRank + 1), 406.0);
            Assert.AreEqual(sheet.GetValue(1, colRank - 1), 4.0);
            Assert.AreEqual(sheet.GetValue(1, colRank + 1), 6.0);
            Assert.AreEqual(sheet.GetValue(50, colRank - 1), 4904.0);
            Assert.AreEqual(sheet.GetValue(50, colRank + 1), 4906.0);
        }

        [TestMethod]
        public void GetBoundaries_Row20Valnull_ReturnAreEqual()
        {
            int rowRank = 20;


            Assert.AreEqual(sheet.GetValue(rowRank - 1, 5), 1804.0);
            Assert.AreEqual(sheet.GetValue(rowRank + 1, 5), 2004.0);
            Assert.AreEqual(sheet.GetValue(rowRank - 1, 1), 1800.0);
            Assert.AreEqual(sheet.GetValue(rowRank + 1, 1), 2000.0);
            Assert.AreEqual(sheet.GetValue(rowRank - 1, 6), 1805.0);
            Assert.AreEqual(sheet.GetValue(rowRank + 1, 6), 2005.0);
        }

        [TestMethod]
        public void GetValues_Col56Val55To955_ReturnAreEqual()
        {
            int colRank = 56;

            Assert.AreEqual(sheet.GetValue(1, colRank), 55.0);
            Assert.AreEqual(sheet.GetValue(5, colRank), 455.0);
            Assert.AreEqual(sheet.GetValue(10, colRank), 955.0);
        }

        [TestMethod]
        public void GetValues_Row20Val1906To1915_ReturnAreEqual()
        {
            int rowRank = 20;

            Assert.AreEqual(sheet.GetValue(rowRank, 7), 1906.0);
            Assert.AreEqual(sheet.GetValue(rowRank, 10), 1909.0);
            Assert.AreEqual(sheet.GetValue(rowRank, 15), 1914.0);
            Assert.AreEqual(sheet.GetValue(rowRank, 16), 1915.0);
        }

        [TestMethod]
        public void GetRandomValues_RowRandomColRandom_ReturnAreEqual()
        {
            for (int i = 0; i < 10; i++)
            {
                var random = new Random();
                int row = random.Next(1, 51);
                int col = random.Next(1, 101);
                double value = (row - 1) * 100 + col - 1;

                Assert.AreEqual(sheet.GetValue(row, col), value);
            }
        }

        [TestMethod]
        public void GetRandomCol_ColRandom_ReturnAreEqual()
        {
            var random = new Random();
            int colRank = random.Next(1, 101);

            Assert.AreEqual(sheet.GetValue(1, colRank), (double)(colRank - 1));
            Assert.AreEqual(sheet.GetValue(16, colRank), (double)(1500 + colRank - 1));
            Assert.AreEqual(sheet.GetValue(40, colRank), (double)(3900 + colRank - 1));
            Assert.AreEqual(sheet.GetValue(50, colRank), (double)(4900 + colRank - 1));
        }

        [TestMethod]
        public void GetRandomRow_RowRandom_ReturnAreEqual()
        {
            var random = new Random();
            int rowRank = random.Next(1, 51);

            Assert.AreEqual(sheet.GetValue(rowRank, 1), (double)((rowRank - 1) * 100));
            Assert.AreEqual(sheet.GetValue(rowRank, 43), (double)((rowRank - 1) * 100 + 42));
            Assert.AreEqual(sheet.GetValue(rowRank, 32), (double)((rowRank - 1) * 100 + 31));
            Assert.AreEqual(sheet.GetValue(rowRank, 100), (double)((rowRank - 1) * 100 + 99));
        }
    }
}
