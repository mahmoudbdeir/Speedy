using Microsoft.Office.Interop.Excel;
using System;

namespace Mastermind.MsOffice
{
    public static class Excellinator
    {
        #region SetRange
        /// <summary>
        /// Sets the value of every cell in a range given an array of objects and the starting cell coordinates
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startRow">Starting row (first row = 1)</param>
        /// <param name="startCol">Starting column (first column = 1)</param>
        /// <param name="array">Array of objects to copy to Excel sheet starting at (startRow, startCol)</param>
        public static void SetRange(this Worksheet sheet, int startRow, int startCol, object[,] array)
        {
            if (startRow < 1 || startCol < 1 || array == null)
                return;

            Range TopLeftCell = (Range)sheet.Cells[startRow, startCol];
            Range BottomRightCell = (Range)sheet.Cells[startRow+ array.GetLength(0) - 1, startCol + array.GetLength(1)-1];
            Range range = sheet.get_Range(TopLeftCell, BottomRightCell);
            range.Value = array;
        }
        /// <summary>
        /// Sets the value of every cell in a range given an array of objects and the starting cell coordinates
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startRow">Starting row (first row = 1)</param>
        /// <param name="startCol">Starting column (first column = "A")</param>
        /// <param name="array">Array of objects to copy to Excel sheet starting at (startRow, startCol)</param>
        public static void SetRange(this Worksheet sheet, int startRow, string startCol, object[,] arr) => SetRange(sheet, startRow, ConvertColumn(startCol), arr);
        #endregion

        #region GetRange


        /// <summary>
        /// Gets the value of every cell in an Excel range given the coordinates of the starting cell (top-left) and ending cell (bottom-right)
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startRow">Starting row (first row = 1)</param>
        /// <param name="startCol">Starting column (first column = 1)</param>
        /// <param name="rows">Number of rows</param>
        /// <param name="cols">Number of columns</param>
        /// <returns>A two-dimensional array of objects</returns>
        public static object[,] GetRange(this Worksheet sheet, int startRow, int startCol, int rows, int cols)
        {
            if (startRow < 1 || startCol < 1 || rows < 1 || cols < 1)
                return null;

            Range TopLeftCell = (Range)sheet.Cells[startRow, startCol];
            Range BottomRightCell = (Range)sheet.Cells[startRow+ rows-1, startCol + cols-1];
            Range range = sheet.get_Range(TopLeftCell, BottomRightCell);
            if(range.Value2.GetType()==typeof(object[,]))
            {
                return (object[,])range.Value;
            }
            else
            {
                var arr = new object[1,1];
                arr[0, 0] = range.Value2;
                return arr;
            }
        }
        /// <summary>
        /// Gets the value of every cell in an Excel range given the coordinates of the starting cell (top-left) and ending cell (bottom-right)
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startRow">Starting row (first row = 1)</param>
        /// <param name="startCol">Starting column (first column = "A")</param>
        /// <param name="rows">Number of rows</param>
        /// <param name="cols">Number of columns</param>
        /// <returns>A two-dimensional array of objects</returns>
        public static object[,] GetRange(this Worksheet sheet, int startRow, string startCol, int rows, int cols) => GetRange(sheet, startRow, ConvertColumn(startCol), rows, cols);

        /// <summary>
        /// Gets the value of a single cell in an Excel range given the coordinates of the cell
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row">The row number (first row = 1)</param>
        /// <param name="col">The column name (first column = "A")</param>
        /// <returns>The value of the cell located at (row,col)</returns>
        public static object GetValue(this Worksheet sheet, int row, string col) => GetValue(sheet, row, ConvertColumn(col));

        /// <summary>
        /// Gets the value of a single cell in an Excel range given the coordinates of the cell
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row">The row number (first row = 1)</param>
        /// <param name="col">The column number (first column = 1)</param>
        /// <returns>The value of the cell located at (row,col)</returns>
        public static object GetValue(this Worksheet sheet, int row, int col) => GetRange(sheet, row, col, 1, 1)[0,0];
        #endregion

        #region Helper Methods
        private static int ConvertColumn(string s)
        {
            int col = 0;
            s = s.ToUpper();
            for (int i = 0; i < s.Length; i++)
            {
                col += (s[i] - 64) * (int)Math.Pow(26, s.Length - 1 - i);
            }
            return col;
        } 
        #endregion
    }
}