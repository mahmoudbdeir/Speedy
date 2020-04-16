# Speedy
*Speedy* is an Ultra-lightweight .NET library to read/write data to Excel using Interop.
It adds extension methods to *Interops'* ```Sheet``` class to quickly read/write data from/to Microsoft Excel. 


#### Usage

```C#
using Mastermind.MsOffice;
......
// To copy data to Excel using an extension method:
Worksheet sheet;
object[,] array = new object[NumOfRows,NumOfColumns];
    //.....fill the array
sheet.SetRange(StartRow, StartCol, array);

// To read data from Excel using an extension method:
object[,] obj = sheet.GetRange(StartRow, StartCol, NumOfRows, NumOfCols);
```