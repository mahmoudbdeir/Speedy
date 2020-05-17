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

#### Testing
Two Test Projects were made for testing:
1. <b>GetMethodsUnitTester:</b> That's mainly concerned about testing the get methods. It contains one class:
    <ol>
        <li><i>UnitTesting:</i> for testing all the possible test cases</li>
    </ol>
     

2. <b>SetMethodsUnitTester:</b> That's mainly concerned about testing the set methods. It contains two classes:
    <ol>
        <li><i>AlphaValuesUnitTesting:</i> that tests the set methods with alpha values, like strings</li>
        <li><i>NumericalValuesUnitTesting:</i> that tests the set methods with numerical values, like integers</li>
    </ol>
        
