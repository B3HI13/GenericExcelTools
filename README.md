# GenericExcelTools
Export generic data to excel &amp; generic data extractor from excel.

This project created with `.NET Core 3.1` and [ClosedXM](https://www.nuget.org/packages/ClosedXML) nuget package. It contains read and export data. 
The methods are generic and can be used with every class model that properties have `Display Annotation`.

In read, the method trys to find columns header text that are same with `Name` value of "Display Annotation" on properties of model.
In Export, the method trys to fill column header textes with "Name" value of "Display Annotation" of properties and then begins to fill cells of the excel file.
