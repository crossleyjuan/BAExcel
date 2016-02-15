Bizagi Excel Util
===

This contains the component library that can be used in Bizagi projects to load and process Excel files, using NPOI, this new implementation works better than automation or ODBC Datasources as the underlying dll from Apache is more powerful and requires less memory.

How to use it
----

To use it you just need to add the BAExcel.dll as part of the component library using the following:

Name: BAExcel
Namespace: BAExcel

and add the NPOI.dll as:
Name: NPOI
Namespace: NPOI

Sample Rule
---
Here's the hello world sample, this will create a new Excel with a Hello World text in the first row, first column and add it to an file attribute:

    var wrapper = BAExcel.ExcelUtil.CreateExcel();
    var swrapper = wrapper.CreateSheet("First");
    var pos = CellPos.CreateCellPos(0, 0);
    swrapper.SetText(pos, "Hello World!");
    var t = wrapper.GetBytes();
    var f = Me.newCollectionItem("Test.uploads");
    f.setXPath("data", t);
    f.setXPath("fileName", "e.xls");

Load an exiting excel file:

    var wrapper2 = BAExcel.ExcelUtil.LoadExcel(t);
    var s2 = w2["Test"];
    var text = s2.GetText(CellPos.CreateCellPos(0, 0)));
