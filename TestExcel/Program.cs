using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BAExcel;

namespace TestExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            BookWrapper wrapper = ExcelUtil.CreateExcel();
            SheetWrapper swrapper = wrapper.CreateSheet("Test");
            CellPos pos = CellPos.CreateCellPos(0, 0);
            swrapper.SetText(pos, "Test");

            Console.WriteLine("Cell: " + swrapper.GetText(pos));
            Console.WriteLine("Cell Empty: " + swrapper.GetText(CellPos.CreateCellPos(1, 0)));

            sbyte[] t = wrapper.GetBytes();
            wrapper.Save("c:\\temp\\test.xlsx");


            BookWrapper w2 = ExcelUtil.LoadExcel(t);
            SheetWrapper s2 = w2["Test"];
            Console.WriteLine(s2.GetText(CellPos.CreateCellPos(0, 0)));

        }
    }
}
