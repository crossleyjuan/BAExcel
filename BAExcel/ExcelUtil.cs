using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using System.IO;
using NPOI.HPSF;

namespace BAExcel
{
    public class ExcelUtil
    {

        private static sbyte[] LoadFile(string fileName)
        {
            FileStream fs = new FileStream(fileName, FileMode.Open);

            byte[] buffer = new byte[1024];
            MemoryStream ms = new MemoryStream();
            int readed = 0;
            while ((readed = fs.Read(buffer, 0, 1024)) > 0)
            {
                ms.Write(buffer, 0, readed);
            }

            fs.Close();

            byte[] result = ms.ToArray();
            sbyte[] sresult = new sbyte[result.Length];
            Buffer.BlockCopy(result, 0, sresult, 0, result.Length);
            return sresult;
        }

        public static BookWrapper LoadExcel(string fileName)
        {
            sbyte[] fileContent = LoadFile(fileName);

            return LoadExcel(fileContent);
        }

        public static BookWrapper LoadExcel(sbyte[] sfile)
        {
            byte[] file = new byte[sfile.Length];
            Buffer.BlockCopy(sfile, 0, file, 0, sfile.Length);
            MemoryStream ms = new MemoryStream(file);

            HSSFWorkbook book = new HSSFWorkbook(ms);

            BookWrapper wrapper = new BookWrapper(book);

            return wrapper;
        }

        public static BookWrapper CreateExcel()
        {
            HSSFWorkbook book = new HSSFWorkbook();
            BookWrapper wrapper = new BookWrapper(book);
            ////create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            book.DocumentSummaryInformation = dsi;

            ////create a entry of SummaryInformation
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            book.SummaryInformation = si;


            return wrapper;
        }

    }
}
