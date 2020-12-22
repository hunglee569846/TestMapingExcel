using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace TestMapingExcel
{
    public class TestReadExcel : ITestfile
    {
        public bool ReadExcel(string file)
        {
            List<StudentTest> studentlist = new List<StudentTest>();
            StudentTest st = new StudentTest();

            var pakage = new ExcelPackage(new FileInfo(file));
            ExcelWorksheet worsheet = pakage.Workbook.Worksheets[1];
            for (int i = worsheet.Dimension.Rows ; i <= 1/*worsheet.Dimension.End.Row*/; i++)
            {
               // StudentTest st = new StudentTest();
                int j = 1;
                st.Email = worsheet.Cells[i, j++].Value.ToString();
                st.hoten = worsheet.Cells[i, j++].Value.ToString();
                st.masv = worsheet.Cells[i, j++].Value.ToString();
                st.lop = worsheet.Cells[i, j++].Value.ToString();
            }
            if (st.Email == "Email" && st.hoten == "ho ten" && st.Email == "masv" && st.Email == "lop")
            {
                return true;
            }
            else return false;
        }
        
    }
}
