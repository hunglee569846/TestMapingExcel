using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using TestMapingExcel;

namespace TestExcel2
{
    class Program
    {
       
        static void Main(string[] args)
        {
            List<StudentTest> studentlist = new List<StudentTest>();
           
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //mở file excel
            using (var pakage = new ExcelPackage(new FileInfo("ImportSv.xlsx")))
            {
                StudentTest st = new StudentTest();
                //chon sheets dau tien
                ExcelWorksheet worsheet = pakage.Workbook.Worksheets[0];
                for (int i = worsheet.Dimension.Start.Row; i <= 1 /*worsheet.Dimension.End.Row*/; i++)
                {
                    // StudentTest st = new StudentTest();
                    int j = 1;
                    st.Email = worsheet.Cells[i, j++].Value.ToString();
                    st.hoten = worsheet.Cells[i, j++].Value.ToString();
                    st.masv = worsheet.Cells[i, j++].Value.ToString();
                    st.lop = worsheet.Cells[i, j++].Value.ToString();
                }
                //check header cua file, dung voi cac truong csdl moi cho Maping
                if (st.Email == "Email" && st.hoten == "ho ten" && st.masv == "masv" && st.lop == "lop")
                {
                    Console.WriteLine("file dung");
                    for (int i = worsheet.Dimension.Start.Row+1; i <= worsheet.Dimension.End.Row; i++)
                    {
                        //StudentTest st1 = new StudentTest();
                        int j = 1;
                        string Email = worsheet.Cells[i, j++].Value.ToString();
                        string hoten = worsheet.Cells[i, j++].Value.ToString();
                        string masv = worsheet.Cells[i, j++].Value.ToString();
                        string lop = worsheet.Cells[i, j++].Value.ToString();
                        StudentTest student = new StudentTest()
                        {
                            Email = Email,
                            hoten = hoten,
                            masv = masv,
                            lop = lop,
                        };
                        studentlist.Add(student);
                       
                    }
                    foreach (var item in studentlist)
                    {
                        Console.WriteLine(item.Email);
                        Console.WriteLine(item.hoten);
                        Console.WriteLine(item.masv);
                        Console.WriteLine(item.lop);
                    }
                }
                else Console.WriteLine("file loi");

            }

        }
    }
}
