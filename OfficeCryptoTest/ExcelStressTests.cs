using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using System.Diagnostics;
using OfficeOpenXml;
using System.IO;

namespace OfficeOpenXmlCrypto.Test
{
    [TestFixture]
    public class ExcelStressTests
    {

        [Test]
        public void Write__10x_10()
        {
            Write(10, 10);
        }
        [Test]
        public void Write_100x_10()
        {
            Write(100, 10);
        }
        [Test]
        public void Write_500x_10()
        {
            Write(500, 10);
        }

        [Test]
        public void Write5000x_10()
        {
            Write(5000, 10);
        }

        public void Write(int rows, int cols)
        {
            String file = "testWs.xlsx";
            if (File.Exists(file)) { File.Delete(file); }

            TimeSpan start;

            // Write
            Console.WriteLine();
            Console.WriteLine("Test: " + rows + " x " + cols);
            Console.WriteLine("=======");
            Console.WriteLine("Writing");
            start = Process.GetCurrentProcess().TotalProcessorTime;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(file)))
            {
                int div = Math.Max(1, rows / 20);
                ExcelWorksheet ws = package.Workbook.Worksheets.Add("Stress", rows, cols);

                TimeRestart("Create", ref start);
                for (int row = 0; row < rows; row++)
                {
                    for (int col = 0; col < cols; col++)
                    {
                        ws.Cell(row + 1, col + 1).Value = GetVal(row, col);
                    }
                    if (row % div == 0) { Console.Write("*"); }
                }
                Console.WriteLine("done");

                TimeRestart("Write", ref start);

                package.Save();
            }
            TimeRestart("Save", ref start);


            // Read
            Console.WriteLine("Reading");
            start = Process.GetCurrentProcess().TotalProcessorTime;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(file)))
            {
                int div = Math.Max(1, rows / 20);
                ExcelWorksheet ws = package.Workbook.Worksheets["Stress"];

                TimeRestart("Open", ref start);
                for (int row = 0; row < rows; row++)
                {
                    for (int col = 0; col < cols; col++)
                    {
                        ExcelCell cell = ws.Cell(row + 1, col + 1);
                        String val = cell.Value;
                        Assert.AreEqual(GetVal(row, col), cell.Value, "@" + (row + 1) + ", " + (col + 1));
                    }
                    if (row % div == 0) { Console.Write("*"); }
                }
                Console.WriteLine("done");

                TimeRestart("Read", ref start);
            }
            Console.WriteLine(" "); // skip
        }

        static String GetVal(int rowIndex, int colIndex)
        {
            return "Val " + (rowIndex * colIndex + 1);
        }

        static void TimeRestart(String title, ref TimeSpan start)
        {
            TimeSpan total = Process.GetCurrentProcess().TotalProcessorTime - start;
            Console.WriteLine(title + ":\t" + total.TotalMilliseconds + "\tms ");
            start = Process.GetCurrentProcess().TotalProcessorTime;
        }

        [Test]
        public void WriteAndModInMemory()
        {
            String file = "testWs.xlsx";
            if (File.Exists(file)) { File.Delete(file); }

            int rows = 10;
            int cols = 10;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(file)))
            {
                ExcelWorksheet ws = package.Workbook.Worksheets.Add("Stress");
                for (int row = 0; row < rows; row++)
                {
                    for (int col = 0; col < cols; col++)
                    {
                        int val = (row * col) + 1;
                        ws.Cell(row + 1, col + 1).Value = val.ToString();
                    }
                }

                for (int row = 0; row < rows; row++)
                {
                    for (int col = 0; col < cols; col++)
                    {
                        int val = -int.Parse(ws.Cell(row + 1, col + 1).Value);
                        ws.Cell(row + 1, col + 1).Value = val.ToString();
                    }
                }

                for (int row = 0; row < rows; row++)
                {
                    for (int col = 0; col < cols; col++)
                    {
                        int val = -((row * col) + 1);
                        Assert.AreEqual(val.ToString(), ws.Cell(row + 1, col + 1).Value);
                    }
                }
            } 
        }
    }
}
