// Author: Eduardo Santos
// Date: 2015-11-11

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenXmltoExcel;
using System.Configuration;

namespace SpreadsheetCreator
{
    public class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string directoryToFiles = ConfigurationManager.AppSettings["DIRECTORY_EXCEL_FILES"].ToString();
                string filename = "Hello.xlsx";
                string copy = "CopyOfHello.xlsx";

                string existingFile = directoryToFiles + filename;
                string newFile = directoryToFiles + copy;

                System.Console.WriteLine("Export Open XML to Excel 1.0");
                System.Console.WriteLine("Let's do this...\n");

                CopyXlsx xlsxFile = new CopyXlsx();
                xlsxFile.CopyAndSave(newFile, existingFile);
                
                System.Console.WriteLine("The task is complete. Hit ENTER...");
                System.Console.ReadKey();

            }
            catch (Exception e)
            {
            }
        }
    }
}
