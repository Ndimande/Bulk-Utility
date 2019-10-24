

using Microsoft.Office.Core;
using RestSharp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;       //microsoft Excel 14 object in references-> COM tab

namespace RESTAPITest1
{
    public class Read_From_Excel
    {

        private dynamic NewFile;

        public string[] getExcelFile(string excel_Location)
        {


            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(excel_Location);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            string[] names = new string[rowCount];

            for (int i = 2; i < rowCount; i++)

            {
                for (int j = 2; j <= 2; j++)
                {
                    //new line

                    if (j == 2)
                        //  Console.Write("\r\n");



                        //write the value to the console
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {

                            NewFile = xlRange.Cells[i, j].Value2.ToString();
                            names[i] = NewFile;
                            // Console.WriteLine(NewFile + "\t");

                        }


                }

            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            return names;
        }

        public string[] BWS(string excel_Location)
        {


            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(excel_Location);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            string[] bws = new string[rowCount];

            for (int i = 2; i < rowCount; i++)

            {
                for (int j = 2; j <= 2; j++)
                {
                    //new line

                    if (j == 2)
                        //  Console.Write("\r\n");



                        //write the value to the console
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {

                            NewFile = xlRange.Cells[i, j].Value2.ToString();
                            //NewFile.Substring(Length - 5);
                            bws[i] = NewFile.Substring(0, 5);

                            //str.Substring(0, 5)
                            // Console.WriteLine(NewFile + "\t");

                        }


                }

            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            return bws;
        }

        public string[] nodeId(string path)
        {


            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            string[] nodes = new string[rowCount];


            for (int i = 2; i < rowCount; i++)

            {
                for (int j = 1; j <= 1; j++) //First Row
                {
                    //new line

                    if (j == 1)
                        //  Console.Write("\r\n");



                        //write the value to the console
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            NewFile = xlRange.Cells[i, j].Value2.ToString();
                            // nodes = NewFile[i];
                            //NewFile.Substring(Length - 5);



                            nodes[i] = NewFile;
                        }





                    // Console.WriteLine(NewFile.Trim() + "\t");


                }

            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            return nodes;
        }

        public string[] siteName(string path)
        {


            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            string[] nodes = new string[rowCount];


            for (int i = 2; i < rowCount; i++)

            {
                for (int j = 3; j <= 3; j++) //First Row
                {
                    //new line

                    if (j == 3)
                        //  Console.Write("\r\n");



                        //write the value to the console
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            NewFile = xlRange.Cells[i, j].Value2.ToString();
                            // nodes = NewFile[i];
                            //NewFile.Substring(Length - 5);


                            nodes[i] = NewFile;
                        }





                    // Console.WriteLine(NewFile.Trim() + "\t");


                }

            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            return nodes;
        }

        public string[] BWSName(string path)
        {


            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            string[] nodes = new string[rowCount];


            for (int i = 2; i < rowCount; i++)

            {
                for (int j = 2; j <= 2; j++) //First Row
                {
                    //new line

                    if (j == 2)
                        //  Console.Write("\r\n");



                        //write the value to the console
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            NewFile = xlRange.Cells[i, j].Value2.ToString();
                            nodes[i] = NewFile;
                            //NewFile.Substring(Length - 5);



                            // nodes[i] = NewFile;
                        }





                    // Console.WriteLine(NewFile.Trim() + "\t");


                }

            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            return nodes;
        }





        public string[] parentID(string path)
        {

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            string[] nodes = new string[rowCount];


            for (int i = 2; i < rowCount; i++)

            {
                for (int j = 1; j <= 1; j++) //First Row
                {
                    //new line

                    if (j == 1)
                        //  Console.Write("\r\n");



                        //write the value to the console
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            NewFile = xlRange.Cells[i, j].Value2.ToString();
                            // nodes = NewFile[i];
                            //NewFile.Substring(Length - 5);


                            nodes[i] = NewFile;
                        }





                    // Console.WriteLine(NewFile.Trim() + "\t");


                }

            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            return nodes;
        }

        public string[] recordType(string path)


        {

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            string[] nodes = new string[rowCount];


            for (int i = 2; i < rowCount; i++)

            {
                for (int j = 2; j <= 2; j++) //First Row
                {
                    //new line

                    if (j == 2)
                        //  Console.Write("\r\n");



                        //write the value to the console
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            NewFile = xlRange.Cells[i, j].Value2.ToString();
                            // nodes = NewFile[i];
                            //NewFile.Substring(Length - 5);


                            nodes[i] = NewFile;
                        }





                    // Console.WriteLine(NewFile.Trim() + "\t");


                }

            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            return nodes;
        }

        public string[] subFolder(string path)
        {

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            string[] nodes = new string[rowCount];


            for (int i = 2; i < rowCount; i++)

            {
                for (int j = 3; j <= 3; j++) //First Row
                {
                    //new line

                    if (j == 3)
                        //  Console.Write("\r\n");



                        //write the value to the console
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            NewFile = xlRange.Cells[i, j].Value2.ToString();
                            // nodes = NewFile[i];
                            //NewFile.Substring(Length - 5);


                            nodes[i] = NewFile;
                        }





                    // Console.WriteLine(NewFile.Trim() + "\t");


                }

            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            return nodes;
        }

        public string[] filePath(string path)
        {

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            string[] nodes = new string[rowCount];


            for (int i = 2; i < rowCount; i++)

            {
                for (int j = 4; j <= 4; j++) //First Row
                {
                    //new line

                    if (j == 4)
                        //  Console.Write("\r\n");



                        //write the value to the console
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            NewFile = xlRange.Cells[i, j].Value2.ToString();
                            // nodes = NewFile[i];
                            //NewFile.Substring(Length - 5);


                            nodes[i] = NewFile;
                        }





                    // Console.WriteLine(NewFile.Trim() + "\t");


                }

            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            return nodes;
        }


        public string[] categoryID(string path)
        {

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            string[] nodes = new string[rowCount];


            for (int i = 2; i < rowCount; i++)

            {
                for (int j = 7; j <= 7; j++) //First Row
                {
                    //new line

                    if (j == 7)
                        //  Console.Write("\r\n");



                        //write the value to the console
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            NewFile = xlRange.Cells[i, j].Value2.ToString();
                            // nodes = NewFile[i];
                            //NewFile.Substring(Length - 5);


                            nodes[i] = NewFile;
                        }





                    // Console.WriteLine(NewFile.Trim() + "\t");


                }

            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            return nodes;
        }



        public string[] parentIDForFolderCreation(string path)
        {

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            string[] nodes = new string[rowCount];


            for (int i = 2; i < rowCount; i++)

            {
                for (int j = 4; j <= 4; j++) //First Row
                {
                    //new line

                    if (j == 4)
                        //  Console.Write("\r\n");



                        //write the value to the console
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            NewFile = xlRange.Cells[i, j].Value2.ToString();
                            // nodes = NewFile[i];
                            //NewFile.Substring(Length - 5);


                            nodes[i] = NewFile;
                        }





                    // Console.WriteLine(NewFile.Trim() + "\t");


                }

            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            return nodes;
        }

        public string[] FolderName(string path)
        {

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            string[] nodes = new string[rowCount];


            for (int i = 2; i < rowCount; i++)

            {
                for (int j = 3; j <= 3; j++) //First Row
                {
                    //new line

                    if (j == 3)
                        //  Console.Write("\r\n");



                        //write the value to the console
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            NewFile = xlRange.Cells[i, j].Value2.ToString();
                            // nodes = NewFile[i];
                            //NewFile.Substring(Length - 5);


                            nodes[i] = NewFile;
                        }





                    // Console.WriteLine(NewFile.Trim() + "\t");


                }

            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            return nodes;
        }


        public string [] FolderValues(string path)
        {

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            string[] nodes = new string[rowCount];


            for (int i = 2; i < rowCount; i++)

            {
                for (int j = 2; j <= 2; j++) //First Row
                {
                    //new line

                    if (j == 2)
                        //  Console.Write("\r\n");



                        //write the value to the console
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            NewFile = xlRange.Cells[i, j].Value2.ToString();
                            // nodes = NewFile[i];
                            //NewFile.Substring(Length - 5);


                            nodes[i] = NewFile;
                        }





                    // Console.WriteLine(NewFile.Trim() + "\t");


                }

            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            return nodes;
        }
    }

    

}
