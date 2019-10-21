using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace ImportPassword
{

    public class ImportPassword
    {
        private static string userHMF = Environment.UserName;
        private string default_path = "C:\\Users\\" + userHMF + "\\ImportFiles\\ImportFile.xlsx";
        private int row_count = 0;
        private int col_count = 0;
        private int first_row = 1;
        private int second_row = 2;
        public System.Data.DataTable ImportDataTable = new System.Data.DataTable("TEST");


        public ImportPassword()
        {

            Application xlApp;
            Workbook xlWorkBook;
            Worksheet xlWorkSheet;
            Range range;

            string excelCellContents;
            string excelColumns;


            xlApp = new Application();
            xlWorkBook = xlApp.Workbooks.Open(default_path, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            row_count = range.Rows.Count;
            col_count = range.Columns.Count;

            string[,] insertArray = new string[row_count, col_count];


            //insert columns
            for (int y = 0; y <= col_count; y++)
            {
                excelColumns = (string)(range.Cells[first_row, y + first_row] as Range).Value2;
                ImportDataTable.Columns.Add(excelColumns, typeof(string));

            }

            //insert rows            
            for (int i = 0; i < row_count - first_row; i++)
            {
                DataRow drNew = ImportDataTable.NewRow(); // Has to be dynamically allocated
                for (int y = 0; y < col_count; y++)
                {
                    excelCellContents = (string)(range.Cells[i + second_row, y + first_row] as Range).Value2;
                    insertArray[i, y] = excelCellContents;

                    drNew[y] = excelCellContents;
                }
                ImportDataTable.Rows.Add(drNew);

            }

            //release from memory
            xlWorkBook.Close(true, null, null);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlWorkSheet);

        }

        public string getUser(string source)
        {

            string DataTableSource = null;
            string UserReturnValue = null;

            for (int i = 0; i <= row_count; i++)
            {
                try
                {
                    DataTableSource = (string)ImportDataTable.Rows[i]["Source"];
                    UserReturnValue = (string)ImportDataTable.Rows[i]["Username"];

                    i = (DataTableSource.ToUpper() == source.ToUpper()) ? row_count : i; // once value found, end loop
                }
                catch (System.IndexOutOfRangeException e)
                {
                    Console.WriteLine("Username and/or password not found. Please check spelling or check excel file in location below:");
                    Console.WriteLine(default_path);
                    Console.WriteLine("Press any key to exit....");
                    Console.ReadLine();
                    System.Environment.Exit(1);
                }
            }

            return UserReturnValue;

        }

        public string getPassword(string source)
        {

            string DataTableSource = null;
            string PWReturnValue = null;

            for (int i = 0; i <= row_count; i++)
            {
                try
                {
                    DataTableSource = (string)ImportDataTable.Rows[i]["Source"];
                    PWReturnValue = (string)ImportDataTable.Rows[i]["Password"];

                    i = (DataTableSource.ToUpper() == source.ToUpper()) ? row_count : i; // once value found, end loop
                }
                catch (System.IndexOutOfRangeException e)
                {
                    Console.WriteLine("Username and/or password not found. Please check spelling or check excel file in location below:");
                    Console.WriteLine(default_path);
                    Console.WriteLine("Press any key to exit....");
                    Console.ReadLine();
                    System.Environment.Exit(1);
                }

            }

            return PWReturnValue;

        }


        public int RowCount()
        {
            return row_count;
        }

        public int ColumnCount()
        {
            return col_count;
        }

        public static List<string> ReadInCSV(string absolutePath)
        {
            List<string> result = new List<string>();
            string value;
            using (TextReader fileReader = File.OpenText(absolutePath))
            {
                var csv = new CsvHelper.CsvReader(fileReader);
                csv.Configuration.HasHeaderRecord = false;
                while (csv.Read())
                {
                    for (int i = 0; csv.TryGetField<string>(i, out value); i++)
                    {
                        result.Add(value);
                    }
                }
            }
            return result;
        }
    }
}
