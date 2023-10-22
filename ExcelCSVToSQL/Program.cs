using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCSVToSQL
{
    internal class Program
    {
        public static readonly string EXCEL_PATH = "..\\Files\\Excel";
        public static readonly string CSV_PATH = "..\\Files\\Csv";
        public static readonly string SQL_PATH = "..\\Files\\Sql";
        public static readonly char DELIMITOR = ';';


        static void Main(string[] args)
        {
            CheckForFoldersExists();


            if (Directory.EnumerateFileSystemEntries(EXCEL_PATH).Any())
            {
                ConvertToCSV();
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey(true);
            }
            if (Directory.EnumerateFileSystemEntries(CSV_PATH).Any())
            {
                var files = Directory.EnumerateFiles(CSV_PATH);

                int counter = 1;

                foreach ( var file in files )
                {
                    //reading data
                    List<string> script = new List<string>();
                    var csvData = File.ReadAllLines(file);
                    //Drop table if exists
                    string drop = string.Format(Strings.Strings.dropTable, Path.GetFileNameWithoutExtension(file));
                    script.Add(drop);
                    //Header
                    List<string> headers = new List<string>();
                    string[] headerRow = csvData[0].Split(DELIMITOR);
                    foreach (var header in headerRow)
                    {
                        headers.Add(string.Format(Strings.Strings.tableColumns, header));
                    }

                    string create = string.Format(Strings.Strings.createTable, 
                        Path.GetFileNameWithoutExtension(file),
                        string.Join(", ", headers));
                    script.Add(create);

                    List<string> csvDataInList = csvData.ToList();
                    csvDataInList.RemoveAt(0);
                    //Values
                    foreach (var line in csvDataInList)
                    {
                        string[] datastream = line.Split(DELIMITOR);
                        List<string> data = new List<string>();
                        headers.Clear();

                        foreach (var item in datastream)
                        {
                            data.Add(string.Format(Strings.Strings.cellItem, item));
                        }
                        foreach (var item in headerRow)
                        {
                            headers.Add(string.Format(Strings.Strings.cellValue, item));
                        }

                        string insert = string.Format(Strings.Strings.tableInsert, 
                            Path.GetFileNameWithoutExtension(file), 
                            string.Join(", ", headers));

                        string valuesForInsert = string.Format(Strings.Strings.insertColumns, string.Join(", ", data));

                        script.Add(insert);
                        script.Add(valuesForInsert);
                    }

                    File.WriteAllLines($"{SQL_PATH}\\{Path.GetFileNameWithoutExtension(file)}.sql", script.ToArray());

                    Console.WriteLine($"{Path.GetFileName(file)} {counter}/{files.Count()} ----> {Path.GetFileNameWithoutExtension(file)}.sql Created!");
                    counter++;
                }
            }
            else
            {
                Console.WriteLine($"Folders {Path.GetFileName(Path.GetDirectoryName(EXCEL_PATH))} and {Path.GetFileName(Path.GetDirectoryName(CSV_PATH))} are empty!");
                Console.WriteLine("Press any key to finish...");
                Console.ReadKey(true);
            }





        }

        private static void ConvertToCSV()
        {
            var files = Directory.EnumerateFiles(EXCEL_PATH);

            int counter = 1;

            foreach (var file in files)
            {
                var workbook = new XLWorkbook(file);
                IXLWorksheet worksheet = workbook.Worksheets.Worksheet(1);

                var lastCellAddress = worksheet.RangeUsed().LastCell().Address;
                File.WriteAllLines($"{CSV_PATH}\\{Path.GetFileNameWithoutExtension(file)}.csv", worksheet.Rows(1, lastCellAddress.RowNumber)
                    .Select(r => string.Join(";", r.Cells(1, lastCellAddress.ColumnNumber)
                            .Select(cell =>
                            {
                                var cellValue = cell.GetValue<string>();
                                return cellValue.Contains(";") ? $"\"{cellValue}\"" : cellValue;
                            }))));

                Console.WriteLine($"{Path.GetFileName(file)} {counter}/{files.Count()}");
                counter++;
            }
        }

        private static void CheckForFoldersExists()
        {
            if (!Directory.Exists(EXCEL_PATH))
                Directory.CreateDirectory(EXCEL_PATH);
            if (!Directory.Exists(CSV_PATH))
                Directory.CreateDirectory(CSV_PATH);
            if (!Directory.Exists(SQL_PATH))
                Directory.CreateDirectory(SQL_PATH);
            Console.WriteLine($"Put files you want to convert in folders {Path.GetFileName(Path.GetDirectoryName(EXCEL_PATH))} and {Path.GetFileName(Path.GetDirectoryName(CSV_PATH))}");
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey(true);
        }
    }
}
