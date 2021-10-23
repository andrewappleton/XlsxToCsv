using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Collections.Specialized;

namespace XlsxToCsv
{
    class XlsxToCsv
    {
        static void Main(string[] args)
        {
            if (args.Length != 1)
            {
                System.Console.WriteLine("XlsxToCsv v.0.1");
                System.Console.WriteLine("Usage: XlsxToCsv.exe filename.xlsx");
            }
            else
            {
                new XlsxToCsv(args[0]);
            }
        }

        public XlsxToCsv(string filename)
        {
            System.Console.WriteLine("Reading {0}...",filename);
            ReadExcelFile(filename);
        }

        private bool WriteCsvFile(DataSet dsExcelData, string inputFileName)
        {
            string sDelimiter = ConfigurationManager.AppSettings.Get("Delimiter");
            string sQualifier = (ConfigurationManager.AppSettings.Get("UseTextQualifier") == "0" ? "" : "\"");
            Boolean bOverwrite = (ConfigurationManager.AppSettings.Get("OverwriteCSV") == "1");
            System.Console.WriteLine("Using Qualifier: {0}",(sQualifier != ""));
            System.Console.WriteLine("Using Delimiter: {0}",sDelimiter);
            System.Console.WriteLine("Overwriting CSV: {0}",bOverwrite);
            System.Console.WriteLine("Found dataset with {0} table(s) and {1} rows",dsExcelData.Tables.Count, dsExcelData.Tables[0].Rows.Count);
            StringBuilder content = new StringBuilder();
            if (dsExcelData.Tables.Count >= 1)
            {
                try
                {
                    FileStream fs = new FileStream(SetOutputFileName(inputFileName), (bOverwrite ? FileMode.Create : FileMode.CreateNew), FileAccess.Write, FileShare.None);
                    StreamWriter writer = new StreamWriter(fs, Encoding.UTF8);
                    
                    DataTable table = dsExcelData.Tables[0];
                    if (table.Rows.Count > 0)
                    {
                        DataRow dr1 = (DataRow)table.Rows[0];
                        int intColumnCount = dr1.Table.Columns.Count;
                        int index = 1;

                        foreach (DataColumn item in dr1.Table.Columns)
                        {
                            content.Append(String.Format("{0}{1}{2}",sQualifier,item.ColumnName,sQualifier));
                            if (index < intColumnCount)
                                content.Append(sDelimiter);
                            else
                                content.Append("\r\n");
                            index++;
                        }

                        foreach (DataRow currentRow in table.Rows)
                        {
                            string strRow = string.Empty;
                            for (int y = 0; y <= intColumnCount - 1; y++)
                            {
                                strRow += String.Format("{0}{1}{2}", sQualifier, currentRow[y].ToString(), sQualifier);
                                if (y < intColumnCount - 1 && y >= 0)
                                    strRow += sDelimiter;
                            }
                            content.Append(strRow + "\r\n");
                            System.Console.WriteLine(strRow);
                            writer.WriteLine(strRow);
                        }

                    }
                    writer.Close();
                    fs.Close();
                }
                catch (IOException e)
                {
                    System.Console.WriteLine("ERROR: " + e.Message);
                }
            }
            return true;
        }

        private bool ReadExcelFile(string filename)
        {
            try
            {
                using (var stream = File.Open(@filename, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        System.Console.WriteLine("Found {0} rows and {1} columns on sheet '{2}'",
                            reader.RowCount, reader.FieldCount, reader.Name);
                        System.Console.WriteLine("Reading data from workbook...");
                        WriteCsvFile(reader.AsDataSet(),filename);
                    }
                }
            }
            catch (IOException e)
            {
                System.Console.WriteLine("ERROR: File not found!");
            }
            catch (Exception e)
            {
                System.Console.WriteLine("Error occurred: {0}", e.Message);
            }
            return true;
        }

        private string SetOutputFileName(string inputFileName)
        {
            return Path.ChangeExtension(inputFileName, ".csv");
        }
    }
}
