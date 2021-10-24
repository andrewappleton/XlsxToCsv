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
using NLog;
using System.Diagnostics;

namespace XlsxToCsv
{

    class XlsxToCsv
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        static void Main(string[] args)
        {
            if (args.Length != 1)
            {
                logger.Info("XlsxToCsv v.0.1");
                logger.Info("Usage: XlsxToCsv.exe filename.xlsx");
            }
            else
            {
                _ = new XlsxToCsv(args[0]);
                if (Debugger.IsAttached) //only pause exit if running from IDE.
                    System.Console.ReadLine();
            }
        }

        public XlsxToCsv(string filename)
        {
            logger.Info("Reading {0}...",filename);
            ReadExcelFile(filename);
        }

        private bool WriteCsvFile(DataSet dsExcelData, string inputFileName)
        {
            string sDelimiter = ConfigurationManager.AppSettings.Get("Delimiter");
            string sQualifier = (ConfigurationManager.AppSettings.Get("UseTextQualifier") == "0" ? "" : "\"");
            Boolean bOverwrite = (ConfigurationManager.AppSettings.Get("OverwriteCSV") == "1");
            string sOverwriteMessage = String.Format("Overwriting CSV: {0}", bOverwrite);
            logger.Info("Using Qualifier: {0}",(sQualifier != ""));
            logger.Info("Using Delimiter: {0}",sDelimiter);
            if (bOverwrite)
                logger.Warn(sOverwriteMessage);
            else
                logger.Info(sOverwriteMessage);
            logger.Info("Found dataset with {0} table(s) and {1} rows",dsExcelData.Tables.Count, dsExcelData.Tables[0].Rows.Count);
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
                            logger.Trace(strRow);
                            writer.WriteLine(strRow);
                        }

                    }
                    writer.Close();
                    fs.Close();
                }
                catch (IOException e)
                {
                    logger.Error("ERROR: " + e.Message);
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
                        logger.Info("Found {0} rows and {1} columns on sheet '{2}'",
                            reader.RowCount, reader.FieldCount, reader.Name);
                        logger.Info("Reading data from workbook...");
                        WriteCsvFile(reader.AsDataSet(),filename);
                    }
                }
            }
            catch (IOException e)
            {
                logger.Error("ERROR: {0}!",e.Message);
            }
            catch (Exception e)
            {
                logger.Error("Error occurred: {0}", e.Message);
            }
            return true;
        }

        private string SetOutputFileName(string inputFileName)
        {
            return Path.ChangeExtension(inputFileName, ".csv");
        }
    }
}
