using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyFlashReport
{
    class Program
    {
        static void Main(string[] args)
        {
            string datetime = DateTime.Now.ToString("yyyyMMddHHmmss");
            var LineConnection = ConfigurationManager.AppSettings["LineConnection"];
            var DBSchema = ConfigurationManager.AppSettings["DBSchema"];
            var prevMonth = DateTime.Now.AddMonths(-1).Month;
            var year = DateTime.Now.Year;

            if (prevMonth == 12)
            {
                year = year - 1;
            }

            string StartDate = prevMonth + "/1/" + year;
            DateTime today = DateTime.Today;
            DateTime EndDate = new DateTime(today.Year, today.AddMonths(-1).Month, DateTime.DaysInMonth(today.Year, today.AddMonths(-1).Month));

            try
            {
                //Declare Variables and provide values
                //string FileNamePart = "TrackIt Flash Report";//Datetime will be added to it
                //string DestinationFolder = @"C:\Clients\_NONGIT\OLV\";
                //string StoredProcedureName = "uspSelectAllFromWoundType";//Provide SP name,you Can provide with Parameter if you like
                //string FileDelimiter = ","; //You can provide comma or pipe or whatever you like
                //string FileExtension = ".csv"; //Provide the extension you like such as .txt or .csv
                                               //\nDefault Month is set to: " + DateTime.Now.AddMonths(-1).ToString("MMMM"));
                bool validDate;

                Console.WriteLine("Console Application for TrackIt Flash Report.\nDefault Start and End dates will be set to last month: " + DateTime.Now.AddMonths(-1).ToString("MMMM") + "\n");
                Console.WriteLine("Default Start Date is: " + StartDate);
                Console.WriteLine("Default End Date is: " + EndDate.ToString("d") + "\n");
                Console.WriteLine("Hit Enter to Run Report with default dates or enter start date (ex. 01/01/2016)");
                var start = Console.ReadLine();
                if (start != "")
                {
                    validDate = DateValidation(start);

                    if (validDate == true)
                    {
                        StartDate = start.ToString();
                    }

                    Console.WriteLine("Enter end date");
                    var end = Console.ReadLine();
                    validDate = DateValidation(end);
                    if (validDate == true)
                    {
                        EndDate = Convert.ToDateTime(end);
                    }
                }

                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("Running TrackIt Report from " + StartDate + " to " + EndDate.ToString("d"));
                ExportToCSV(StartDate, EndDate.ToString("d"));
                Console.ReadLine();

                //Create Connection to SQL Server in which you like to load files
                //SqlConnection SQLConnection = new SqlConnection();
                //SQLConnection.ConnectionString = LineConnection;

                ////Execute Stored Procedure and save results in data table
                //string query = "EXEC " + StoredProcedureName + " " + "'Surgical Wound'" + "," + "'Import'";
                //SqlCommand cmd = new SqlCommand(query, SQLConnection);
                //SQLConnection.Open();
                //DataTable d_table = new DataTable();
                //d_table.Load(cmd.ExecuteReader());
                //SQLConnection.Close();

                ////Prepare the file path 
                //StartDate = StartDate.Replace("/", "_");
                ////defaultEndDate = defaultEndDate.ToString("d").Replace("/", "_");
                //string FileFullPath = DestinationFolder + "\\" + FileNamePart + "_" + StartDate + " to " + EndDate.ToString("d").Replace("/","_") + FileExtension;

                //StreamWriter sw = null;
                //sw = new StreamWriter(FileFullPath, false);

                //// Write the Header Row to File
                //int ColumnCount = d_table.Columns.Count;
                //for (int ic = 0; ic < ColumnCount; ic++)
                //{
                //    sw.Write(d_table.Columns[ic]);
                //    if (ic < ColumnCount - 1)
                //    {
                //        sw.Write(FileDelimiter);
                //    }
                //}
                //sw.Write(sw.NewLine);

                //// Write All Rows to the File
                //foreach (DataRow dr in d_table.Rows)
                //{
                //    for (int ir = 0; ir < ColumnCount; ir++)
                //    {
                //        if (!Convert.IsDBNull(dr[ir]))
                //        {
                //            sw.Write(dr[ir].ToString());
                //        }
                //        if (ir < ColumnCount - 1)
                //        {
                //            sw.Write(FileDelimiter);
                //        }
                //    }
                //    sw.Write(sw.NewLine);

                //}

                //sw.Close();

            }

            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.ReadLine();
            }
        }

        static bool DateValidation(string date)
        {
            var dateFormats = new[] { "dd.MM.yyyy", "dd-MM-yyyy", "dd/MM/yyyy", "d.M.yyyy", "d-M-yyyy", "d/M/yyyy" };
            bool validate = true;

            DateTime scheduleDate;
            if (DateTime.TryParseExact(date, dateFormats, DateTimeFormatInfo.InvariantInfo, DateTimeStyles.None, out scheduleDate))
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Valid date");
                Console.ForegroundColor = ConsoleColor.Gray;
                validate = false;
                return true;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Invalid date: \"{0}\"", date);
                Console.ForegroundColor = ConsoleColor.Gray;
                Console.WriteLine("Enter valid date: MM/dd/yyyy");
                string b = Console.ReadLine();
                DateValidation(b);
                return false;
            }
        }

        static void ExportToCSV(string StartDate, string EndDate)
        {
            try
            {
                //Declare Variables and provide values
                string FileNamePart = "Flash Report";//Datetime will be added to it
                string DestinationFolder = @"C:\Clients\_NONGIT\OLV\";
                string StoredProcedureName = "uspSelectAllFromWoundType";//Provide SP name,you Can provide with Parameter if you like
                string FileDelimiter = ","; //You can provide comma or pipe or whatever you like
                string FileExtension = ".csv"; //Provide the extension you like such as .txt or .csv
                var LineConnection = ConfigurationManager.AppSettings["LineConnection"];
                var DBSchema = ConfigurationManager.AppSettings["DBSchema"];

                //Create Connection to SQL Server in which you like to load files
                SqlConnection SQLConnection = new SqlConnection();
                SQLConnection.ConnectionString = LineConnection;

                //Execute Stored Procedure and save results in data table
                string query = "EXEC " + StoredProcedureName + " " + "'Surgical Wound'" + "," + "'Import'";
                SqlCommand cmd = new SqlCommand(query, SQLConnection);
                SQLConnection.Open();
                DataTable d_table = new DataTable();
                d_table.Load(cmd.ExecuteReader());
                SQLConnection.Close();

                //Prepare the file path 
                StartDate = StartDate.Replace("/", "_");
                EndDate = EndDate.Replace("/", "_");
                string FileFullPath = DestinationFolder + "\\" + FileNamePart + " " + StartDate + " to " + EndDate.Replace("/", "_") + FileExtension;

                StreamWriter sw = null;
                sw = new StreamWriter(FileFullPath, false);

                // Write the Header Row to File
                int ColumnCount = d_table.Columns.Count;
                for (int ic = 0; ic < ColumnCount; ic++)
                {
                    sw.Write(d_table.Columns[ic]);
                    if (ic < ColumnCount - 1)
                    {
                        sw.Write(FileDelimiter);
                    }
                }
                sw.Write(sw.NewLine);

                // Write All Rows to the File
                foreach (DataRow dr in d_table.Rows)
                {
                    for (int ir = 0; ir < ColumnCount; ir++)
                    {
                        if (!Convert.IsDBNull(dr[ir]))
                        {
                            sw.Write(dr[ir].ToString());
                        }
                        if (ir < ColumnCount - 1)
                        {
                            sw.Write(FileDelimiter);
                        }
                    }
                    sw.Write(sw.NewLine);

                }

                sw.Close();
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.ReadLine();
            }
        }
    }
}
