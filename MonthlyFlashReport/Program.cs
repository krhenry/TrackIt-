using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;


namespace MonthlyFlashReport
{
    class Program
    {
        static void Main(string[] args)
        {
            var prevMonth = DateTime.Now.AddMonths(-1).Month;
            var year = DateTime.Now.Year;
            DateTime today = DateTime.Today;
            DateTime EndDate = new DateTime(today.Year, today.AddMonths(-1).Month, DateTime.DaysInMonth(today.Year, today.AddMonths(-1).Month));

            if (prevMonth == 12)
            {
                year = year - 1;
                EndDate = new DateTime(today.AddYears(-1).Year, today.AddMonths(-1).Month, DateTime.DaysInMonth(today.AddYears(-1).Year, today.AddMonths(-1).Month));
            }

            string StartDate = prevMonth + "/1/" + year;
            
            try
            {
                Console.WriteLine("Console Application for TrackIt Flash Report.\nDefault Start and End dates will be set to last month: " + DateTime.Now.AddMonths(-1).ToString("MMMM") + "\n");
                Console.WriteLine("Default Start Date is: " + StartDate);
                Console.WriteLine("Default End Date is: " + EndDate.ToString("d") + "\n");

                UserInput(StartDate, EndDate);
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.ReadLine();
            }
        }

        static void UserInput(string StartDate, DateTime EndDate)
        {
            string validDate;
            try
            {
                Console.WriteLine("Hit Enter to Run Report with default dates or enter start date (MM/dd/yyyy)");
                var start = Console.ReadLine();
                if (start != "")
                {
                    validDate = DateValidation(start);

                    if (!string.IsNullOrEmpty(validDate))
                    {
                        StartDate = validDate.ToString();
                    }

                    Console.WriteLine("Enter end date");
                    var end = Console.ReadLine();
                    validDate = DateValidation(end);
                    if (!string.IsNullOrEmpty(validDate))
                    {
                        EndDate = Convert.ToDateTime(validDate);
                    }
                }

                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("Running TrackIt Report from " + StartDate + " to " + EndDate.ToString("d"));
                ExportToCSV(StartDate, EndDate.ToString("d"));
                Console.ReadLine();
            }

            catch(Exception ex)
            {
                Console.WriteLine(ex);
                Console.ReadLine();
            }
        }

        static string DateValidation(string date)
        {
            var dateFormats = new[] { "MM.dd.yyyy", "MM-dd-yyyy", "MM/dd/yyyy", "M.d.yyyy", "M-d-yyyy", "M/d/yyyy", "M.dd.yyyy", "M-dd-yyyy", "M/dd/yyyy", "MM.d.yyyy", "MM-d-yyyy", "MM/d/yyyy" };
            bool validate = true;

            DateTime scheduleDate;
            if (DateTime.TryParseExact(date, dateFormats, DateTimeFormatInfo.InvariantInfo, DateTimeStyles.None, out scheduleDate))
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Valid date");
                Console.ForegroundColor = ConsoleColor.Gray;
                validate = false;
                return date;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Invalid date: \"{0}\"", date);
                Console.ForegroundColor = ConsoleColor.Gray;
                Console.WriteLine("Enter valid date");
                string dt = Console.ReadLine();
                DateValidation(dt);
                return dt;
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
                string FileFullPath = DestinationFolder + "\\" + FileNamePart + " " + StartDate + " to " + EndDate + FileExtension;

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
