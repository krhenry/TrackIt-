using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;

/// <summary>
/// Purpose: An existing Stored Procedure is being executed at the beginning of each month to be emailed to the client for the previous months records. 
///     Ex -> If today's date is Feb 1st, SP is executed and sent in an excel file to client from Jan 1st -> Jan 31st records.
///     Client would like not to have to ask us to request this each month and be able to run this console application with user inputs. This CA will set the default
///     dates to be last month's dates.
/// 
/// </summary>


namespace MonthlyFlashReport
{
    class Program
    {
        /// <summary>
        /// Main: 
        ///     * Sets Defaults Dates.
        ///     * Determines default date if the current month is January.
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            // Set Pevious Month
            var prevMonth = DateTime.Now.AddMonths(-1).Month;
            // Set Current Month
            var year = DateTime.Now.Year;
            // Set End Date. -> Previous Month. Last Day.
            string EndDate = prevMonth.ToString() + "/" + DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.AddMonths(-1).Month) + "/" + year;
            // Set Start Date. -> Previous Month. First Day.
            string StartDay = prevMonth + "/1/" + year;


            // Sets the date and month to Decemeber and previous year if the current month is January.
            if (DateTime.Now.Month == 1)
            {
                year = year - 1;

                EndDate = ("12/" + DateTime.DaysInMonth(DateTime.Now.AddYears(-1).Year, 1) + "/" + year).ToString();
                StartDay = "12/1/" + year;
            }

            string StartDate = prevMonth + "/1/" + year;
            
            try
            {
                Console.WriteLine("Console Application for TrackIt Flash Report.\nDefault Start and End dates will be set to last month: " + DateTime.Now.AddMonths(-1).ToString("MMMM") + "\n");
                Console.WriteLine("Default Start Date is: " + StartDate);
                Console.WriteLine("Default Start Date is: " + EndDate);

                UserInput(StartDate, EndDate);
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.ReadLine();
            }
        }

        /// <summary>
        /// StartDate and EndDate will go through validation to determine if the dates are valid.
        /// If end date comes before start date, dates will flip so that it is passed to the stored procedure correctly.
        /// </summary>
        /// <param name="StartDate"></param>
        /// <param name="EndDate"></param>
        static void UserInput(string StartDate, string EndDate)
        {
            bool validDate;
            try
            {
                Console.WriteLine("Hit Enter to Run Report with default dates or enter start date (MM/dd/yyyy)");
                var start = Console.ReadLine();
                if (start != "")
                {
                    validDate = DateValidation(start);

                    while (validDate == false)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Invalid date: \"{0}\"", start);
                        Console.ForegroundColor = ConsoleColor.Gray;
                        Console.WriteLine("Enter valid date");
                        start = Console.ReadLine();
                        validDate = DateValidation(start);
                    }
                      
                    StartDate = start.ToString();

                    Console.WriteLine("Enter end date");
                    var end = Console.ReadLine();
                    validDate = DateValidation(end);
                    while (validDate == false)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Invalid date: \"{0}\"", end);
                        Console.ForegroundColor = ConsoleColor.Gray;
                        Console.WriteLine("Enter valid date");
                        end = Console.ReadLine();
                        validDate = DateValidation(end);
                    }
                    EndDate = end.ToString();
                }

                Console.ForegroundColor = ConsoleColor.White;
                if (Convert.ToDateTime(StartDate) > Convert.ToDateTime(EndDate))
                {
                    Console.WriteLine("Running TrackIt Report from " + EndDate + " to " + StartDate);
                    ExportToCSV(EndDate, StartDate);
                } else
                {
                    Console.WriteLine("Running TrackIt Report from " + StartDate + " to " + EndDate);
                    ExportToCSV(StartDate, EndDate);
                }
                Console.ReadLine();
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.ReadLine();
            }
        }

        /// <summary>
        /// Date Validation will check the format the user has inputted. There are several ways to pass in a valid date based on
        /// the variable 'dateFormats'. 
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        static bool DateValidation(string date)
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
                return true;
            }
            else
            {  
                return false;
            }
        }

        /// <summary>
        /// Function will execute SP and be sent to a specific file directory path in an excel file.
        /// </summary>
        /// <param name="StartDate"></param>
        /// <param name="EndDate"></param>
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
                string query = "EXEC " + StoredProcedureName + " " + StartDate + "," + EndDate;
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
