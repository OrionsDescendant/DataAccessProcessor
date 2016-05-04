using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;

using Excel = Microsoft.Office.Interop.Excel;

namespace DAProcessor
{
    class Processor
    {
        string ProcessPath = string.Empty;
        string SQLFileToProcess = string.Empty;
        string LogFilePrefix = DateTime.Now.Year.ToString() + "." + DateTime.Now.Month.ToString() + "." + DateTime.Now.Day.ToString() + "_" + DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString() + "_";
        string RunLogFile = string.Empty;
        string SqlFolder = string.Empty;
        string ColsolidatedExcelFileName = string.Empty;

        bool IsProduction = true;
        bool UseTransaction = false;
        bool ExecuteSQLFilesSeparately = false;
        bool GenerateExcelFilesFromSelectQueryFiles = false;
        bool ConsolidateExcelFilesIntoSingleFileOnSeperateTabs = false;

        int DefaultTimeOut = 120;
        int TimeOut = 30;

        string connectionString = string.Empty;
        List<string> SqlFilesToProcess = new List<string>();

        List<SqlFileToProcess> ParsedSqlFilesToProcess = new List<SqlFileToProcess>();
        List<DataSet> DataSetsForExcel = new List<DataSet>();

        /// <summary>
        /// Generic processor for executing scripts and generating Excel files. It will read a list of sql files named Step_xx* and then execute them according to the settings.
        ///     If it is executed using a transaction, if there are any errors, then the transaction is automatically rolledback
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            Processor p = new Processor();
            DateTime ProcessStartTime = DateTime.Now;
            bool OverrideSettings = false;

            // list of required parameters that is used only when the overrides are passed in.
            List<string> RequiredParameters = new List<string>();
            RequiredParameters.Add("IsProduction");
            RequiredParameters.Add("ProcessPath");
            RequiredParameters.Add("FileToProcess");
            RequiredParameters.Add("TimeOut");
            RequiredParameters.Add("GenerateExcelFilesFromSelectQueryFiles");
            RequiredParameters.Add("ConsolidateExcelFilesIntoSingleFileOnSeperateTabs");

            try
            {
                if (args.Length == 0)
                {
                    // path not supplied, use current path
                    p.ProcessPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                }
                else if (args.Length == 1)
                {
                    p.ProcessPath = args[0].ToString();

                    // 4/10/2015 remove last \ from path if exists in attempt to avoid the The path is not of a legal form. error
                    // an extra " was being placed at the end of the path automatically by .net code which resulted in "" at
                    // the end of the path. It also added an extra \ at the end as well resulting in the legal for error
                    // By removing the double "" at the end, it fixed the legal for error
                    p.ProcessPath = p.ProcessPath.Replace(@"""", "");

                    if (!Directory.Exists(p.ProcessPath))
                    {
                        throw new Exception("CRITICAL - Supplied PATH NOT FOUND. Be sure to encompass the path in double quotes. Path supplied: " + p.ProcessPath );
                    }
                }
                else if (args.Length != 1) /* change this according to the number of required parameters*/
                {
                    OverrideSettings = OverrideSettingsUsingCLIParameters(args, p, OverrideSettings, RequiredParameters);
                }

                if (p.ProcessPath.Length > 0)
                {
                    if (!Directory.Exists(p.ProcessPath))
                    {
                        throw new Exception("CRITICAL - Supplied PATH NOT FOUND. Be sure to encompass the path in double quotes. Path supplied: " + p.ProcessPath);
                    }
                }

                if (p.SQLFileToProcess.Length > 0)
                {
                    if (!File.Exists(p.SQLFileToProcess))
                    {
                        throw new Exception("CRITICAL - Supplied FileToProcess NOT FOUND. Be sure to encompass the path in double quotes. Path supplied: " + p.SQLFileToProcess);
                    }
                }


                if (p.ProcessPath.Length == 0 && p.SQLFileToProcess.Length == 0)
                {
                    throw new Exception("CRITICAL - Please supply a valid ProcessPath or supply a FileToProcess");
                }

                string FilePath ;

                if (p.ProcessPath.Length > 0)
                {
                    FilePath = p.ProcessPath;
                    p.SqlFolder = FilePath + "\\SQL";
                    p.RunLogFile = FilePath + "\\Logs\\" + p.LogFilePrefix + "RunLog.txt";
                }
                else
                {
                    p.RunLogFile = (Path.GetDirectoryName(p.SQLFileToProcess)) + "\\Logs\\" + p.LogFilePrefix + "RunLog.txt";
                }


                if (RequiredParameters.Count > 0 && OverrideSettings)
                {
                    p.WriteToLogAndConsole("CRITICAL - Please supply the below REQUIRED parameters: ");
                    foreach (string item in RequiredParameters)
                    {
                        p.WriteToLogAndConsole(" * " + item);
                    }

                    throw new Exception("CRITICAL - missing parameters. Check the console output and the log file for more information");
                }

                p.WriteToLogAndConsole("Process start ");

                // check to see if we need to get the settings from the app.config file or not
                if (!OverrideSettings)
                {
                    GetSettings(p);
                }
                else
                {
                    p.WriteToLogAndConsole("Overrides provided so skipping config file");
                }

                if (p.IsProduction)
                {
                    p.connectionString = File.ReadAllText(@"C:\DataSources\SQLProduction.txt");
                }
                else
                {
                    // default to BOLTEST for now
                    p.connectionString = File.ReadAllText(@"C:\DataSources\SQLTest.txt");
                }

                // only search for the sql directory when a file to process is not passed in
                if (p.SQLFileToProcess.Length == 0)
                {
                    p.WriteToLogAndConsole("Get SQL files");
                    p.WriteToLogAndConsole("Searching for SQL files in : " + p.SqlFolder);
                    if (!Directory.Exists(p.SqlFolder))
                    {
                        p.WriteToLogAndConsole("CRITICAL - SQL FOLDER NOT FOUND");
                        p.WriteToLogAndConsole("Aborting process");

                        throw new Exception("CRITICAL - SQL FOLDER NOT FOUND");
                    }
                    // pull files that start with the prefix of 'step'
                    p.SqlFilesToProcess = Directory.GetFiles(p.SqlFolder, "step*.sql").ToList();

                    // sort them so they will be processed in a specific order
                    p.SqlFilesToProcess.Sort();
                }
                else
                {
                    // go ahead and add the single file passed in so the rest of the logic will be applied normally
                    p.SqlFilesToProcess.Add(p.SQLFileToProcess);
                }

                p.WriteToLogAndConsole("Found " + p.SqlFilesToProcess.Count + " files to process");

                p.ProcessSQLFiles(ref p);

                if (p.ConsolidateExcelFilesIntoSingleFileOnSeperateTabs)
                    p.GenerateConsolodatedExcelFile();

                DateTime ProcessEndTime = DateTime.Now;
                TimeSpan ProcessTime = ProcessEndTime - ProcessStartTime;

                p.WriteToLogAndConsole("     Total Process Time = " + ProcessTime.Hours.ToString() + ":" + ProcessTime.Minutes.ToString() + ":" + ProcessTime.Seconds.ToString());

                p.WriteToLogAndConsole("Process complete ");
            }
            catch (Exception ex)
            {
                p.WriteToLogAndConsole("Exception occurred!");
                p.WriteToLogAndConsole("Exception.Message = " + ex.Message);
                p.WriteToLogAndConsole("");
                p.WriteToLogAndConsole("Exception.GetBaseException() = " + ex.GetBaseException().ToString());

                if (System.Diagnostics.Debugger.IsAttached)
                {
                    Console.ReadLine();
                }
                else
                {
                    // bubble up the exception to the caller
                    throw;
                }
            }
        }

        /// <summary>
        /// Override all of the required settings with values provided via the command line.
        /// </summary>
        /// <param name="args"></param>
        /// <param name="p"></param>
        /// <param name="OverrideSettings"></param>
        /// <param name="RequiredParameters"></param>
        /// <returns></returns>
        private static bool OverrideSettingsUsingCLIParameters(string[] args, Processor p, bool OverrideSettings, List<string> RequiredParameters)
        {
            /* This logic block is for overriding certain parameters\settings\variables
               This is so this program can be used in a manner in which makes it more dynamic such as non developers
                    whom will do some reporting which would need datasets saved into Excel files. */

            for (int i = 0; i < args.Length; i++)
            {
                if (args[i].ToLower().ToString() == "isproduction")
                {
                    if (args[i + 1].ToLower().ToString() == "1" || args[i + 1].ToLower().ToString() == "true")
                    {
                        p.IsProduction = true;
                        RequiredParameters.Remove("IsProduction");
                    }
                    else if (args[i + 1].ToLower().ToString() == "0" || args[i + 1].ToLower().ToString() == "false")
                    {
                        p.IsProduction = false;
                        RequiredParameters.Remove("IsProduction");
                    }
                    else
                    {
                        Console.WriteLine("CRITICAL - Invalid value supplied for parameter isproduction. Provided value: " + args[i + 1].ToString());
                    }

                    // increment to skip reading the value of the override
                    i++;
                }
                else if (args[i].ToLower().ToString() == "processpath")
                {
                    p.ProcessPath = args[i + 1].ToString();

                    // increment to skip reading the value of the override
                    i++;

                    RequiredParameters.Remove("ProcessPath");
                }
                else if (args[i].ToLower().ToString() == "filetoprocess")
                {
                    p.SQLFileToProcess = args[i + 1].ToString();

                    // increment to skip reading the value of the override
                    i++;

                    RequiredParameters.Remove("FileToProcess");
                }
                else if (args[i].ToLower().ToString() == "timeout")
                {
                    int tempTimeOut = 0;

                    try
                    {
                        tempTimeOut = int.Parse(args[i + 1].ToString());

                        p.TimeOut = tempTimeOut;

                        // increment to skip reading the value of the override
                        i++;
                    }
                    catch (Exception)
                    {
                        throw new Exception("TimeOut - The supplied value is not a valid value. Value provided: " + args[i].ToString() + "Please supply a TimeOut that is between 10 - 1200. (Note that the TimeOut value is in secods)");
                    }

                    RequiredParameters.Remove("TimeOut");
                }
                else if (args[i].ToLower().ToString() == "generateexcelfilesfromselectqueryfiles")
                {
                    if (args[i + 1].ToLower().ToString() == "1" || args[i + 1].ToLower().ToString() == "true")
                    {
                        p.GenerateExcelFilesFromSelectQueryFiles = true;
                        RequiredParameters.Remove("GenerateExcelFilesFromSelectQueryFiles");
                    }
                    else if (args[i + 1].ToLower().ToString() == "0" || args[i + 1].ToLower().ToString() == "false")
                    {
                        p.GenerateExcelFilesFromSelectQueryFiles = false;
                        RequiredParameters.Remove("GenerateExcelFilesFromSelectQueryFiles");
                    }

                    i++;
                }
                else if (args[i].ToLower().ToString() == "consolidateexcelfilesintosinglefileonseperatetabs")
                {
                    if (args[i + 1].ToLower().ToString() == "1" || args[i + 1].ToLower().ToString() == "true")
                    {
                        p.ConsolidateExcelFilesIntoSingleFileOnSeperateTabs = true;
                        RequiredParameters.Remove("ConsolidateExcelFilesIntoSingleFileOnSeperateTabs");
                    }
                    else if (args[i + 1].ToLower().ToString() == "0" || args[i + 1].ToLower().ToString() == "false")
                    {
                        p.ConsolidateExcelFilesIntoSingleFileOnSeperateTabs = false;
                        RequiredParameters.Remove("ConsolidateExcelFilesIntoSingleFileOnSeperateTabs");
                    }

                    i++;
                }
            }

            // set this variable to true to skip over processing the app.config file
            OverrideSettings = true;

            // only 1 can be supplied so if 1 is supplied then remove it from the list of required parameters
            if (p.SQLFileToProcess.Length > 0 || p.ProcessPath.Length > 0)
            {
                RequiredParameters.Remove("filetoprocess");
                RequiredParameters.Remove("processpath");

                if (p.SQLFileToProcess.Length > 0)
                    p.ProcessPath = string.Empty;
            }
            return OverrideSettings;
        }

        /// <summary>
        /// Read the app.settings file
        /// </summary>
        /// <param name="p"></param>
        private static void GetSettings(Processor p)
        {
            // initialize the settings 
            p.WriteToLogAndConsole("Reading settings...");
            try
            {
                p.IsProduction = bool.Parse(ConfigurationManager.AppSettings["IsProduction"].ToString());
                p.WriteToLogAndConsole("    Using setting IsProduction = " + p.IsProduction.ToString());
            }
            catch (Exception)
            {
                p.WriteToLogAndConsole("    Invalid IsProduction setting. Defaulting to IsProduction = true");
                p.IsProduction = true;
            }

            try
            {
                p.TimeOut = int.Parse(ConfigurationManager.AppSettings["TimeOut"].ToString());
                p.WriteToLogAndConsole("    Using setting TimeOut = " + p.TimeOut.ToString());
            }
            catch (Exception)
            {
                p.WriteToLogAndConsole("    Invalid TimeOut setting. Defaulting to TimeOut = " + p.DefaultTimeOut.ToString());
                p.TimeOut = p.DefaultTimeOut;
            }

            try
            {
                p.UseTransaction = bool.Parse(ConfigurationManager.AppSettings["UseTransaction"].ToString());
                p.WriteToLogAndConsole("    Using setting UseTransaction = " + p.UseTransaction.ToString());
            }
            catch (Exception)
            {
                p.WriteToLogAndConsole("    Invalid UseTransaction setting. Defaulting to UseTransaction = false ");
                p.UseTransaction = false;
            }

            try
            {
                p.ExecuteSQLFilesSeparately = bool.Parse(ConfigurationManager.AppSettings["ExecuteSQLFilesSeparately"].ToString());
                p.WriteToLogAndConsole("    Using setting ExecuteSQLFilesSeparately = " + p.ExecuteSQLFilesSeparately.ToString());
            }
            catch (Exception)
            {
                p.WriteToLogAndConsole("    Invalid ExecuteSQLFilesSeparately setting. Defaulting to ExecuteSQLFilesSeparately = false ");
                p.ExecuteSQLFilesSeparately = false;
            }

            try
            {
                p.GenerateExcelFilesFromSelectQueryFiles = bool.Parse(ConfigurationManager.AppSettings["GenerateExcelFilesFromSelectQueryFiles"].ToString());
                p.WriteToLogAndConsole("    Using setting GenerateExcelFilesFromSelectQueryFiles = " + p.GenerateExcelFilesFromSelectQueryFiles.ToString());
            }
            catch (Exception)
            {
                p.WriteToLogAndConsole("    Invalid GenerateExcelFilesFromSelectQueryFiles setting. Defaulting to GenerateExcelFilesFromSelectQueryFiles = false ");
                p.GenerateExcelFilesFromSelectQueryFiles = false;
            }

            try
            {
                p.ConsolidateExcelFilesIntoSingleFileOnSeperateTabs = bool.Parse(ConfigurationManager.AppSettings["ConsolidateExcelFilesIntoSingleFileOnSeperateTabs"].ToString());
                p.WriteToLogAndConsole("    Using setting ConsolidateExcelFilesIntoSingleFileOnSeperateTabs = " + p.ConsolidateExcelFilesIntoSingleFileOnSeperateTabs.ToString());
            }
            catch (Exception)
            {
                p.WriteToLogAndConsole("    Invalid ConsolidateExcelFilesIntoSingleFileOnSeperateTabs setting. Defaulting to ConsolidateExcelFilesIntoSingleFileOnSeperateTabs = false ");
                p.ConsolidateExcelFilesIntoSingleFileOnSeperateTabs = false;
            }

            p.WriteToLogAndConsole("Settings read...");
        }

        /// <summary>
        /// Write messages to the console and also to the current log file
        /// </summary>
        /// <param name="p_Message"></param>
        void WriteToLogAndConsole(string p_Message)
        {
            //string DateTimeStamp = DateTime.Now.Year.ToString() + "." + DateTime.Now.Month.ToString() + "." + DateTime.Now.Day.ToString() + "_h" + DateTime.Now.Hour.ToString() + "_m" + DateTime.Now.Minute.ToString() + "_s" + DateTime.Now.Second.ToString() + "_ms" + DateTime.Now.Millisecond.ToString();
            string DateTimeStamp = DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss ");
            DateTimeStamp = DateTimeStamp.PadRight(25, '.'); // add a date time stamp for each message. also try to keep the formatting of the file the same for consistency

            Console.WriteLine(DateTimeStamp + " : " + p_Message);

            try
            {
                if (!Directory.Exists(Path.GetDirectoryName(this.RunLogFile)))
                    Directory.CreateDirectory(Path.GetDirectoryName(this.RunLogFile));

                FileStream fs;

                if (File.Exists(this.RunLogFile))
                {
                    fs = new FileStream(this.RunLogFile, FileMode.Append);
                }
                else
                {
                    fs = new FileStream(this.RunLogFile, FileMode.OpenOrCreate);
                }

                using (StreamWriter sw = new StreamWriter(fs))
                {
                    sw.WriteLine(DateTimeStamp + " : " + p_Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ex.Message : " + ex.Message.ToString());
                Console.WriteLine("");
                Console.WriteLine("Ex.GetBaseException() : " + ex.GetBaseException().ToString());

                if (System.Diagnostics.Debugger.IsAttached)
                    Console.ReadLine();
            }
        }

        /* decided to use just the single write to file so all available things can be written immediately to a file instead of batches.
            Doing it in batches would be more efficient as far as resources goes but neglible unless the log file gets huge.
        public void WriteToLogAndConsole(string p_File, List<string> p_Messages)
        {
            using (StreamWriter sw = new StreamWriter(p_File))
            {
                foreach (string message in p_Messages)
                {
                    string DateTimeStamp = DateTime.Now.Year.ToString() + "." + DateTime.Now.Month.ToString() + "." + DateTime.Now.Day.ToString() + "_" + DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString();
                    DateTimeStamp = DateTimeStamp.PadRight(25, '.'); // add a date time stamp for each message. also try to keep the formatting of the file the same for consistency

                    sw.WriteLine(DateTimeStamp + " : " + message);

                    Console.WriteLine(DateTimeStamp + " : " + message);
                }
            }
        }
         */

        /// <summary>
        /// Executes the sql files according to the settings
        /// </summary>
        /// <param name="p_Processor"></param>
        public void ProcessSQLFiles(ref Processor p_Processor)
        {
            // read all of the sql files.
            // it will populate ParsedSqlFilesToProcess list.
            // if the setting is to execute the files separately, then ParsedSqlFilesToProcess will contain an element for each file of sql data

            ReadAndFormatSQLFiles(p_Processor);

            try
            {
                p_Processor.WriteToLogAndConsole("    Executing SQL batch ");

                using (System.Data.SqlClient.SqlConnection connection = new System.Data.SqlClient.SqlConnection(p_Processor.connectionString))
                {
                    p_Processor.WriteToLogAndConsole("    Opening connection");
                    connection.Open();
                    p_Processor.WriteToLogAndConsole("    Connection Open");
                    foreach (SqlFileToProcess sqlfile in p_Processor.ParsedSqlFilesToProcess)
                    {
                        string transactionName = "DAPProcessor_" + DateTime.Now.Year.ToString() + "_" + DateTime.Now.Month.ToString() + "_" + DateTime.Now.Day.ToString() + "_" + DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString();
                        SqlTransaction transaction = null;

                        DateTime ProcessStartTime = DateTime.Now;
                        if (UseTransaction)
                        {
                            transaction = connection.BeginTransaction(System.Data.IsolationLevel.ReadUncommitted, transactionName);
                            p_Processor.WriteToLogAndConsole("        Begin transaction...'" + transactionName + "'");
                        }

                        try
                        {
                            p_Processor.WriteToLogAndConsole("            SQL Executing");

                            p_Processor.WriteToLogAndConsole("                Executing file " + sqlfile.FileName);
                            
                            SqlCommand cmd = new SqlCommand("sp_executesql ", connection);

                            cmd.CommandTimeout = p_Processor.TimeOut;

                            if (UseTransaction)
                                cmd.Transaction = transaction;

                            cmd.CommandType = System.Data.CommandType.StoredProcedure;
                            cmd.Parameters.Add("@statement", System.Data.SqlDbType.NText).Value = sqlfile.RawSQL;

                            if (p_Processor.GenerateExcelFilesFromSelectQueryFiles && sqlfile.FileName.ToLower().IndexOf("select") > 0)
                            {
                                DataSet ds = new DataSet();
                                SqlDataAdapter da = new SqlDataAdapter(cmd);
                                da.Fill(ds);

                                if (!ConsolidateExcelFilesIntoSingleFileOnSeperateTabs)
                                {
                                    p_Processor.GenerateExcelFile(ref ds);
                                }
                                else
                                {
                                    p_Processor.DataSetsForExcel.Add(ds);
                                }
                            }
                            else
                            {
                                cmd.ExecuteNonQuery();
                            }

                            p_Processor.WriteToLogAndConsole("            SQL Executed successfully");

                            try
                            {
                                if (UseTransaction)
                                {
                                    transaction.Commit();
                                    p_Processor.WriteToLogAndConsole("        Transaction committed");
                                }
                            }
                            catch (Exception ex1)
                            {
                                // This catch block will handle any errors that may have occurred 
                                // on the server that would cause the rollback to fail, such as 
                                // a closed connection.
                                p_Processor.WriteToLogAndConsole("        Rollback Exception Type: " + ex1.GetType().ToString());
                                p_Processor.WriteToLogAndConsole("          Message: " + ex1.Message.ToString());
                                throw;
                            }

                            DateTime ProcessEndTime = DateTime.Now;
                            TimeSpan ProcessTime = ProcessEndTime - ProcessStartTime;

                            p_Processor.WriteToLogAndConsole("        Process Time = " + ProcessTime.Hours.ToString() + ":" + ProcessTime.Minutes.ToString() + ":" + ProcessTime.Seconds.ToString());
                        }
                        catch (Exception ex)
                        {
                            p_Processor.WriteToLogAndConsole("        Commit Exception Type: " + ex.GetType().ToString());
                            p_Processor.WriteToLogAndConsole("        Message: " + ex.Message.ToString());

                            // Attempt to roll back the transaction. 
                            try
                            {
                                if (UseTransaction)
                                {
                                    transaction.Rollback();
                                    p_Processor.WriteToLogAndConsole("        Transaction rolledback");
                                }
                            }
                            catch (Exception ex2)
                            {
                                // This catch block will handle any errors that may have occurred 
                                // on the server that would cause the rollback to fail, such as 
                                // a closed connection.
                                p_Processor.WriteToLogAndConsole("        Rollback Exception Type: " + ex2.GetType().ToString());
                                p_Processor.WriteToLogAndConsole("          Message: " + ex2.Message.ToString());
                                throw;
                            }
                            throw;
                        }
                    }
                }

                p_Processor.WriteToLogAndConsole("    Closing connection");
            }
            catch (Exception ex)
            {
                p_Processor.WriteToLogAndConsole("Exception occurred!");
                p_Processor.WriteToLogAndConsole("Exception.Message = " + ex.Message);
                p_Processor.WriteToLogAndConsole("");
                p_Processor.WriteToLogAndConsole("Exception.GetBaseException() = " + ex.GetBaseException().ToString());

                throw;
            }
        }

        /// <summary>
        /// Reads the sql files and applies some simple formatting in order to prepare them to be executed
        /// </summary>
        /// <param name="p_Processor"></param>
        private static void ReadAndFormatSQLFiles(Processor p_Processor)
        {
            string sql = string.Empty;

            if (!p_Processor.ExecuteSQLFilesSeparately)
            {
                p_Processor.WriteToLogAndConsole("    Combining SQL files into a single string for batch execution");
            }

            foreach (string file in p_Processor.SqlFilesToProcess)
            {
                try
                {
                    p_Processor.WriteToLogAndConsole("        Reading file : " + file);

                    FileStream fs = new FileStream(file, FileMode.Open);
                    StreamReader reader = new StreamReader(fs);

                    while (!reader.EndOfStream)
                    {
                        string currentline = reader.ReadLine();

                        // remove comments since they are not needed for executing the sql
                        currentline = RemoveSQLComments(currentline);

                        sql += currentline + " "; // add a space to each line of the file to help prevent certain syntax errors
                    }
                    reader.Close();

                    p_Processor.WriteToLogAndConsole("        File read");

                    // lighten the sql foot print by removing sql formatting
                    sql = sql.Replace("\t", " ");

                    // replace all of the double spacing with single space
                    sql = System.Text.RegularExpressions.Regex.Replace(sql, @"\s{2,}", " ");

                    //repeat incase any was missed
                    sql = sql.Replace("\t", " ");
                    sql = System.Text.RegularExpressions.Regex.Replace(sql, @"\s{2,}", " ");

                    sql = sql.Replace("'", "''");

                    // prepare the query to be executed through parameterization
                    // this also has a good chance that the query plan will be cached and re-used for frequently used queries
                    sql = "DECLARE @stmt nvarchar(MAX)='" + sql + "' EXEC sp_executesql @stmt ";

                    if (!p_Processor.IsProduction)
                    {
                        sql = sql.ToLower().Replace("bol..", "BOLTest..");
                    }

                    // if we need to process them separately then add a new element to the list
                    if (p_Processor.ExecuteSQLFilesSeparately)
                    {
                        try
                        {
                            p_Processor.ParsedSqlFilesToProcess.Add(new SqlFileToProcess(Path.GetFileName(file), sql));
                        }
                        catch (Exception)
                        {
                            p_Processor.ParsedSqlFilesToProcess.Add(new SqlFileToProcess("", sql));
                        }
                        //p_Processor.ParsedSqlFilesToProcess.Add(new SqlFileToProcess(;
                        sql = string.Empty;
                    }
                }
                catch
                {
                    throw;
                }
            }

            // if we need to process all files in a single batch then add a single element to the list
            if (!p_Processor.ExecuteSQLFilesSeparately)
            {
                string BatchFileName = "FilesExecutedInBatch";
                if(p_Processor.GenerateExcelFilesFromSelectQueryFiles)
                {
                    foreach (string sqlfile in p_Processor.SqlFilesToProcess)
                    {
                        if (sqlfile.ToLower().IndexOf("select") > -1)
                        {
                            BatchFileName += "_select";
                            break;
                        }
                    }
                }
                p_Processor.ParsedSqlFilesToProcess.Add(new SqlFileToProcess("FilesExecutedInBatch", sql));
                sql = string.Empty;
            }

            p_Processor.WriteToLogAndConsole("    SQL files read and formatted");
        }

        /// <summary>
        /// Removes single line SQL comments only
        /// </summary>
        /// <param name="currentline"></param>
        /// <returns></returns>
        private static string RemoveSQLComments(string currentline)
        {
            int beginBlockComment = -1;
            int endBlockComment = -1;

            //does not properly work since it reads 1 line at a time, multiple lined block comments are not removed properly
            // // remove block comments
            // while (currentline.IndexOf("/*") > -1)
            // {
            //     beginBlockComment = currentline.IndexOf("/*");
            //
            //     // if the current line contains a full block comment then remove that whole comment, else remove the rest of the line starting with /*
            //     if (currentline.IndexOf("*/") > -1)
            //     {
            //         endBlockComment = currentline.IndexOf("*/") + 2;
            //     }
            //     else
            //     {
            //         // if the end does not exist, then remove the rest of the line.
            //         endBlockComment = currentline.Length - 1;
            //     }
            //
            //     currentline = currentline.Remove(beginBlockComment, (endBlockComment - beginBlockComment));
            // }
            //
            // while (currentline.IndexOf("*/") > -1)
            // {
            //     endBlockComment = currentline.IndexOf("*/") + 2;
            //
            //     // if the current line contains a full block comment then remove that whole comment, else remove the rest of the line starting with /*
            //     if (currentline.IndexOf("/*") > -1)
            //     {
            //         beginBlockComment = currentline.IndexOf("/*");
            //     }
            //     else
            //     {
            //         // if the begin does not exist, then remove the rest of the line starting with the beginning of the string
            //         beginBlockComment = 0;
            //     }
            //
            //     currentline = currentline.Remove(beginBlockComment, (endBlockComment - beginBlockComment));
            // }

            // remove singlie line comments
            while (currentline.IndexOf("--") > -1)
            {
                beginBlockComment = currentline.IndexOf("--");
                endBlockComment = currentline.Length;
                currentline = currentline.Remove(beginBlockComment, (endBlockComment - beginBlockComment));
            }
            return currentline;
        }

        /// <summary>
        /// Generate an Excel file from a dataset
        /// </summary>
        /// <param name="p_ds"></param>
        private void GenerateExcelFile(ref DataSet p_ds)
        {
            this.WriteToLogAndConsole("                Generating Excel file");
            string ExcelFileName = string.Empty;
            string TabName = string.Empty;
            object MisValue = System.Reflection.Missing.Value;

            Excel.Application xlApp = new Excel.Application();

            Excel.Workbook xlWorkBook = xlApp.Workbooks.Add(System.Reflection.Missing.Value);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            
            try
            {
                foreach (DataRow dr in p_ds.Tables[0].Rows)
                {
                    ExcelFileName = dr["FileName"].ToString();

                    if (ExcelFileName.Length > 0)
                        break;
                }

                ExcelFileName = ExcelFileName.Replace(".xlsx", ".xls");
                if (ExcelFileName.ToLower().IndexOf(".xls") < 0)
                {
                    ExcelFileName = ExcelFileName + ".xls";
                }

                ExcelFileName = this.ProcessPath + @"\" + ExcelFileName;

                for (int i = 0; i <= p_ds.Tables[0].Rows.Count - 1; i++)
                {
                    // write out the column names
                    if (i == 0)
                    {
                        DataColumnCollection columns = p_ds.Tables[0].Columns;
                        for (int i2 = 1; i2 < columns.Count; i2++)
                        {
                            xlWorkSheet.Cells[(i + 1), (i2)] = columns[i2].ColumnName;
                        }
                    }

                    // write the each data row to the excel document skipping the first field which is the FileName.
                    for (int j = 1; j <= p_ds.Tables[0].Columns.Count - 1; j++)
                    {
                        xlWorkSheet.Cells[(i + 2), j] = p_ds.Tables[0].Rows[i].ItemArray[j].ToString();
                    }
                }

                xlWorkBook.SaveAs(ExcelFileName, Excel.XlFileFormat.xlWorkbookNormal, MisValue, MisValue, MisValue, MisValue, Excel.XlSaveAsAccessMode.xlExclusive, MisValue, MisValue, MisValue, MisValue, MisValue);
                xlWorkBook.Close(true, MisValue, MisValue);
                xlApp.Quit();

                releaseObject(xlApp);
                releaseObject(xlWorkBook);
                releaseObject(xlWorkSheet);
            }
            catch (Exception)
            {
                xlWorkBook.Close(true, MisValue, MisValue);
                xlApp.Quit();

                releaseObject(xlApp);
                releaseObject(xlWorkBook);
                releaseObject(xlWorkSheet);

                throw;
            }

            this.WriteToLogAndConsole("                Excel file generated");
        }

        /// <summary>
        /// Generates a single excel file with each dataset in a separate tab. Tabname is provided in each script
        /// </summary>
        private void GenerateConsolodatedExcelFile()
        {
            this.WriteToLogAndConsole("                Generating Excel file");
            string FileName = string.Empty;
            string TabName = string.Empty;
            object MisValue = System.Reflection.Missing.Value;
            
            Excel.Application xlApp = new Excel.Application();
            
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Add(System.Reflection.Missing.Value);
            
            try
            {
                int ExcelSheetIndex = 1;

                foreach (DataSet ds in this.DataSetsForExcel)
                {
                    Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(ExcelSheetIndex);

                    foreach (DataRow dr in ds.Tables[0].Rows)
                    {
                        if (FileName.Length == 0)
                        {
                            FileName = dr["FileName"].ToString();
                            FileName = this.ProcessPath + @"\" + FileName;
                            FileName = FileName.Replace(".xlsx", ".xls");
                            if (FileName.ToLower().IndexOf(".xls") < 0)
                            {
                                FileName = FileName + ".xls";
                            }
                        }

                        TabName = dr["TabName"].ToString();

                        if (FileName.Length > 0 && TabName.Length > 0)
                            break;
                    }


                    for (int i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
                    {
                        // write out the column names
                        if (i == 0)
                        {
                            DataColumnCollection columns = ds.Tables[0].Columns;
                            for (int i2 = 2; i2 < columns.Count; i2++)
                            {
                                xlWorkSheet.Cells[(i + 1), (i2 - 1)] = columns[i2].ColumnName;
                            }
                        }

                        // write the each data row to the excel document skipping the first field which is the FileName.
                        for (int j = 2; j <= ds.Tables[0].Columns.Count - 1; j++)
                        {
                            xlWorkSheet.Cells[(i + 2), j - 1] = ds.Tables[0].Rows[i].ItemArray[j].ToString();
                        }
                    }

                    xlWorkSheet.Name = TabName;

                    ExcelSheetIndex++;
                }
                xlWorkBook.SaveAs(FileName, Excel.XlFileFormat.xlWorkbookNormal, MisValue, MisValue, MisValue, MisValue, Excel.XlSaveAsAccessMode.xlExclusive, MisValue, MisValue, MisValue, MisValue, MisValue);
                xlWorkBook.Close(true, MisValue, MisValue);
                xlApp.Quit();

                releaseObject(xlApp);
                releaseObject(xlWorkBook);
            }
            catch (Exception)
            {
                xlWorkBook.Close(true, MisValue, MisValue);
                xlApp.Quit();

                releaseObject(xlApp);
                releaseObject(xlWorkBook);

                throw;
            }

            this.WriteToLogAndConsole("                Excel file generated");
        }

        /// <summary>
        /// Releases the supplied Excel object 
        /// </summary>
        /// <param name="obj"></param>
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
