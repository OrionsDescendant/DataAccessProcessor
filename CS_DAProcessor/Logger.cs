using System;
using System.Collections.Generic;
using System.IO;

namespace DAProcessor
{
    /// <summary>
    /// 
    /// </summary>
    public class Logger
    {
        public string ErrorLogFile = "ErrorLog.txt";
        public string RunLogFile = "RunLog.txt";

        public Logger(string p_LogPath)
        {
            string LogFilePrefix = DateTime.Now.Year.ToString() + "." + DateTime.Now.Month.ToString() + "." + DateTime.Now.Day.ToString() + "_" + DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString() + "_";
            ErrorLogFile = p_LogPath + "\\Logs\\" + LogFilePrefix + ErrorLogFile;
            RunLogFile = p_LogPath + "\\Logs\\" + LogFilePrefix + RunLogFile;
        }

        /// <summary>
        /// Write messages to the current log file
        /// </summary>
        /// <param name="p_File"></param>
        /// <param name="p_Messages"></param>
        public void WriteLog(string p_File, List<string> p_Messages)
        {
            using (StreamWriter sw = new StreamWriter(p_File))
            {
                foreach (string message in p_Messages)
                {
                    string DateTimeStamp = DateTime.Now.Year.ToString() + "." + DateTime.Now.Month.ToString() + "." + DateTime.Now.Day.ToString() + "_" + DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString();
                    DateTimeStamp = DateTimeStamp.PadRight(25, '.'); // add a date time stamp for each message. also try to keep the formatting of the file the same for consistency

                    sw.WriteLine(p_Messages);
                }
            }
        }
    }
}
