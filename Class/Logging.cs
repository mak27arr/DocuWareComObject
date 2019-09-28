using System;
using System.Collections.Generic;
using System.IO;
using System.IO.IsolatedStorage;
using System.Runtime.InteropServices;
using System.Text;

namespace DocuWare.Class
{

    public static class Logging
    {
        private static object syncObj = new object();
        private static string setting_file_name = "log.txt";
        private static long log_file_max_size = 20 * 1024 * 1024;//20Mb
        public static void Log(string msg, Object obj = null)
        {
            string patc = getLogFileLocation();
            string msg_formated = formatMg(msg, obj);
            cheakLogFileSize();
            lock (syncObj)
            {
                try
                {
                    using (StreamWriter sw = new StreamWriter(patc, true, System.Text.Encoding.Default))
                    {
                        sw.WriteLine(msg_formated);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
            }
        }

        private static string formatMg(string msg, object obj)
        {
            string msg_ret = Environment.NewLine +
                DateTime.Now.ToString() +
                Environment.NewLine +
                "Message: " + msg;
            Exception ex = obj as Exception;
            if (ex != null)
            {
                msg_ret += Environment.NewLine +
                    "Error mssage: " + ex.Message + Environment.NewLine +
                    "Class: " + ex.Source + Environment.NewLine +
                    "Full : " + ex.ToString();
            }
            return msg_ret;
        }
        private static string getLogFileLocation()
        {
            string patch = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
            patch += "\\" + setting_file_name;
            return patch;
        }

        private static void cheakLogFileSize()
        {
            string patch = getLogFileLocation();
            if (System.IO.File.Exists(patch))
            {
                if (log_file_max_size < new System.IO.FileInfo(patch).Length)
                {
                    lock (syncObj)
                    {
                        using (StreamWriter sw = new StreamWriter(new IsolatedStorageFileStream(patch, FileMode.Truncate, null)))
                        {
                            sw.WriteLine(patch, "Log file clean: " + DateTime.Now.ToLongDateString());
                            sw.Close();
                        }
                    }
                }

            }
        }

    }
}
