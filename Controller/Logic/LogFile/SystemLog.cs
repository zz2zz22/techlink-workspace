using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace techlink_workspace.Controller.Logic.LogFile
{
    public class SystemLog
    {
        static SystemLog s_myInstance = null;
        string m_startUpPath = "";
        public enum MSG_TYPE
        {
            Nor,
            Err,
            War
        };
        private SystemLog()
        {
            string dirPath = AppDomain.CurrentDomain.BaseDirectory + "\\Logfile";
            System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(dirPath);
            if (dir.Exists == false)
                dir.Create();
            m_startUpPath = dirPath + "\\Log_";
        }
        public static void DeleteOldLog() //Delete 3 months old logs
        {
            string dirPath = AppDomain.CurrentDomain.BaseDirectory + "\\Logfile";
            System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(dirPath);
            if (dir.Exists == false)
                dir.Create();

            string[] files = Directory.GetFiles(dirPath);

            foreach (string file in files)
            {
                FileInfo fi = new FileInfo(file);
                if (fi.CreationTime < DateTime.Now.AddMonths(-3))
                    fi.Delete();
            }
        }

        public static void Output(MSG_TYPE msgType, string name, string str)
        {
            if (s_myInstance == null)
                s_myInstance = new SystemLog();
            s_myInstance.logout(msgType, name, str);
        }
        public static void Output(MSG_TYPE msgType, string name, string format, params object[] args)
        {
            if (s_myInstance == null)
                s_myInstance = new SystemLog();
            s_myInstance.logout(msgType, name, string.Format(format, args));
        }
        private void logout(MSG_TYPE msgType, string name, string str)
        {
            string filePath = m_startUpPath + DateTime.Now.ToString("yyyyMMdd") + ".txt";
            string output = name + " : " + str + "\r\n";
            try
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@filePath, true))
                {
                    file.WriteLine(DateTime.Now.ToString("HH:mm:ss.fff ") + output);
                }
            }
            catch (Exception) { }

            EventBroker.AsyncSend(EventBroker.EventID.etLog, new EventBroker.EventParam(this, (int)msgType, output));
        }

        private string getFilePath()
        {
            string filePath = m_startUpPath + DateTime.Now.ToString("yyyyMMdd") + ".txt";
            return filePath;
        }

        public static int logcount(string key, string data)
        {
            if (s_myInstance == null)
                s_myInstance = new SystemLog();
            try
            {
                int count = File.ReadLines(s_myInstance.getFilePath())
                    .Count(line => line.Contains(key) && line.Contains(data));
                return count;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading log file: {ex.Message}");
                return 0;
            }
        }

    }
}
