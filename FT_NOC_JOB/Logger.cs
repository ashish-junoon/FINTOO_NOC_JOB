using System;
using System.Configuration;
using System.IO;
using System.Runtime.CompilerServices;

class Logger
{
    private static readonly string logDirectory = ConfigurationManager.AppSettings["LogDirectory"];
    private static readonly string logFile = Path.Combine(logDirectory, ConfigurationManager.AppSettings["LogFileName"]);

    static Logger()
    {
        if (!Directory.Exists(logDirectory))
        {
            Directory.CreateDirectory(logDirectory);
        }
    }

    public static void LogMessage(string message)
    {
        using (StreamWriter sw = new StreamWriter(logFile, true))
        {
            sw.WriteLine($"{DateTime.Now}: {message}");
        }
    }

}
