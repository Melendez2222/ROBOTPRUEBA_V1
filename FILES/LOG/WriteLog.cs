using ROBOTPRUEBA_V1.CONFIG;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ROBOTPRUEBA_V1.FILES.LOG
{
    internal class WriteLog
    {
        public void Log(string message)
        {
            string logMessage = $"{DateTime.Now}: {message}{Environment.NewLine}";
            File.AppendAllText(GlobalSettings.logFilePath, logMessage);
        }
    }
}
