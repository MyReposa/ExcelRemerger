using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelRemerger
{
    class Logger
    {
        public event EventHandler<string> sendMsgToDisplay;

        public void Log(string textToSend)
        {
            sendMsgToDisplay?.Invoke(this, textToSend);
        }
    }
}
