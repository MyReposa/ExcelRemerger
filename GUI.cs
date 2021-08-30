using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelRemerger
{
    public partial class GUI : Form
    {
        Demerger demerger;
        Logger logger;

        public GUI()
        {
            InitializeComponent();
            this.logger = new Logger();
            logger.sendMsgToDisplay += DisplayMsg;
            this.demerger = new Demerger(this.logger);
        }
        
        private void ButtonStart_Click(object sender, EventArgs e)
        {
             demerger.Work(textBoxFilePath.Text);
        }

        void DisplayMsg(object sender, string msgToDisplay)
        {
            textBoxMainDisplay.AppendText($"• {msgToDisplay}\r\n\r\n");
        }
    }
}
