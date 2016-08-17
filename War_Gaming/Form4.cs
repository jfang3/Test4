using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Mail;

using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Core;
// using OutlookApp = Microsoft.Office.Interop.Outlook._Application;


namespace WindowsFormsApplication2
{
    public partial class Form4 : Form
    {
        OpenFileDialog OpenFileDialog1 = new OpenFileDialog();

        public Form4()
        {
            InitializeComponent();
        }

        private void Form4_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textAttachment.Text = OpenFileDialog1.FileName.ToString();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("Software\\microsoft\\windows\\currentversion\\app paths\\OUTLOOK.EXE");
            string path = (string)key.GetValue("Path");
            if (path != null)
            {
                // System.Diagnostics.Process.Start("OUTLOOK.EXE");
                string subjected = textSubject.Text;
                string attached = textAttachment.Text;

                System.Diagnostics.Process.Start("OUTLOOK.EXE", "/c ipm.note /m jfang3@ford.com&subject=War_Gaming");
            }
            else
                MessageBox.Show("There is no Outlook in this computer!", "SystemError", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
           

        }
    }
}
