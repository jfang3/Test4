using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication2
{
    public partial class Form0 : Form
    {
        bool isFlashing;
        public Form0()
        {
            InitializeComponent();
            this.FormClosing += Form0_FormClosing;

        }

        private void Form0_FormClosing(object sender, FormClosingEventArgs e)
        {
            //this.Text = "Cancelled";
        }

        private void button1_Click(object sender, EventArgs e)
        {
        }

        private bool hasWriteAccessToFolder(string folderPath)
        {
            try
            {
                // Attempt to get a list of security permissions from the folder. 
                // This will raise an exception if the path is read only or do not have access to view the permissions. 
                System.Security.AccessControl.DirectorySecurity ds = Directory.GetAccessControl(folderPath);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private async void Blink()
        {
            while (isFlashing)
            {
                await Task.Delay(500);
                label4.BackColor = label4.BackColor == Color.Yellow ? this.BackColor : Color.Yellow;
            }
        }

        private void butLocal_Click(object sender, EventArgs e)
        {
            Form1.local = true;
            this.Close();
            return;
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            isFlashing = true;
            string txt = textBox1.Text.Trim();
            if (txt.Length == 0)
            {
                label4.Text = "Please fill the text box with the central machine ID.";
                button3.Enabled = false;
                Blink();
                return;
            }
            else
            {
                string path = @"\\" + txt + @"\sandbox\players";
                if (hasWriteAccessToFolder(path))
                {
                    this.Text = txt;
                    label4.Text = "A connection between you and the central player is established!";
                    label4.BackColor = Color.LawnGreen;
                    button1.Enabled = true;
                    button3.Enabled = false;
                }
                else
                {
                    button3.Enabled = false;
                    Blink();
                    string msg = "Connection failed:" + Environment.NewLine;
                    msg += "1) Make sure the correct machine ID is entered" + Environment.NewLine;
                    msg += "2) Make sure the network is working properly" + Environment.NewLine;
                    msg += "3) Contact the central player for more information";
                    label4.Text = msg;
                    //MessageBox.Show(msg);
                }
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            isFlashing = false;
            textBox1.Text = "";
            button1.Enabled = false;
            button3.Enabled = true;
            label4.Text = "";
            label4.BackColor = this.BackColor;

            this.Text = "Welcome Page";
        }

        private void button1_Click_1(object sender, EventArgs e)
        {

        }

        private void butDefault_Click(object sender, EventArgs e)
        {
            textBox1.Text = "WGC100D5NQW52";
        }
    }
}
