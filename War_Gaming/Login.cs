using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication2
{
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
        }

        private void butLogin_Click(object sender, EventArgs e)
        {

        }

        private void butExit_Click(object sender, EventArgs e)
        {

        }

        private void butExit_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void butLogin_Click_1(object sender, EventArgs e)
        {
            if (textUser.Text != "jfang3")
                MessageBox.Show("You are not in Database user list");
            else
            {
                Form1 form1 = new Form1();
                form1.Show();
            }
        }
    }
}
