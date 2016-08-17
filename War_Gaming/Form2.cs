using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace WindowsFormsApplication2
{
    public partial class Form2 : Form
    {
        //List<string> filenames;
        string[] oemList, filenames;
        string path;
        string selectedOemName;
        int[] claimCount;
        public string returnRequestFile { get; set; }

        public Form2()
        {
            InitializeComponent();
            CenterToScreen();
            button1.Enabled = false;
            bool form2_local = Form1.local;

            // central player machine ID
            string machineID = "WGC100D5NQW52";

            if ( !form2_local )
                path = @"\\" + machineID + @"\sandbox";
            else
                path = "c:/WG/sandBox";

            int cc = 0;
            string playlistfile = path + @"\announcement\OEM_list.txt";
            using (StreamReader reader = new StreamReader(playlistfile))
            {
                while (!reader.EndOfStream)
                {
                    string line = reader.ReadLine();
                    string[] pair = line.Split(',');
                    cc++;
                }
            }
            oemList = new string[cc];
            filenames = new string[cc];
            using (StreamReader reader = new StreamReader(playlistfile))
            {
                for (int i = 0; i < cc; i++)
                {
                    string line = reader.ReadLine();
                    string[] pair = line.Split(',');
                    oemList[i] = pair[0];
                    filenames[i] = pair[1];
                }
            }

/*            filenames = new List<string>();
            filenames.Add("ford");
            filenames.Add("toyota");
            filenames.Add("nissan");
            filenames.Add("honda");
            filenames.Add("bmw");
            filenames.Add("hyundai");*/
//            oemList = new string[] { "FORD", "BMW", "DAIMLER", "FIAT", "FUJI HEAVY", "GMC", "HONDA", "HYUNDAI",
//               "MAZDA", "MITSUBISHI", "NISSAN", "TATA", "TOYOTA", "VOLKSWAGEN", "VOLVO" };
//           filenames = new string[] { "ford", "bmw", "benz", "chrysler", "fuji", "gm", "honda", "hyundai",
//               "mazda", "mitsubishi", "nissan", "tata", "toyota", "vw", "volvo" };

            /* int count = 0;
            claimCount = new int[filenames.Length];
            for (var ii = 0; ii < filenames.Length; ii++)
            {
                string directoryPath = path + @"\" + oemList[ii];
                if (File.Exists(directoryPath + @"\" + @"*.request"))
                {
                    count = claimCount[ii] = 1;
                    MessageBox.Show(count.ToString());
                }

            } */

            claimCount = new int[filenames.Length];
            for (var ii = 0; ii < filenames.Length; ii++)
            {
                string directoryPath = path + @"\" + oemList[ii];
                if (File.Exists(directoryPath + @"\" + @"*.request")) claimCount[ii] = 1;
            }


            // add picturebox controls to FlowLayoutPanel
            for (var ii = 0; ii < filenames.Length; ii++)
            {
                string oemName = filenames[ii];
                string oemName1 = oemName + "1.jpg";
                PictureBox pic = new PictureBox();
                pic.ClientSize = new Size(100, 86);
                pic.BorderStyle = BorderStyle.FixedSingle;
                Image image;   // add by JF

                if (claimCount[ii] == 0)
                {
                    image = Image.FromFile("C:\\WG\\Resources\\"+oemName1);
                    // pic.Image = (Image)Properties.Resources.ResourceManager.GetObject(oemName + Convert.ToString(1));
                    pic.Image = image;
                    pic.Enabled = true;
                    toolTip1.SetToolTip(pic,oemName.ToUpper() + " is available for players");
                }
                else
                {
                    // pic.Image = (Image)Properties.Resources.ResourceManager.GetObject(oemName + Convert.ToString(2));
                    image = Image.FromFile("C:\\War_Gaming\\War_Gaming\\Resources\\" + oemName1);
                    pic.Image = image;
                    pic.Enabled = false;
                    toolTip1.SetToolTip(pic, oemName.ToUpper() + " is claimed by other players");
                }

                // If the image is too big, zoom.
                if ((pic.Image.Width > 100) || (pic.Image.Height > 86))
                {
                    pic.SizeMode = PictureBoxSizeMode.Zoom;
                }
                else
                {
                    pic.SizeMode = PictureBoxSizeMode.CenterImage;
                }

                // Add the Click event handler.
                pic.Click += PictureBoxesClick;
                pic.Tag = oemName;
                pic.Parent = flpThumbnails;
                //pic.BackColor = Color.LawnGreen;

            }

        }

        private IEnumerable<Control> ChildControls(Control parent)
        {
            List<Control> controls = new List<Control>();
            controls.Add(parent);
            foreach (Control ctrl in parent.Controls)
            {
                controls.AddRange(ChildControls(ctrl));
            }
            return controls;
        }

        private void PictureBoxesClick(object sender, EventArgs e)
        {
            if (sender is PictureBox)
            {
                if (((PictureBox)sender).Enabled == false) return;

                // the selected OEM has not been claimed
                DialogResult dialogResult = MessageBox.Show("Are you sure you want to select " + ((PictureBox)sender).Tag.ToString().ToUpper() + "?", "Selection of OEM", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_ffffff");

                    string OemMachinName = Environment.MachineName;
                    string oemClaimed = ((PictureBox)sender).Tag.ToString();
                    string userID = Environment.UserName;
                    int keyIndex = Array.FindIndex(filenames, w => w.Contains(oemClaimed));
                    //MessageBox.Show(tickName);

                    string directoryPath = path + @"\" + oemList[keyIndex];
                    string[] filePaths = Directory.GetFiles(directoryPath + @"\", "*.request");
                    if (filePaths.Count() > 0) 
                    {
                        MessageBox.Show("This OEM has been claimed. Please select another one.");
                        return;
                    }

                    string tickName = path + @"\" + oemList[keyIndex] + @"\" + userID + "_" + OemMachinName + "_" + oemClaimed + "_" + timestamp + ".request";
                    TextWriter tw = new StreamWriter(tickName, false);
                    tw.WriteLine("");
                    tw.Close();

                    string usermessage = oemClaimed + " is claimed by the player from " + OemMachinName + "\r\n";
                    //richTextBox1.AppendText(Environment.NewLine + usermessage + "\r\n");
                    //richTextBox1.SelectionAlignment = HorizontalAlignment.Right;

                    string grayfname = ((PictureBox)sender).Tag.ToString() + Convert.ToString(2);
                    ((PictureBox)sender).Image = (Image)Properties.Resources.ResourceManager.GetObject(grayfname);
                    ((PictureBox)sender).Enabled = false;
                    selectedOemName = oemClaimed;
                    //selectedOemName = oemList[keyIndex];
                    button1.Enabled = true;

                    // you can only select once in a game
                    var list=ChildControls(flpThumbnails).ToList();
                    foreach (Control ctl in list)
                    {
                        if (ctl is PictureBox) ctl.Enabled = false;
                    }

                    this.returnRequestFile = tickName;
                    
                }
                else if (dialogResult == DialogResult.No)
                {
                    //do something else
                }


            }
        }


        public string getText()
        {
            return selectedOemName;
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }
    }
}
