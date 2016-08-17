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

using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Data.Sql;
using System.Data.Common;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Data.ProviderBase;
using System.Data.SqlTypes;

using System.Net;
using System.Net.Mail;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.Threading;
using System.Diagnostics;

using System.Collections;

using System.Windows.Forms.DataVisualization.Charting;

using System.Text.RegularExpressions;  // JF 2/29/16

using Microsoft.Office.Interop.Outlook;

using OutlookApp = Microsoft.Office.Interop.Outlook.Application;


namespace WindowsFormsApplication2
{
    public partial class Form1 : Form
    {
        int group1_value, group2_value, group3_value;
        bool player = true;
        string[] str = new string[8];
        int total;
        int[,] creditResults = new int[10, 2];
        int[,] CarCredit = new int[17, 2];
        int[,] TruckCredit = new int[17, 2];

        // added for sangdon's code
        private string importFile, exportFile;
        private Decision[] decision;
        // private int launchYr;
        // private int scenario;

        ArrayList vdata1, fordconstraints1, refresh1, basedelta1, fmccdata1, genPopu1, genvehdata1, genfcons1;
        Regex sp = new Regex("( )");

        // added for OEM client               
        private string selectedOem;
        private string sharedOem;
        private string OemMachinName;
       //  private Decision[] decision;
        private int scenario;
        ArrayList dicFordSeg = new ArrayList();
        ArrayList dicFordBrand = new ArrayList();
        private int comboxRowIndex, comboxColumnIndex;
        string[] oemList;
        string[] filenames;
        //List<string> filenames;
        string machineID;
        string path;
        int[] claimCount;
 
        bool rangeflag, runningyearflag, resetflag;
        bool gameflag, mdmflag, authflag;
        int runningyear, runningyear2;
        int beginyear;
        string byear, eyear;
        string newline;

        public static bool local;   // added by JF
        public static bool selectedCompare = false;
        public static bool EPAquery = false;
        public static bool NHTSAquery = false;
        public static bool FordCredit = false;
              public static string machine;

        BackgroundWorker backgroundWorker1 = new BackgroundWorker();
        Form2 frm2;
        string requestTicketName;


#if DEBUG
        const string DLL_PATH = "C:\\War_GamingV1.1\\War_Gaming\\bin\\G_compliance_64d.dll";
#else
        const string DLL_PATH = "C:\\War_GamingV1.1\\War_Gaming\\bin\\WG_compliance_64.dll";
#endif

        [DllImport(DLL_PATH, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Auto)]
        public static extern bool dll_module_init(string strNameplateData, int nNameplateData,
                                                  string strEPACreditReport, int nEPACreditReport,
                                                  string strNHTSACreditReport, int nNHTSACreditReport);

        [DllImport(DLL_PATH, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Auto)]
        public static extern bool dll_oem_init(string strOEM, int nStrOEM, int yearSt, int yearEd);

        [DllImport(DLL_PATH, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Auto)]
        public static extern int dll_oem_push_vol(string strOEM, int nStrOEM, string strNameplate, int nStrNameplate, int[] pVolume, int nSize);

        [DllImport(DLL_PATH, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Auto)]
        public static extern double dll_get_EPA_const_val();

        [DllImport(DLL_PATH, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Auto)]
        public static extern void dll_oem_finalize_vol(string strOEM, int nStrOEM, string strNameplate, int nStrNameplate, int volume, int year);

        [DllImport(DLL_PATH, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Auto)]
        public static extern void dll_module_destroy();

        public Form1()
        {
            InitializeComponent();
            // userListDataGridView.EditingControlShowing += new DataGridViewEditingControlShowingEventHandler(userListDataGridView_EditingControlShowing);
            // userListDataGridView_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e);
        }


        private void set_connect_center()
        {

            OemMachinName = Environment.MachineName;

            //DialogResult dialogOpen = MessageBox.Show("Use the navigation menu to get started.", "Welcome!", MessageBoxButtons.OK);
            //frm0 = new Form0();
            machineID = ShowMyDialogBox();
            // MessageBox.Show(machineID);
 
            machine = machineID;

            // this.FormClosing += mainForm_FormClosing;
            btnReadResults.Enabled = false;

            if ( local )
                MessageBox.Show("Use Local Machine as Central Player");
            if ( !local )
            {
                path = @"\\" + machineID + @"\sandbox";
                if (machineID == "Cancelled")
                {
                    try
                    {
                        Environment.Exit(0); // It will try to close your program "the hard way"
                    }
                    catch (System.Exception)
                    {
                        System.Windows.Forms.Application.Exit(); // If a Win32 exception occurs, then it will be able to close your program "the normal way"
                    }
                }
            }
            else
            {
                path = "c:/WG/sandbox";

            }

            // path = @"\\" + machineID + @"\sandbox";
            newline = Environment.NewLine;

            // this player is online
            string tickName = path + @"\players\" + Environment.UserName + ".on";
            // MessageBox.Show(tickName);
            TextWriter tw = new StreamWriter(tickName, false);
            tw.WriteLine("");
            tw.Close();

            requestTicketName = "";

            mdmflag = false;
            gameflag = false;
            authflag = false;
            rangeflag = false;
            resetflag = false;
            runningyearflag = false;
            runningyear = 0;
            runningyear2 = 0;

            // added in June 2016 by JX
            byear = null;
            eyear = null;
            // if (File.Exists(path + @"\announcement\game_on.txt")) gameflag = true;
            if (Directory.GetFiles(path + @"\announcement\", "from*.txt").Length == 1) rangeflag = true;
            if (File.Exists(path + @"\announcement\OEM_list.txt")) authflag = true;

            richTextBoxOEM.Text = "";
            richTextBoxOEM.TextChanged += new EventHandler(richTextBoxOEM_TextChanged);

            /*
            oemList = new string[] { "FORD", "BMW", "DAIMLER", "FIAT", "FUJI HEAVY", "GMC", "HONDA", "HYUNDAI",
                "MAZDA", "MITSUBISHI", "NISSAN", "TATA", "TOYOTA", "VOLKSWAGEN", "VOLVO" };
            filenames = new string[] { "ford", "bmw", "benz", "chrysler", "fuji", "gm", "honda", "hyundai",
                "mazda", "mitsubishi", "nissan", "tata", "toyota", "vw", "volvo" }; */

            oemList = new string[] { "FORD", "BMW", "DAIMLER", "FIATCHRYSLER", "SUBARU", "GM", "HONDA", "HYUNDAI",
                "MAZDA", "MITSUBISHI", "NISSAN", "JAGUARLANDROVER", "TOYOTA", "VOLKSWAGEN", "VOLVO" };
            filenames = new string[] { "ford", "bmw", "benz", "chrysler", "fuji", "gm", "honda", "hyundai",
                "mazda", "mitsubishi", "nissan", "tata", "toyota", "vw", "volvo" };

            claimCount = new int[filenames.Length];
            for (var ii = 0; ii < filenames.Length; ii++)
            {
                string directoryPath = path + @"\" + oemList[ii];
                var dir = new DirectoryInfo(directoryPath);
                claimCount[ii] = dir.EnumerateFiles("*.request").Count();
                // MessageBox.Show("n= " + claimCount[ii]);
                /*
                if (claimCount[ii] > 0)
                {
                    richTextBoxOEM.Text += oemList[ii] + " has been claimed" + "\r\n";
                    // MessageBox.Show(oemList[ii] + " has been claimed" + "\r\n");
                }*/
            }
            
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.WorkerSupportsCancellation = true;
            backgroundWorker1.DoWork += backgroundWorker1_DoWork;
            backgroundWorker1.ProgressChanged += backgroundWorker1_ProgressChanged;

            // splitContainer1.SplitterMoved += new SplitterEventHandler(splitContainer1_SplitterMoved); removed by JF
            // splitContainer2.SplitterMoved += new SplitterEventHandler(splitContainer2_SplitterMoved);

            dataGridView1.CellMouseDown += new DataGridViewCellMouseEventHandler(dataGridView1_CellMouseDown);
            dataGridView1.MouseClick += new MouseEventHandler(dataGridView1_MouseClick);

            backgroundWorker1.RunWorkerAsync();

            ToolTip ToolTip1 = new ToolTip();
            ToolTip1.SetToolTip(btnSelectOem, "Click to open a selection window");
            
        }

        public string ShowMyDialogBox()
        {
            string machineID;
            Form0 testDialog = new Form0();

            // Show testDialog as a modal dialog and determine if DialogResult = OK.
            if (testDialog.ShowDialog(this) == DialogResult.OK)
            {
                // Read the contents of testDialog's TextBox.
                //this.Text = testDialog.Text;
                machineID = testDialog.Text;
            }
            else
            {
                machineID = "Cancelled";
            }
            testDialog.Dispose();
            // MessageBox.Show(machineID);
            return machineID;
        }

        private void dataGridView1_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex >= 0 && e.RowIndex >= 0)
            {
                dataGridView1.CurrentCell = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
            }
        }

        private void dataGridView1_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                return;
            }
            else
            {
                ContextMenuStrip vehicle_menu = new ContextMenuStrip();
                int position_xy_mouse_row = dataGridView1.HitTest(e.X, e.Y).RowIndex;
                int position_xy_mouse_col = dataGridView1.HitTest(e.X, e.Y).ColumnIndex;

                if (position_xy_mouse_col != 2) return;
                if (position_xy_mouse_row < 0) return;

                comboxRowIndex = position_xy_mouse_row;
                comboxColumnIndex = position_xy_mouse_col;

                for (int ii = 0; ii < dicFordBrand.Count; ii++)
                {
                    object sss = dicFordBrand[ii];
                    vehicle_menu.Items.Add(sss.ToString()).Name = sss.ToString();
                }
                vehicle_menu.Show(dataGridView1, new Point(e.X, e.Y));

                //event menu click
                vehicle_menu.ItemClicked += new ToolStripItemClickedEventHandler(vehicle_menu_ItemClicked);
            }
        }

        private void vehicle_menu_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            string str = e.ClickedItem.Name.ToString();
            dataGridView1.Rows[comboxRowIndex].Cells[comboxColumnIndex].Value = str;
            int SelectedIndex = dicFordBrand.IndexOf(str);
            string seg = dicFordSeg[SelectedIndex].ToString();
            dataGridView1.Rows[comboxRowIndex].Cells[comboxColumnIndex - 1].Value = seg;
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            string str = e.UserState as string; ;
            richTextBoxOEM.Text += str + "\r\n";

            this.Text = str;
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            while (backgroundWorker1.CancellationPending == false)
            {
                Thread.Sleep(1000);

                //check game on/off
                /* if (!File.Exists(path + @"\announcement\game_on.txt") && gameflag == false)
                {
                    string sss = "The central player is offline." + newline + newline;
                    richTextBoxOEM.Invoke((MethodInvoker)delegate { richTextBoxOEM.Text += sss; });
                    gameflag = true;
                }
                if (File.Exists(path + @"\announcement\game_on.txt") && gameflag == true)
                {
                    string sss = "The central player is online." + newline + newline;
                    richTextBoxOEM.Invoke((MethodInvoker)delegate { richTextBoxOEM.Text += sss; });
                    btnSelectOem.Invoke((MethodInvoker)delegate { btnSelectOem.Enabled = true; });
                    gameflag = false;
                } */

                if (File.Exists(path + @"\announcement\OEM_list.txt") && authflag == false)
                {
                    string sss = "The central player has distributed some OEMs for you to select." + newline + newline;
                    richTextBoxOEM.Invoke((MethodInvoker)delegate { richTextBoxOEM.Text += sss; });
                    //                    btnSelectOem.Invoke((MethodInvoker)delegate { btnSelectOem.Enabled = true; });
                    authflag = true;
                }

                //year range announcement
                string[] rangefile = Directory.GetFiles(path + @"\announcement\", "from*.txt");
                if (rangefile.Length == 0 && rangeflag == false)
                {
                    string sss = "The game year range is not announced. Please wait..." + newline + newline;
                    sss += "If you haven't selected an OEM, you can try to select it while waiting." + newline + newline;
                    richTextBoxOEM.Invoke((MethodInvoker)delegate { richTextBoxOEM.Text += sss; });
                    rangeflag = true;
                }
                if (rangefile.Length == 1 && rangeflag == true)
                {
                    byear = rangefile[0].Substring(rangefile[0].Length - 16, 4);
                    eyear = rangefile[0].Substring(rangefile[0].Length - 8, 4);
                    string sss = "A game year range from " + byear + " to " + eyear + " is announced." + newline + newline;
                    richTextBoxOEM.Invoke((MethodInvoker)delegate { richTextBoxOEM.Text += sss; });
                    rangeflag = false;
                    runningyearflag = false;
                    GlobeData.BeginYear = byear;  // added by JF
                }

                
                // game is on but year range is changed
                // if (rangefile.Length == 0 && rangeflag == true) rangeflag = false; */

                // OEM declaration status
                /*
                for (var ii = 0; ii < filenames.Length; ii++)
                {
                    string oemName = oemList[ii];
                    string oempath = path + @"\" + oemName + @"\";
                    string[] filePaths = Directory.GetFiles(oempath, "*.request");

                    int cc = filePaths.Count();

                    if (cc > claimCount[ii])
                    {
                        claimCount[ii] = cc;
                        string requestfname = filePaths[0].Substring(oempath.Length, filePaths[0].Length - oempath.Length);
                        string userID = requestfname.Split('_')[0];
                        backgroundWorker1.ReportProgress(0, String.Format(oemName + " is claimed by " + userID));
                    }


                } */

                //check the market demand model announcement
                /*
                rangefile = Directory.GetFiles(path + @"\announcement\", "from*.txt");
                if (rangefile.Count() == 1 && rangeflag == false)
                {
                    string byear = rangefile[0].Substring(rangefile[0].Length - 16, 4);
                    string eyear = rangefile[0].Substring(rangefile[0].Length - 8, 4);
                    beginyear = Convert.ToInt32(byear);
                    rangeflag = true;
                    backgroundWorker1.ReportProgress(0, String.Format("ATTENTION: A new game is announced."));
                    backgroundWorker1.ReportProgress(0, String.Format("This game is from " + byear + " to " + eyear));
                    btnSelectOem.Invoke((MethodInvoker)delegate { btnSelectOem.Enabled = true; });
                } */

                string[] runningfile = Directory.GetFiles(path + @"\announcement\", "running*.txt");
                if (runningfile.Length == 1 && runningyearflag == false)
                {
                    runningyear = Convert.ToInt32(runningfile[0].Substring(runningfile[0].Length - 8, 4));
                    runningyearflag = true;
                    string sss = "It is time to play year of " + runningyear.ToString() + newline + newline;
                    richTextBoxOEM.Invoke((MethodInvoker)delegate { richTextBoxOEM.Text += sss; });
                    btnMdm.Invoke(new MethodInvoker(delegate { btnMdm.Enabled = true; }));
                    this.Invoke((MethodInvoker)delegate { this.Text = "You are playing year of " + runningyear.ToString(); });
                }

                string curFile = path + @"\announcement\EvaluationVolume_" + runningyear.ToString() + ".xls";
                if (File.Exists(curFile) && mdmflag == false)
                {
                    string str = "Market Demand Model of " + runningyear.ToString() + " is done." + newline + newline;
                    str += "The central player is prepairing input data for the next year." + newline + newline;
                    str += "Please don't touch any part on your interface." + newline + newline;
                    if (runningyear.ToString() == eyear)
                    {
                        str = "The end game year is touched. " + newline + newline;
                    }
                    else
                    {
                        str = "Please wait until the central player announce the game of next year." + newline + newline;
                    }
                    richTextBoxOEM.Invoke((MethodInvoker)delegate { richTextBoxOEM.Text += str; });
                    //backgroundWorker1.ReportProgress(0, str);
                    btnMdm.Invoke(new MethodInvoker(delegate { btnMdm.Enabled = true; }));
                    mdmflag = true;
                }

                // update running year
                if (runningfile.Length == 1)
                {
                    runningyear2 = Convert.ToInt32(runningfile[0].Substring(runningfile[0].Length - 8, 4));
                    if (runningyear2 > runningyear)
                    {
                        runningyear = runningyear2;
                        runningyearflag = false;
                        mdmflag = false;
                    }
                }


                //if (File.Exists(path + @"\announcement\EvaluationVolume_" + eyear + ".xls") && resetflag == false)
                if (!local)
                {
                    if (File.Exists(@"\\" + machineID + @"\" + sharedOem + @"\WG_input_" + eyear + ".xls") && resetflag == false)
                    {
                        // btnReset.Invoke(new MethodInvoker(delegate { btnReset.Enabled = true; }));
                        string sss = "The last year market demanding results are available: " + newline + path + @"\" + sharedOem + @"\WG_input_" + eyear + ".xls" + newline + newline;
                        sss += "Please go to the view area to see the results" + newline + newline;
                        richTextBoxOEM.Invoke((MethodInvoker)delegate { richTextBoxOEM.Text += sss; });
                        resetflag = true;
                    }
                    if (!File.Exists(@"\\" + machineID + @"\" + sharedOem + @"\WG_input_" + eyear + ".xls") && resetflag == true)
                    {
                        resetflag = false;
                    }
                }
                else
                {
                    // MessageBox.Show("eyear= " + eyear);
          
                    if (File.Exists("C:/WG/OEM/" + sharedOem + "/WG_input_" + eyear + ".xls") && resetflag == false)
                    {
                        MessageBox.Show("resetflag= " + resetflag);
                        btnReadResults.Invoke(new MethodInvoker(delegate { btnReadResults.Enabled = true; }));
                        string sss = "The last year market demanding results are available: " + newline + path + "/" + sharedOem + "/WG_input_" + eyear + ".xls" + newline + newline;
                        sss += "Please go to the view area to see the results" + newline + newline;
                        richTextBoxOEM.Invoke((MethodInvoker)delegate { richTextBoxOEM.Text += sss; });
                        resetflag = true;
                    }
                    if (!File.Exists("C:/WG/OEM/" + sharedOem + "/WG_input_" + eyear + ".xls") && resetflag == true)
                    {
                        MessageBox.Show("resetflag=" + resetflag);
                        resetflag = false;
                    }
                }
            
                /*
                rangefile = Directory.GetFiles(path + @"\announcement\", "running*.txt");
                if (rangefile.Count() == 1 && runningyearflag == false)
                {
                    runningyear = Convert.ToInt32(rangefile[0].Substring(rangefile[0].Length - 8, 4));
                    runningyearflag = true;
                    backgroundWorker1.ReportProgress(0, String.Format("Now we are playing year of " + runningyear.ToString()));
                    backgroundWorker1.ReportProgress(0, String.Format("You can review market results of last year"));
                    btnMdm.Invoke(new MethodInvoker(delegate { btnMdm.Enabled = true; }));
                }

                // update running year
                if (rangefile.Count() == 1)
                {
                    runningyear2 = Convert.ToInt32(rangefile[0].Substring(rangefile[0].Length - 8, 4));
                    if (runningyear2 > runningyear)
                    {
                        runningyearflag = false;
                        mdmflag = false;
                    }
                } */
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int Records = 0;

            if (Demo.SelectedIndex == 5)
            {
                // Add mothed for Client of OEM module
                set_connect_center();
            }
            
            if ( Demo.SelectedIndex == 6 )
            {
                string connetionString = null;
                // string OEM = selectOEM.SelectedItem.ToString();


                SqlConnection cnn;
                connetionString = "Data Source=Srl0ad79.srl.ford.com; Database = WarGaming; Integrated Security=SSPI;";
                cnn = new SqlConnection(connetionString);
                try
                {
                    cnn.Open();
                    // MessageBox.Show("Connection Open ! ");
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("Can not open sql server database! ");
                }

                string sSQL = "SELECT COUNT(*) FROM FORD_2016";

                try
                {
                    SqlCommand cmd = new SqlCommand(sSQL, cnn);
 
                    Records = int.Parse(cmd.ExecuteScalar().ToString());
                    if (Records > 0 )
                        selectOEM.Items.Add("FORD");
   
                }
                catch
                {
                    // MessageBox.Show("the table is not exist");
                }

                sSQL = "SELECT COUNT(*) FROM GM_2016";

                try
                {
                    SqlCommand cmd = new SqlCommand(sSQL, cnn);
                    Records = int.Parse(cmd.ExecuteScalar().ToString());
                    if (Records > 0)
                        selectOEM.Items.Add("GM");
                }
                catch
                {
                    // MessageBox.Show("the table is not exist");
                }

                sSQL = "SELECT COUNT(*) FROM FCA_2016";

                try
                {
                    // ANSI SQL way.  Works in PostgreSQL, MSSQL, MySQL.  
                    SqlCommand cmd = new SqlCommand(sSQL, cnn);
                    Records = int.Parse(cmd.ExecuteScalar().ToString());
                    if (Records > 0)
                        selectOEM.Items.Add("FCA");
                }
                catch
                {
                    // MessageBox.Show("the table is not exist");
                }

                sSQL = "SELECT COUNT(*) FROM HONDA_2016";

                try
                {
                    SqlCommand cmd = new SqlCommand(sSQL, cnn);
                    Records = int.Parse(cmd.ExecuteScalar().ToString());
                    if (Records > 0)
                        selectOEM.Items.Add("HONDA");
                }
                catch
                {
                    // MessageBox.Show("the table is not exist");
                }

                sSQL = "SELECT COUNT(*) FROM TOYOTA_2016";

                try
                {
                    SqlCommand cmd = new SqlCommand(sSQL, cnn);
                    Records = int.Parse(cmd.ExecuteScalar().ToString());
                    if (Records > 0)
                        selectOEM.Items.Add("TOYOTA");
                }
                catch
                {
                    // MessageBox.Show("the table is not exist");
                }


                sSQL = "SELECT COUNT(*) FROM NISSAN_2016";

                try
                {
                    SqlCommand cmd = new SqlCommand(sSQL, cnn);
                    Records = int.Parse(cmd.ExecuteScalar().ToString());
                    if (Records > 0 )
                         selectOEM.Items.Add("NISSAN");
                }
                catch
                {
                    // MessageBox.Show("the table is not exist");
                }

                sSQL = "SELECT COUNT(*) FROM BMW_2016";
                try
                {
                    SqlCommand cmd = new SqlCommand(sSQL, cnn);
                    Records = int.Parse(cmd.ExecuteScalar().ToString());
                    if (Records > 0)
                        selectOEM.Items.Add("BMW");
                }
                catch
                {
                    // MessageBox.Show("the table is not exist");
                }

                sSQL = "SELECT COUNT(*) FROM HYUNDAI_2016";
                try
                {
                    SqlCommand cmd = new SqlCommand(sSQL, cnn);
                    Records = int.Parse(cmd.ExecuteScalar().ToString());
                    if (Records > 0)
                        selectOEM.Items.Add("HYUNDAI");
                }
                catch
                {
                    // MessageBox.Show("the table is not exist");
                }

     

                cnn.Close();
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void Results_1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Range range;

            string str;
            int rCnt = 0;
            int cCnt = 0;

            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Cursor Files|*.xlsx";
            openFileDialog1.Title = "Select a Cursor File";

            // Show the Dialog.
            // If the user clicked OK in the dialog and
            // a .CUR file was selected, open it.
            String fileName = "";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileName = openFileDialog1.FileName;
            }

            // MessageBox.Show("Open Excel file name is: " + fileName);

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(fileName);
            Excel._Worksheet xlWorkSheet = xlWorkBook.Sheets[1];

            range = xlWorkSheet.UsedRange;


            for (rCnt = 1; rCnt < 4; rCnt++)
            {
                str = "";
                for (cCnt = 1; cCnt <= 6; cCnt++)
                {
                    str = str + "    " + (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;

                }
            }

            // xlWorkBook.Close(true, null, null);
            xlApp.Quit();
        }


        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void label24_Click(object sender, EventArgs e)
        {

        }

        private void label20_Click(object sender, EventArgs e)
        {

        }

        private void richTextBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This Function is underdevelopment");
        }


        private void comboBox13_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox14_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkBox13_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox14_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label21_Click(object sender, EventArgs e)
        {

        }

        private void label22_Click(object sender, EventArgs e)
        {

        }

        private void checkBox15_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void comboBox15_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label23_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            Excel.Range range;

            string str;
            int rCnt = 0;
            int cCnt = 0;

            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Cursor Files|*.xlsx";
            openFileDialog1.Title = "Select a Cursor File";

            // Show the Dialog.
            // If the user clicked OK in the dialog and
            // a .CUR file was selected, open it.
            String fileName = "";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileName = openFileDialog1.FileName;
            }

            // MessageBox.Show("Open Excel file, display the first row,  add 2 strings to row 4");
            // Output_2.Text = "The First Row will be display one column by one column";

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(fileName);
            Excel._Worksheet xlWorkSheet = xlWorkBook.Sheets[1];

            // write formula as text 
            xlWorkSheet.Cells[4, "A"] = "7*3+2";
            // write formula as formula
            xlWorkSheet.Cells[4, "B"] = "=7*3+2";
            range = xlWorkSheet.UsedRange;

            for (rCnt = 1; rCnt < 2; rCnt++)
            {
                for (cCnt = 1; cCnt <= 6; cCnt++)
                {
                    str = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                    MessageBox.Show(str);
                }
            }

            // xlWorkBook.Close(true, null, null);
            xlApp.Quit();

        }

        private void button5_Click(object sender, EventArgs e)
        {
            Excel.Range range;

            string str;
            int rCnt = 0;
            int cCnt = 0;

            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Cursor Files|*.xlsx";
            openFileDialog1.Title = "Select a Cursor File";

            // Show the Dialog.
            // If the user clicked OK in the dialog and
            // a .CUR file was selected, open it.
            String fileName = "";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileName = openFileDialog1.FileName;
            }

            // MessageBox.Show("Open Excel file, display the first row,  add 2 strings to row 4");
            EPA_Results.Text = "The First Row will be display one column by one column";

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(fileName);
            Excel._Worksheet xlWorkSheet = xlWorkBook.Sheets[1];

            // write formula as text 
            xlWorkSheet.Cells[4, "A"] = "7*3+2";
            // write formula as formula
            xlWorkSheet.Cells[4, "B"] = "=7*3+2";
            range = xlWorkSheet.UsedRange;

            for (rCnt = 1; rCnt < 2; rCnt++)
            {
                for (cCnt = 1; cCnt <= 6; cCnt++)
                {
                    str = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                    MessageBox.Show(str);
                }
            }

            // xlWorkBook.Close(true, null, null);
            xlApp.Quit();
        }

        private void button7_Click(object sender, EventArgs e)
        {

        }

        private void Process_II_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This Function is underdevelopment");

        }

        private void Process_III_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This Function is underdevelopment");
        }

        private void Process_IV_Click(object sender, EventArgs e)
        {

        }

        private void fileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // MessageBox.Show("This Function is underdevelopment");
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This Function is underdevelopment");
        }

        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This Function is underdevelopment");
        }

        private void openDatabaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This Function is underdevelopment");
        }

        private void closeDatabaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This Function is underdevelopment");
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
            //MessageBox.Show("This Function is underdevelopment");
        }

        private void runToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This Function is underdevelopment");
        }

        private void debugToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This Function is underdevelopment");
        }

        private void aboutUsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This Function is underdevelopment");
        }

        private void onlineHelpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("file:///C:/War_Gaming/War_Gaming/bin/html/index.html");
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This Function is underdevelopment");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string connetionString = null;
            SqlConnection cnn;
            connetionString = "Data Source=Srl0ad79.srl.ford.com; Database = carsdotcom; Integrated Security=SSPI;";
            cnn = new SqlConnection(connetionString);
            try
            {
                cnn.Open();
                MessageBox.Show("Connection Open ! ");

                cnn.Close();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Can not open connection ! ");
            }
        }

        public class Manufactures
        {
            private int RecID;
            public int recID
            {
                get { return RecID; }
                set { RecID = value; }
            }
            private string Mfg_Name;
            public string mfgName
            {
                get { return Mfg_Name; }
                set { Mfg_Name = value; }
            }
            private string WMI;
            public string wmi
            {
                get { return WMI; }
                set { WMI = value; }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'warGamingDataSet1.UserList' table. You can move, or remove it, as needed.
            this.userListTableAdapter2.Fill(this.warGamingDataSet1.UserList);
            // TODO: This line of code loads data into the 'warGamingDataSet2.UserList' table. You can move, or remove it, as needed.
            this.userListTableAdapter1.Fill(this.warGamingDataSet2.UserList);
            // TODO: This line of code loads data into the 'carsdotcomDataSet2.userList' table. You can move, or remove it, as needed.
            this.userListTableAdapter.Fill(this.carsdotcomDataSet2.userList);
            // TODO: This line of code loads data into the 'carsdotcomDataSet.Manufacturers_S' table. You can move, or remove it, as needed.
            this.manufacturers_STableAdapter.Fill(this.carsdotcomDataSet.Manufacturers_S);
            // TODO: This line of code loads data into the 'localDatabaseDataSet1.test1' table. You can move, or remove it, as needed.
            this.test1TableAdapter.Fill(this.localDatabaseDataSet1.test1);
            // TODO: This line of code loads data into the 'testDatabaseDataSet1.EmployeeInfo' table. You can move, or remove it, as needed.
            // this.employeeInfoTableAdapter.Fill(this.testDatabaseDataSet1.EmployeeInfo);

            // tabSVL.SelectedTab = tabIED; // JF 3/7/16

        }

        private void addNew_Click(object sender, EventArgs e)
        {
            this.employeeInfoBindingSource1.AddNew();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            // this.employeeInfoBindingSource1.AddNew();
            this.test1BindingSource.AddNew();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.test1BindingSource.EndEdit();
            // this.tableAdapterManager1.UpdateAll(this.testDatabaseDataSet1);
            this.tableAdapterManager.UpdateAll(this.testDatabaseDataSet1);
        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            this.test1BindingSource.RemoveCurrent();
        }

        private void label32_Click(object sender, EventArgs e)
        {

        }

        private void radioButton12_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void Done_1_Click(object sender, EventArgs e)
        {
            if (!player)
            {
                total++;
                gameDialog.AppendText("Step= " + total + "\n");
            }
            if (group1_1.Checked)
            {
                group1_value--;
                disable(1, group1_value);
                gameDialog.AppendText("Take 1 in Group 1 by player\n");
                Console.Beep();
                play();
            }
            else if (group1_2.Checked)
            {
                group1_value = group1_value - 2;
                disable(1, group1_value);
                gameDialog.AppendText("Take 2 in Group 1 by player\n");
                Console.Beep();
                Console.Beep();
                play();
            }
            else if (group1_3.Checked)
            {
                group1_value = group1_value - 3;
                disable(1, group1_value);
                gameDialog.AppendText("Take 3 in Group 1 by player\n");
                Console.Beep();
                Console.Beep();
                Console.Beep();
                play();
            }
            else if (group2_1.Checked)
            {
                group2_value--;
                disable(2, group2_value);
                gameDialog.AppendText("Take 1 in Group 2 by player\n");
                Console.Beep();
                play();
            }
            else if (group2_2.Checked)
            {
                group2_value = group2_value - 2;
                disable(2, group2_value);
                gameDialog.AppendText("Take 2 in Group 2 by player\n");
                Console.Beep();
                Console.Beep();
                play();
            }

            else if (group2_3.Checked)
            {
                group2_value = group2_value - 3;
                disable(2, group2_value);
                gameDialog.AppendText("Take 3 in Group 2 by player\n");
                Console.Beep();
                Console.Beep();
                Console.Beep();
                play();
            }

            else if (group2_4.Checked)
            {
                group2_value = group2_value - 4;
                disable(2, group2_value);
                gameDialog.AppendText("Take 4 in Group 2 by player\n");
                Console.Beep();
                Console.Beep();
                Console.Beep();
                Console.Beep();
                play();
            }

            else if (group2_5.Checked)
            {
                group2_value = group2_value - 5;
                disable(2, group2_value);
                gameDialog.AppendText("Take 5 in Group 2 by player\n");
                Console.Beep();
                Console.Beep();
                Console.Beep();
                Console.Beep();
                Console.Beep();
                play();
            }

            else if (group3_1.Checked)
            {
                group3_value = group3_value - 1;
                disable(3, group3_value);
                gameDialog.AppendText("Take 1 in Group 3 by player\n");
                Console.Beep();
                play();
            }
            else if (group3_2.Checked)
            {
                group3_value = group3_value - 2;
                disable(3, group3_value);
                gameDialog.AppendText("Take 2 in Group 3 by player\n");
                Console.Beep();
                Console.Beep();
                play();
            }

            else if (group3_3.Checked)
            {
                group3_value = group3_value - 3;
                disable(3, group3_value);
                gameDialog.AppendText("Take 3 in Group 3 by player\n");
                Console.Beep();
                Console.Beep();
                Console.Beep();
                play();
            }

            else if (group3_4.Checked)
            {
                group3_value = group3_value - 4;
                disable(3, group3_value);
                gameDialog.AppendText("Take 4 in Group 3 by player\n");
                Console.Beep();
                Console.Beep();
                Console.Beep();
                Console.Beep();
                play();
            }
            else if (group3_5.Checked)
            {
                group3_value = group3_value - 5;
                disable(3, group3_value);
                gameDialog.AppendText("Take 5 in Group 3 by player\n");
                Console.Beep();
                Console.Beep();
                Console.Beep();
                Console.Beep();
                Console.Beep();
                play();
            }
        }

        private void Start_Click(object sender, EventArgs e)
        {
            group1_value = 3;
            group2_value = 5;
            group3_value = 7;
            group1_pic.Image = Image.FromFile("./3balls-s.jpg");
            group2_pic.Image = Image.FromFile("./5balls-s.jpg");
            group3_pic.Image = Image.FromFile("./7balls-s.jpg");
            group1_1.Enabled = true;
            group1_2.Enabled = true;
            group1_3.Enabled = true;
            group2_1.Enabled = true;
            group2_2.Enabled = true;
            group2_3.Enabled = true;
            group2_4.Enabled = true;
            group2_5.Enabled = true;
            group3_1.Enabled = true;
            group3_2.Enabled = true;
            group3_3.Enabled = true;
            group3_4.Enabled = true;
            group3_5.Enabled = true;
            total = 0;
            player = pcFirst.Checked;
            gameDialog.Text = "Start, please select\n";
            if (player)
            {
                play();
            }
            Execute.Enabled = true;
        }

        void disable(int group, int ind)
        {
            if (group == 1)
            {
                if (ind == 2)
                {
                    group1_3.Enabled = false;
                    group1_pic.Image = Image.FromFile("./2balls-s.jpg");
                }
                else if (ind == 1)
                {
                    group1_3.Enabled = false;
                    group1_2.Enabled = false;
                    group1_pic.Image = Image.FromFile("./ball-s.jpg");
                }
                else if (ind == 0)
                {
                    group1_3.Enabled = false;
                    group1_2.Enabled = false;
                    group1_1.Enabled = false;
                    group1_pic.Image = Image.FromFile("./0ball-s.jpg");
                }
            }
            else if (group == 2)
            {

                if (ind == 4)
                {
                    group2_5.Enabled = false;
                    group2_pic.Image = Image.FromFile("./4balls-s.jpg");
                }
                else if (ind == 3)
                {
                    group2_4.Enabled = false;
                    group2_5.Enabled = false;
                    group2_pic.Image = Image.FromFile("./3balls-s.jpg");
                }
                else if (ind == 2)
                {
                    group2_3.Enabled = false;
                    group2_4.Enabled = false;
                    group2_5.Enabled = false;
                    group2_pic.Image = Image.FromFile("./2balls-s.jpg");
                }
                else if (ind == 1)
                {
                    group2_2.Enabled = false;
                    group2_3.Enabled = false;
                    group2_4.Enabled = false;
                    group2_5.Enabled = false;
                    group2_pic.Image = Image.FromFile("./ball-s.jpg");
                }
                else if (ind == 0)
                {
                    group2_5.Enabled = false;
                    group2_4.Enabled = false;
                    group2_3.Enabled = false;
                    group2_2.Enabled = false;
                    group2_1.Enabled = false;
                    group2_pic.Image = Image.FromFile("./0ball-s.jpg");
                }
            }

            else if (group == 3)
            {
                if (ind == 6)
                {
                    group3_pic.Image = Image.FromFile("./6balls-s.jpg");
                }
                else if (ind == 5)
                {
                    group3_pic.Image = Image.FromFile("./5balls-s.jpg");
                }
                else if (ind == 4)
                {
                    group3_5.Enabled = false;
                    group3_pic.Image = Image.FromFile("./4balls-s.jpg");
                }
                else if (ind == 3)
                {
                    group3_4.Enabled = false;
                    group3_5.Enabled = false;
                    group3_pic.Image = Image.FromFile("./3balls-s.jpg");
                }
                else if (ind == 2)
                {
                    group3_3.Enabled = false;
                    group3_4.Enabled = false;
                    group3_5.Enabled = false;
                    group3_pic.Image = Image.FromFile("./2balls-s.jpg");
                }
                else if (ind == 1)
                {
                    group3_2.Enabled = false;
                    group3_3.Enabled = false;
                    group3_4.Enabled = false;
                    group3_5.Enabled = false;
                    group3_pic.Image = Image.FromFile("./ball-s.jpg");
                }
                else if (ind == 0)
                {
                    group3_5.Enabled = false;
                    group3_4.Enabled = false;
                    group3_3.Enabled = false;
                    group3_2.Enabled = false;
                    group3_1.Enabled = false;
                    group3_pic.Image = Image.FromFile("./0ball-s.jpg");
                }
            }

        }

        void play()
        {
            int num = group1_value * 100 + group2_value * 10 + group3_value;
            if (player)
            {
                total++;
                gameDialog.AppendText("Step= " + total + "\n");
            }

            if (num == 357)
            {
                System.DateTime moment = DateTime.UtcNow;
                int second = moment.Second;
                int mod = second % 3 + 1;

                if (mod == 1)
                {
                    group1_value = group1_value - 1;
                    disable(1, group1_value);
                    gameDialog.AppendText("Take 1 from Group 1 by computer\n");
                }
                else if (mod == 2)
                {
                    group2_value = 2;
                    disable(2, group2_value);
                    gameDialog.AppendText("Take 3 from Group 2 by computer\n");
                }
                else if (mod == 3)
                {
                    group3_value--;
                    disable(3, group3_value);
                    gameDialog.AppendText("Take 1 from Group 3 by computer\n");
                }

            }
            else if (group1_value + group2_value + group3_value == 1)
            {
                MessageBox.Show("Congrauation! You Win.");
                Execute.Enabled = false;
            }

            // (1,3, >3) -> (1,3,2) 
            else if (num >= 134 && num <= 137 || num >= 313 && num <= 317)
            {
                group3_value = 2;
                disable(3, group3_value);
                gameDialog.AppendText(" Left 2 in Group 3 by computer\n");
            }
            // (1,>=3,3) -> (1,2,3) 
            else if (num == 153 || num == 143 || num == 133)
            {
                group2_value = 2;
                disable(2, group2_value);
                gameDialog.AppendText(" Left 2 in Group 2 by computer\n");
            }

            // (1,5,>5) -> (1,5,4) 
            else if (num == 156 || num == 157)
            {
                group3_value = 4;
                disable(3, group3_value);
                gameDialog.AppendText(" Left 4 in Group 3 by computer\n");
            }

            // (1,5,>4) -> (1,5,4)
            else if (num > 154 && num <= 157)
            {
                group3_value = 4;
                disable(3, group3_value);
                gameDialog.AppendText(" Left 4 in Group 3 by computer\n");
            }

            // (1,0,1) -> (0,0,1) and (1,1,0) -> (0,1,0)
            else if (num == 101 || num == 110)
            {
                group1_value = 0;
                disable(1, group1_value);
                gameDialog.AppendText("Left 0 in Group 1 by computer\n");
                MessageBox.Show("Compter Win!");
                Execute.Enabled = false;
            }

            // (0,1,1) -> (0,0,1)
            else if (num == 11)
            {
                group2_value = 0;
                disable(2, group2_value);
                gameDialog.AppendText("Left 0 in Group 2 by computer\n");
                MessageBox.Show("Compter Win!");
                Execute.Enabled = false;

            }

            // (>1,1,1) -> (1,1,1) 
            else if (group1_value > 1 && group2_value == 1 && group3_value == 1)
            {
                group1_value = 1;
                disable(1, group1_value);
                gameDialog.AppendText("Left l in Group 1 by computer\n");
            }

            // (1,1,>1) -> (1,1,1) 
            else if (num > 111 && num <= 116)
            {
                group3_value = 1;
                disable(3, group3_value);
                gameDialog.AppendText("Left l in Group 3 by computer\n");
            }
            // (1,>1,1) -> (1,1,1) 
            else if (group2_value > 1 && group1_value == 1 && group3_value == 1)
            {
                group2_value = 1;
                disable(2, group2_value);
                gameDialog.AppendText("Left l in Group 2 by computer\n");
            }

            // (0,0,>1) -> (0,0,1) 
            else if (num > 1 && num <= 6)
            {
                group3_value = 1;
                disable(3, group3_value);
                gameDialog.AppendText("Left l in Group 3 by computer\n");
                MessageBox.Show("Computer Win!");
                Execute.Enabled = false;

            }
            // (0,>1,0) -> (0,1,0) 
            else if (group1_value + group3_value == 0 && group2_value > 1)
            {
                group2_value = 1;
                disable(2, group2_value);
                gameDialog.AppendText("Left l in Group 2 by computer\n");
                MessageBox.Show("Computer Win!");
                Execute.Enabled = false;

            }
            // (>1,0,0) -> (1,0,0) 
            else if (group2_value + group3_value == 0 && group1_value > 1)
            {
                group1_value = 1;
                disable(1, group1_value);
                gameDialog.AppendText("Left l in Group 1 by computer\n");
                MessageBox.Show("Computer Win!");
                Execute.Enabled = false;

            }
            // (0, 1, 7)  -> (0,1,6)  or  (1, 0, 7)  -> (1,0,6)  
            else if (group1_value + group2_value == 1 && group3_value == 7)
            {
                group3_value = 6;
                disable(3, group3_value);
                gameDialog.AppendText("Left 6 in Group 3 by computer\n");
            }
            // (0, 1, 6)  -> (0,1,5)  or  (1, 0, 7)  -> (0,1,6)
            else if (group1_value + group2_value == 1 && group3_value == 6)
            {
                group3_value = 5;
                disable(3, group3_value);
                gameDialog.AppendText("Left 5 in Group 3 by computer\n");
            }

            // (0, 1, >1)  -> (0,1,0) or  (1, 0, >1)  -> (1,0,0) 
            else if (group1_value + group2_value == 1 && group3_value > 1)
            {
                group3_value = 0;
                disable(3, group3_value);
                gameDialog.AppendText("Take all in Group 3 by computer\n");
                MessageBox.Show("Computer Win!");
                Execute.Enabled = false;
            }

            else if (group1_value + group3_value == 1 && group2_value > 1)
            {
                group2_value = 0;
                disable(2, group2_value);
                gameDialog.AppendText("Take all in Group 2 by computer\n");
                MessageBox.Show("Computer Win!");
                Execute.Enabled = false;

            }
            else if (group2_value + group3_value == 1 && group1_value > 1)
            {
                group1_value = 0;
                disable(1, group1_value);
                gameDialog.AppendText("Take all in Group 1 by computer\n");
                MessageBox.Show("Computer Win!");
                Execute.Enabled = false;
            }

            else if (num > 111 && num <= 116)
            {
                group3_value = 1;
                disable(3, group3_value);
                gameDialog.AppendText("Left l in Group 3 by computer\n");
            }
            //  (1,1,7) -> (0,1,7)
            else if (num == 117)
            {
                group1_value = 0;
                disable(1, group1_value);
                gameDialog.AppendText("Left 0 in Group 1 by computer\n");
            }
            //  (0,1,7) -> (0,0,7)
            else if (num == 17)
            {
                group2_value = 0;
                disable(2, group2_value);
                gameDialog.AppendText("Left 0 in Group 2 by computer\n");
            }
            //  (1,0,7) -> (0,0,7)
            else if (num == 107)
            {
                group1_value = 0;
                disable(1, group1_value);
                gameDialog.AppendText("Left 0 in Group 1 by computer\n");
            }
            //  (0,0,7) -> (0,0,6)
            else if (num == 7)
            {
                group3_value = 6;
                disable(3, group3_value);
                gameDialog.AppendText("Take 1 in Group 3 by computer\n");
            }
            //  (0,0,>1) -> (0,0,1)
            else if (num <= 6)
            {
                group3_value = 1;
                disable(3, group3_value);
                gameDialog.AppendText("Left 1 in Group 3 by computer\n");
                MessageBox.Show("Computer Win!");
                Execute.Enabled = false;
            }
            else if (num == 126)
            {
                group3_value = 3;
                disable(3, group3_value);
                gameDialog.AppendText("Take 3 in Group 3 by computer\n");
            }
            else if (num == 141)
            {
                group2_value = 2;
                disable(2, group2_value);
                gameDialog.AppendText("Take 2 in Group 2 by computer\n");
            }
            else if (num == 142)
            {
                group2_value = 3;
                disable(2, group2_value);
                gameDialog.AppendText("Take 2 in Group 2 by computer\n");
            }
            else if (num == 143)
            {
                group2_value = 2;
                disable(2, group2_value);
                gameDialog.AppendText("Take 2 in Group 2 by computer\n");
            }

            else if (num == 145)
            {
                group2_value = 3;
                disable(2, group2_value);
                gameDialog.AppendText("Take 1 in Group 2 by computer\n");
            }
            else if (num == 146 || num == 147)
            {
                group3_value = 5;
                disable(3, group3_value);
                gameDialog.AppendText("Left 5 in Group 3 by computer\n");
            }
            else if (num == 152)
            {
                group2_value = 3;
                disable(2, group2_value);
                gameDialog.AppendText("Take 2 in Group 2 by computer\n");
            }

            else if (num == 227 || num == 226)
            {
                group3_value--;
                disable(3, group3_value);
                gameDialog.AppendText("Take 1 in Group 3 by computer\n");
            }
            // (2,3,6) -> (2, 3, 1)
            else if (num == 236)
            {
                group3_value = 1;
                disable(3, group3_value);
                gameDialog.AppendText("Take 5 in Group 3 by computer\n");
            }
            // (2,>=4,1) -> (2, 3, 1)
            else if (num == 241 || num == 251)
            {
                group2_value = 3;
                disable(2, group2_value);
                gameDialog.AppendText("Left 3 in Group 2 by computer\n");
            }
            else if (num == 245)
            {
                group1_value = 1;
                disable(1, group1_value);
                gameDialog.AppendText("Take 1 in Group 1 by computer\n");
            }
            // (2,4,6) -> (2, 2, 6)
            else if (num == 256 || num == 246)
            {
                group2_value = 2;
                disable(2, group2_value);
                gameDialog.AppendText("Left 2 in Group 2 by computer\n");
            }
            // (2,>3,3) -> (2, 1, 3)
            else if (num == 243 || num == 253)
            {
                group2_value = 1;
                disable(2, group2_value);
                gameDialog.AppendText("Left 1 in Group 2 by computer\n");
            }
            // (2,5,4) -> (1, 5, 4)
            else if (num == 254)
            {
                group1_value = 1;
                disable(1, group1_value);
                gameDialog.AppendText("Take 1 in Group 1 by computer\n");
            }
            // (2,5,7) -> (2, 3, 7)
            else if (num == 257)
            {
                group2_value = 3;
                disable(2, group2_value);
                gameDialog.AppendText("Take 2 in Group 2 by computer\n");
            }
            // (3,2,6) -> (3, 2, 1)
            else if (num == 326)
            {
                group3_value = 1;
                disable(3, group3_value);
                gameDialog.AppendText("Left 1 in Group 3 by computer\n");
            }
            // (3,3,7) -> (3, 3, 6)
            else if (num == 336 || num == 337)
            {
                group3_value--;
                disable(3, group3_value);
                gameDialog.AppendText("Take 1 in Group 3 by computer\n");
            }
            // (3,4,5) -> (1, 4, 5)
            else if (num == 345)
            {
                group1_value = 1;
                disable(1, group1_value);
                gameDialog.AppendText("Take 2 in Group 1 by computer\n");
            }
            // (3,4,6) -> (3, 3, 6)
            else if (num == 346)
            {
                group2_value = 3;
                disable(2, group2_value);
                gameDialog.AppendText("Take 1 in Group 2 by computer\n");
            }
            // (3,4,7) -> (3, 2, 7)
            else if (num == 347 || num == 337)
            {
                group2_value = 2;
                disable(2, group2_value);
                gameDialog.AppendText("Left 2 in Group 2 by computer\n");
            }
            // (3,5,1) -> (3,2,1)
            else if (num == 351)
            {
                group2_value = 2;
                disable(2, group2_value);
                gameDialog.AppendText("Take 3 in Group 2 by computer\n");
            }
            // (3,5,2) -> (3,1,2)
            else if (num == 352)
            {
                group2_value = 1;
                disable(2, group2_value);
                gameDialog.AppendText("Left 1 in Group 2 by computer\n");
            }
            // (3,5,4) -> (1,5,4)
            else if (num == 354)
            {
                group1_value = 1;
                disable(1, group1_value);
                gameDialog.AppendText("Take 2 in Group 1 by computer\n");
            }
            else if (group1_value > 1 && group2_value > 1 && group3_value == 0 && group1_value != group2_value)
            {
                if (group2_value > group1_value)
                {
                    group2_value = group1_value;
                    disable(2, group2_value);
                    gameDialog.AppendText("Make group1 and group2 equal by Computer\n");
                }
                else if (group2_value < group1_value)
                {
                    group1_value = group2_value;
                    disable(1, group1_value);
                    gameDialog.AppendText("Make group1 and group2 equal by Computer\n");
                }
            }

            else if (group1_value > 1 && group3_value > 1 && group2_value == 0 && group1_value != group3_value)
            {
                if (group1_value > group3_value)
                {
                    group1_value = group3_value;
                    disable(1, group1_value);
                    gameDialog.AppendText("Make group1 and group3 equal by Computer\n");
                }
                else if (group1_value < group3_value)
                {
                    group3_value = group1_value;
                    disable(3, group3_value);
                    gameDialog.AppendText("Make group1 and group3 equal by Computer\n");
                }
            }

            else if (group2_value > 1 && group3_value > 1 && group1_value == 0 && group2_value != group3_value)
            {
                if (group2_value > group3_value)
                {
                    group2_value = group3_value;
                    disable(2, group2_value);
                    gameDialog.AppendText("Make group2 and group3 equal by Computer\n");
                }
                else if (group2_value < group3_value)
                {
                    group3_value = group2_value;
                    disable(3, group3_value);
                    gameDialog.AppendText("Make group2 and group3 equal by Computer\n");
                }
            }

            // two groups have equal number, and another is 0
            else if (group1_value == group2_value && group3_value == 0 || group1_value == group3_value && group2_value == 0)
            {
                group1_value--;
                disable(1, group1_value);
                gameDialog.AppendText("Left 1 in Group 1 by computer\n");
            }
            else if (group2_value == group3_value && group1_value == 0)
            {
                group2_value--;
                disable(2, group2_value);
                gameDialog.AppendText("Left 1 in Group 2 by computer\n");
            }

            // two groups have equal number, and another is larger 0
            else if (group1_value == group2_value && group3_value * group1_value > group3_value && group3_value < 6)
            {
                group3_value = 0;
                disable(3, group3_value);
                gameDialog.AppendText(" Take all in Group 3 by computer\n");
            }
            // two groups have equal number, and another is larger 0
            else if (group1_value == group3_value && group2_value * group1_value > group2_value)
            {
                group2_value = 0;
                disable(2, group2_value);
                gameDialog.AppendText("Take all in Group 2 by computer\n");
            }

            // two groups have equal number, and another is larger 0
            else if (group2_value == group3_value && group2_value * group1_value > group1_value)
            {
                group1_value = 0;
                disable(1, group1_value);
                gameDialog.AppendText("Take all in Group 1 by computer\n");
            }

            // the groups is (2,4,7) -> (2,4,6)
            else if (group1_value == 2 && group2_value == 4 && group3_value * group1_value > 6)
            {
                group3_value = 6;
                disable(3, group3_value);
                gameDialog.AppendText("Take 1 in Group 3 by computer\n");
            }

            // other condition
            else if (num > 1 && group1_value > 0)
            {
                group1_value--;
                disable(1, group1_value);
                gameDialog.AppendText("Take 1 in Group 1 by computer\n");
            }

            else if (num > 1 && group2_value > 0)
            {
                group2_value--;
                disable(2, group2_value);
                gameDialog.AppendText("Take 1 in Group 2 by computer\n");
            }
            else if (num > 1 && group3_value > 0)
            {
                group3_value--;
                disable(3, group3_value);
                gameDialog.AppendText("Take 1 in Group 3 by computer\n");
            }

        }

        private void pictureBox9_Click(object sender, EventArgs e)
        {

        }

        private void gotoWeb_Click(object sender, EventArgs e)
        {
            webBrowser1.Navigate(webAddress.Text);
        }

        private void emailMeToolStripMenuItem_Click(object sender, EventArgs e)
        {

            Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("Software\\microsoft\\windows\\currentversion\\app paths\\OUTLOOK.EXE");
            string path = (string)key.GetValue("Path");
            if (path != null)
            {
                // System.Diagnostics.Process.Start("OUTLOOK.EXE");
                System.Diagnostics.Process.Start("OUTLOOK.EXE", "/c ipm.note /m jfang3@ford.com&subject=WarGaming");
            }
            else
                MessageBox.Show("There is no Outlook in this computer!", "SystemError", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            /*
            try
            {
                Form4 Email = new Form2();
                Email.Show();
            }
            catch (System.Exception s)
            {

            }
            */

        }

        private void goHome_Click(object sender, EventArgs e)
        {
            webBrowser1.Navigate("https://team.sp.ford.com/sites/PDSA/WarGaming");
        }

        private void Foward_Click(object sender, EventArgs e)
        {
            webBrowser1.GoForward();
        }

        private void Back_Click(object sender, EventArgs e)
        {
            webBrowser1.GoBack();
        }

        private void Refresh_Click(object sender, EventArgs e)
        {
            webBrowser1.Refresh();
        }

        private void goToolStripMenuItem_Click(object sender, EventArgs e)
        {
            webBrowser1.Navigate(webAddress.Text);
        }

        private void homeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            webBrowser1.Navigate("https://team.sp.ford.com/sites/PDSA/WarGaming");
        }

        private void fowardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            webBrowser1.GoForward();
        }

        private void backToolStripMenuItem_Click(object sender, EventArgs e)
        {
            webBrowser1.GoBack();
        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            webBrowser1.Refresh();
        }

        private void helpToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }


        private void butDLLini_Click(object sender, EventArgs e)
        {

        }

        private void button12_Click(object sender, EventArgs e)
        {
            EPA_Results.Text = "";
            // clearTextBox();
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // DLL init (to create objects in DLL)
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            string strNameplateData = Nameplate.Text;
            string strEPACreditReport = EPA_2014.Text;
            string strNHTSACreditReport = NHTSA.Text;

            int nSizeNameplateData = strNameplateData.Length;
            int nSizeEPACreditReport = strEPACreditReport.Length;
            int nSizeNHTSACreditReport = strNHTSACreditReport.Length;
            bool bRtn = dll_module_init(strNameplateData, nSizeNameplateData, strEPACreditReport, nSizeEPACreditReport, strNHTSACreditReport, nSizeNHTSACreditReport);

            if (bRtn)
                EPA_Results.AppendText("DLL Initialized\n");
            else
                EPA_Results.AppendText("DLL Initialization Failed\n");


            // DataTable dt = getDataTableFromCSVFile("C:\\Users\\JFANG3\\Documents\\Visual Studio 2012\\WarGamingGUI-v1\\WG_const_dll_test_20160210\\OEM_Vol_test.csv");
            DataTable dt = getDataTableFromCSVFile(textOEM.Text);

            Dictionary<string, int[]> dict = new Dictionary<string, int[]>();
            Random rnd1 = new Random();
            foreach (DataRow row in dt.Rows)
            {
                string strNameplate = row["Nameplate"].ToString();
                int[] vol = Enumerable.Repeat(0, 32).ToArray();
                vol[0] = Int32.Parse(row["2015"].ToString());
                vol[1] = vol[0] * rnd1.Next(95, 105) / 100;
                vol[2] = vol[0] * rnd1.Next(95, 105) / 100;
                vol[3] = vol[0] * rnd1.Next(95, 105) / 100;
                vol[4] = vol[0] * rnd1.Next(95, 105) / 100;
                vol[5] = vol[0] * rnd1.Next(95, 105) / 100;
                vol[6] = vol[0] * rnd1.Next(95, 105) / 100;
                vol[7] = vol[0] * rnd1.Next(95, 105) / 100;
                vol[8] = vol[0] * rnd1.Next(95, 105) / 100;
                vol[9] = vol[0] * rnd1.Next(95, 105) / 100;
                vol[10] = vol[0] * rnd1.Next(95, 105) / 100;
                dict.Add(strNameplate, vol);
            }



            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // OEM init (to tell DLL for OEM to pull out a spectific OEM dataset
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            string strOEM = "FORD";
            int nSizeOEM = strOEM.Length;
            dll_oem_init(strOEM, nSizeOEM, 2015, 2025);

            var sw = Stopwatch.StartNew();
            double ttLHS = .0;
            for (int i = 0; i < 1000; i++)
            {
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                // OEM volume push
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                foreach (KeyValuePair<string, int[]> entry in dict)
                {
                    int nArray = 11;
                    int random = rnd1.Next(0, 10);
                    int orgVol = entry.Value[random];
                    entry.Value[random] = entry.Value[random] * rnd1.Next(95, 105) / 100;
                    string strNameplate = entry.Key.ToString();
                    dll_oem_push_vol(strOEM, nSizeOEM, strNameplate, strNameplate.Length, entry.Value, nArray);
                }


                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                // get the evaluation value (LHS): positive negative in your bank account over the time horizon
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                double lhs = dll_get_EPA_const_val();
                ttLHS += lhs;


            }

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Once OEM optimization gets an optimal volume for the particular year, update the dll
            // so that the dll remembers the updated volume and manage EPA & NHTSA credit banks
            // then we move to 2016 optimization....
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            foreach (KeyValuePair<string, int[]> entry in dict)
            {
                string strNameplate = entry.Key.ToString();
                dll_oem_finalize_vol(strOEM, nSizeOEM, strNameplate, strNameplate.Length, entry.Value[0], 2015);
            }

            sw.Stop();
            TimeSpan tmSpan = sw.Elapsed;
            // pushText("1000 evaluation function calls: " + tmSpan.TotalSeconds.ToString() + " (sec)");
            // pushText("This is avg. LHS val = " + (ttLHS/1000).ToString());
            EPA_Results.AppendText("1000 evaluation function calls: " + tmSpan.TotalSeconds.ToString() + " (sec)\n");
            EPA_Results.AppendText("This is LHS val = " + (ttLHS / 1000).ToString() + "\n");


            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // DLL delete (to release all memory allocated in DLL)
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            dll_module_destroy();

            // EPA_Results.AppendText("DLL Release from Memory\n");

            // save credit in array creditResults
            creditResults[0, 0] = 2015;
            creditResults[0, 1] = 1000;
            creditResults[1, 0] = 2016;
            creditResults[1, 1] = 1100;

            creditResults[2, 0] = 2017;
            creditResults[2, 1] = 1200;
            creditResults[3, 0] = 2018;
            creditResults[3, 1] = 1150;

            creditResults[4, 0] = 2019;
            creditResults[4, 1] = 1300;
            creditResults[5, 0] = 2020;
            creditResults[5, 1] = 1400;

            creditResults[6, 0] = 2021;
            creditResults[6, 1] = 1250;
            creditResults[7, 0] = 2022;
            creditResults[7, 1] = 1450;

            creditResults[8, 0] = 2023;
            creditResults[8, 1] = 1750;
            creditResults[9, 0] = 2024;
            creditResults[9, 1] = 1650;
        }

        private static DataTable getDataTableFromCSVFile(string strFilePath)
        {
            DataTable dt = new DataTable();
            var fs = new FileStream(strFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using (StreamReader sr = new StreamReader(fs))
            {
                string[] hdrs = sr.ReadLine().Split(',');
                foreach (string hdr in hdrs)
                {
                    dt.Columns.Add(hdr);
                }
                while (!sr.EndOfStream)
                {
                    string[] rows = sr.ReadLine().Split(',');
                    DataRow dr = dt.NewRow();
                    for (int i = 0; i < hdrs.Length; i++)
                    {
                        dr[i] = rows[i];
                    }
                    dt.Rows.Add(dr);
                }
            }
            return dt;
        }

        private void button13_Click(object sender, EventArgs e)
        {

        }

        private void butOEM_Click(object sender, EventArgs e)
        {

            OpenFileDialog openFileDialog2 = new OpenFileDialog();
            openFileDialog2.Filter = "Cursor Files|*.csv";
            openFileDialog2.Title = "Select a CSV File";

            // Show the Dialog.
            // If the user clicked OK in the dialog and
            // a .CUR file was selected, open it.
            String fileName = "";
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                fileName = openFileDialog1.FileName;
            }
            NHTSA.Text = fileName;


        }

        private void butNameplate_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog2 = new OpenFileDialog();
            openFileDialog2.Filter = "Cursor Files|*.csv";
            openFileDialog2.Title = "Select a CSV File";

            // Show the Dialog.
            // If the user clicked OK in the dialog and
            // a .CUR file was selected, open it.
            String fileName = "";
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                fileName = openFileDialog2.FileName;
            }
            Nameplate.Text = fileName;
        }

        private void butEPA_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog2 = new OpenFileDialog();
            openFileDialog2.Filter = "Cursor Files|*.csv";
            openFileDialog2.Title = "Select a CSV File";

            // Show the Dialog.
            // If the user clicked OK in the dialog and
            // a .CUR file was selected, open it.
            String fileName = "";
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                fileName = openFileDialog2.FileName;
            }
            EPA_2014.Text = fileName;
        }

        private void butSelectFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                textChoose.Text = openFileDialog1.FileName;
                butLoadXslx.Enabled = true;

            }
        }

        private void butLoadXslx_Click(object sender, EventArgs e)
        {
            
            // Excel.ApplicationClass ExcelObj = new Excel.ApplicationClass();
            Microsoft.Office.Interop.Excel.Application ExcelObj = new Microsoft.Office.Interop.Excel.Application();
            // Excel.Workbook theWorkbook = ExcelObj.Workbooks.Open(textChoose.Text);

            Excel.Workbook theWorkbook = ExcelObj.Workbooks.Open(textChoose.Text, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);

            Excel.Sheets sheets = theWorkbook.Worksheets;

            Excel.Worksheet worksheet = (Excel.Worksheet)sheets.get_Item(1);

            Excel.Range range = worksheet.UsedRange;

            System.Array myvalues = (System.Array)range.Cells.Value2;

            int vertical = myvalues.GetLength(0);
            int horizontal = myvalues.GetLength(1);

            string[] headers = new string[horizontal];
            string[] data = new string[horizontal];

            DataTable ResultsHeader = new DataTable();
            DataSet ds = new DataSet();


            for (int x = 1; x <= vertical; x++)
            {
                // Utils.inicializarArrays(datos);
                for (int y = 1; y <= horizontal; y++)
                {
                    if (x == 1)
                    {
                        headers[y - 1] = myvalues.GetValue(x, y).ToString();
                        // MessageBox.Show(headers[y-1]);
                    }
                    else
                    {
                        string auxdata = "";
                        if (myvalues.GetValue(x, y) != null)
                            auxdata = myvalues.GetValue(x, y).ToString();
                        data[y - 1] = auxdata;
                        // MessageBox.Show(data[y-1]);
                    }

                }
                if (x == 1) //headers
                {
                    for (int w = 0; w < horizontal; w++)
                    {
                        // ResultsHeader.Columns.Add(new DataColumn(headers[w], GetType()));
                        ResultsHeader.Columns.Add(new DataColumn(headers[w], typeof(string)));
                    }
                    ds.Tables.Add(ResultsHeader);
                }
                else
                {
                    DataRow dataRow = ds.Tables[0].NewRow();
                    for (int w = 0; w < horizontal; w++)
                    {
                        // MessageBox.Show(headers[w]);
                        // MessageBox.Show(data[w]);

                        dataRow[headers[w]] = data[w];
                    }
                    ds.Tables[0].Rows.Add(dataRow);
                }
            }
            DataView myDataView = new DataView();
            myDataView.Table = ds.Tables[0];
            // dataGridView3.CurrentPageIndex = 0;
            dataGridView3.DataSource = myDataView;
            // dataGridView3.DataBind();
            // dataGridView3.Dispose();
            theWorkbook.Close();
            ExcelObj.Quit(); 
       
            
        }

        private void butPlot_Click(object sender, EventArgs e)
        {



            foreach (var series in chartCredit.Series)
            {
                series.Points.Clear();
            }

            for (int i = 0; i < plotNum.Value; i++)
            {
                chartCredit.Series["Credit"].Points.Add(creditResults[i, 1]);
                chartCredit.Series["Credit"].Points[i].Color = Color.Blue;
                chartCredit.Series["Credit"].Points[i].AxisLabel = creditResults[i, 0].ToString();
                chartCredit.Series["Credit"].Points[i].LegendText = creditResults[i, 0].ToString(); ;
                chartCredit.Series["Credit"].Points[i].Label = creditResults[i, 1].ToString();
            }

        }

        private void chartCredit_Click(object sender, EventArgs e)
        {

        }


        private void loadVars(string brand)
        {
            // chklstESSVars.Items.Clear();

            int idx;
            vdata1 = new ArrayList();
            genPopu1 = new ArrayList();
            genfcons1 = new ArrayList();
            fordconstraints1 = new ArrayList();
            basedelta1 = new ArrayList();
            refresh1 = new ArrayList();
            genvehdata1 = new ArrayList();
            fmccdata1 = new ArrayList();

            /*
            idx = Array.IndexOf(Routines.xlsSheets, "vehdata");
            vdata1.Add("vehdata price"); vdata1.Add("vehdata volume");
            if(yr==Routines.baseYear) for (int i = 0; i < vdata1.Count; i++) chklstESSVars.Items.Add(vdata1[i]);
            */

            if (brand == "Ford")
            {
                idx = Array.IndexOf(Routines.xlsSheets, "macro");
                genPopu1.Add("genpopu population");
                genPopu1.Add("genpopu cpi");
                genPopu1.Add("genpopu gasprice");
                genPopu1.Add("genpopu income");
                genPopu1.Add("genpopu carcafestd");
                genPopu1.Add("genpopu truckcafestd");
                genPopu1.Add("genpopu totadbudget");
                genPopu1.Add("genpopu dealadbudget");
                //for (int i = 0; i < genPopu1.Count; i++) chklstESSVars.Items.Add(genPopu1[i]);

                idx = Array.IndexOf(Routines.xlsSheets, "fordconstraints");
                for (int i = 0; i < Routines.shtNames[idx].Length; i++) genfcons1.Add("genfcons " + Routines.shtNames[idx][i]);
                genfcons1.Remove("genfcons oem"); genfcons1.Remove("genfcons name"); genfcons1.Remove("genfcons transyear");
                genfcons1.Remove("genfcons rentalelast"); genfcons1.Remove("genfcons rentalvol0");
                genfcons1.Remove("genfcons rentalriskelast"); genfcons1.Remove("genfcons rentalriskvol0");
                genfcons1.Remove("genfcons fleetelast"); genfcons1.Remove("genfcons fleetvol0");
                // for (int i = 0; i < genfcons1.Count; i++) chklstESSVars.Items.Add(genfcons1[i]);

                idx = Array.IndexOf(Routines.xlsSheets, "basedelta");
                for (int i = 0; i < Routines.shtNames[idx].Length; i++) basedelta1.Add("basedelta " + Routines.shtNames[idx][i]);
                basedelta1.Remove("basedelta oem"); basedelta1.Remove("basedelta name"); basedelta1.Remove("basedelta type");
                basedelta1.Remove("basedelta segment"); basedelta1.Remove("basedelta mgroupid"); basedelta1.Remove("basedelta tgroupid"); basedelta1.Remove("basedelta sgroupid");
                // basedelta1.Remove("basedelta modelyear");
                //for (int i = 0; i < basedelta1.Count; i++) chklstESSVars.Items.Add(basedelta1[i]);

                idx = Array.IndexOf(Routines.xlsSheets, "refresh");
                refresh1.Add("refresh successrate"); refresh1.Add("refresh increasedvolrate");
                // for (int i = 0; i < refresh1.Count; i++) chklstESSVars.Items.Add(refresh1[i]);

                /*   //David changed o n 7/24/06
               for (int i = 0; i < Routines.shtNames[idx].Length; i++) fmccdata1.Add("fmccdata "+Routines.shtNames[idx][i]);
               fmccdata1.Remove("fmccdata oem"); fmccdata1.Remove("fmccdata name");
               if (yr == Routines.baseYear) for (int i = 0; i < fmccdata1.Count; i++) chklstESSVars.Items.Add(fmccdata1[i]);
               */

            }
            else
            {
                idx = Array.IndexOf(Routines.xlsSheets, "vehdata");
                //genvehdata1.Add("vehdata price"); 
                genvehdata1.Add("genvehdata volume");
                //for (int i = 0; i < Routines.shtNames[idx].Length; i++) genvehdata1.Add("genvehdata " + Routines.shtNames[idx][i]);
                //genvehdata1.Remove("genvehdata oem"); genvehdata1.Remove("genvehdata name"); genvehdata1.Remove("genvehdata type");
                //genvehdata1.Remove("genvehdata transyear"); genvehdata1.Remove("genvehdata segment"); genvehdata1.Remove("genvehdata modelyear");
                //genvehdata1.Remove("genvehdata sgroupid"); genvehdata1.Remove("genvehdata mgroupid");
                //  for (int i = 0; i < genvehdata1.Count; i++) chklstESSVars.Items.Add(genvehdata1[i]);
            }

        }

        private void loadNewVars(string brand)
        {
            // chklstESSVars.Items.Clear();

            int idx;
            vdata1 = new ArrayList();
            genPopu1 = new ArrayList();
            genfcons1 = new ArrayList();
            fordconstraints1 = new ArrayList();
            basedelta1 = new ArrayList();
            refresh1 = new ArrayList();
            genvehdata1 = new ArrayList();
            fmccdata1 = new ArrayList();

            /*
            idx = Array.IndexOf(Routines.xlsSheets, "vehdata");
            vdata1.Add("vehdata price"); vdata1.Add("vehdata volume");
            if(yr==Routines.baseYear) for (int i = 0; i < vdata1.Count; i++) chklstESSVars.Items.Add(vdata1[i]);
            */

            if (brand == "Ford")
            {


                idx = Array.IndexOf(Routines.xlsSheets, "fordconstraints");
                for (int i = 0; i < Routines.shtNames[idx].Length; i++) genfcons1.Add("genfcons " + Routines.shtNames[idx][i]);
                genfcons1.Remove("genfcons oem"); genfcons1.Remove("genfcons name"); genfcons1.Remove("genfcons transyear");
                // for (int i = 0; i < genfcons1.Count; i++) chklstESSVars.Items.Add(genfcons1[i]);


                idx = Array.IndexOf(Routines.xlsSheets, "basedelta");
                for (int i = 0; i < Routines.shtNames[idx].Length; i++) basedelta1.Add("basedelta " + Routines.shtNames[idx][i]);
                basedelta1.Remove("basedelta oem"); basedelta1.Remove("basedelta name"); basedelta1.Remove("basedelta type");
                // for (int i = 0; i < basedelta1.Count; i++) chklstESSVars.Items.Add(basedelta1[i]);



                /*   //David changed o n 7/24/06
               for (int i = 0; i < Routines.shtNames[idx].Length; i++) fmccdata1.Add("fmccdata "+Routines.shtNames[idx][i]);
               fmccdata1.Remove("fmccdata oem"); fmccdata1.Remove("fmccdata name");
               if (yr == Routines.baseYear) for (int i = 0; i < fmccdata1.Count; i++) chklstESSVars.Items.Add(fmccdata1[i]);
               */

            }
            else
            {
                idx = Array.IndexOf(Routines.xlsSheets, "vehdata");
                //genvehdata1.Add("vehdata price"); 
                genvehdata1.Add("genvehdata volume");
                //for (int i = 0; i < Routines.shtNames[idx].Length; i++) genvehdata1.Add("genvehdata " + Routines.shtNames[idx][i]);
                //genvehdata1.Remove("genvehdata oem"); genvehdata1.Remove("genvehdata name"); genvehdata1.Remove("genvehdata type");
                //genvehdata1.Remove("genvehdata transyear"); genvehdata1.Remove("genvehdata segment"); genvehdata1.Remove("genvehdata modelyear");
                //genvehdata1.Remove("genvehdata sgroupid"); genvehdata1.Remove("genvehdata mgroupid");
                // for (int i = 0; i < genvehdata1.Count; i++) chklstESSVars.Items.Add(genvehdata1[i]);
            }

        }


        private void btnRALoad_Click(object sender, EventArgs e)
        {
            if (Routines.population == null || Routines.genVehData == null) return;

            //decision = dfRoutines();
            //write to excel
            //.....
            string dataDir = Directory.GetCurrentDirectory();
            string dataFile = dataDir + "\\EditDec.xls";
            if (scenario < 4)
            {

                if (File.Exists(dataFile))
                    Routines.viewEXCEL(dataFile);
                else
                {
                    MessageBox.Show("Decision does not exit!");
                    return;
                }
            }
            else
            {
                dataFile = dataDir + "\\EditPrice.xls";

                if (File.Exists(dataFile))
                    Routines.viewEXCEL(dataFile);
                else
                {
                    MessageBox.Show("Decision does not exit!");
                    return;
                }
            }
        }


        private void butNHTSA_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Cursor Files|*.csv";
            openFileDialog1.Title = "Select a CSV File";

            // Show the Dialog.
            // If the user clicked OK in the dialog and
            // a .CUR file was selected, open it.
            String fileName = "";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileName = openFileDialog1.FileName;
            }
            NHTSA.Text = fileName;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog2 = new OpenFileDialog();
            openFileDialog2.Filter = "Cursor Files|*.csv";
            openFileDialog2.Title = "Select a CSV File";

            // Show the Dialog.
            // If the user clicked OK in the dialog and
            // a .CUR file was selected, open it.
            String fileName = "";
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                fileName = openFileDialog2.FileName;
            }
            textOEM.Text = fileName;
        }

        private void saveRec_Click(object sender, EventArgs e)
        {

        }

        private void removeRec_Click(object sender, EventArgs e)
        {

        }

        private void userListDataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void buSaveDatabase_Click(object sender, EventArgs e)
        {
            string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            // MessageBox.Show("You are: " + userName );
            string connetionString = null;
            SqlConnection cnn;
            connetionString = "Data Source=Srl0ad79.srl.ford.com; Database = WarGaming; Integrated Security=SSPI;";
            cnn = new SqlConnection(connetionString);
            try
            {
                cnn.Open();
                // MessageBox.Show("Connection Open ! ");
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Can not open sql server database! ");
            }

            cnn.Close();
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void butLoadData_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            // Display Qurry Form on another Form
            DatabaseQurry form3 = new DatabaseQurry();
            form3.Show();
        }

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }

        private void butNHTSAOEM_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog2 = new OpenFileDialog();
            openFileDialog2.Filter = "Cursor Files|*.csv";
            openFileDialog2.Title = "Select a CSV File";

            // Show the Dialog.
            // If the user clicked OK in the dialog and
            // a .CUR file was selected, open it.
            String fileName = "";
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                fileName = openFileDialog2.FileName;
            }
            textBoxOEM.Text = fileName;
        }

        private void buttNHTSA_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Cursor Files|*.csv";
            openFileDialog1.Title = "Select a CSV File";

            // Show the Dialog.
            // If the user clicked OK in the dialog and
            // a .CUR file was selected, open it.
            String fileName = "";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileName = openFileDialog1.FileName;
            }
            textBoxNHTSA.Text = fileName;
        }

        private void buttNHSTANameplate_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog2 = new OpenFileDialog();
            openFileDialog2.Filter = "Cursor Files|*.csv";
            openFileDialog2.Title = "Select a CSV File";

            // Show the Dialog.
            // If the user clicked OK in the dialog and
            // a .CUR file was selected, open it.
            String fileName = "";
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                fileName = openFileDialog2.FileName;
            }
            textBoxNameplate.Text = fileName;
        }

        private void buttNHSTAEPA_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog2 = new OpenFileDialog();
            openFileDialog2.Filter = "Cursor Files|*.csv";
            openFileDialog2.Title = "Select a CSV File";

            // Show the Dialog.
            // If the user clicked OK in the dialog and
            // a .CUR file was selected, open it.
            String fileName = "";
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                fileName = openFileDialog2.FileName;
            }
            textBoxEPA.Text = fileName;
        }

        private void button9_Click_1(object sender, EventArgs e)
        {
            NHTSAResults.Text = "";
            // clearTextBox();
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // DLL init (to create objects in DLL)
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            string strNameplateData = textBoxNameplate.Text;
            string strEPACreditReport = textBoxEPA.Text;
            string strNHTSACreditReport = textBoxNHTSA.Text;

            int nSizeNameplateData = strNameplateData.Length;
            int nSizeEPACreditReport = strEPACreditReport.Length;
            int nSizeNHTSACreditReport = strNHTSACreditReport.Length;
            bool bRtn = dll_module_init(strNameplateData, nSizeNameplateData, strEPACreditReport, nSizeEPACreditReport, strNHTSACreditReport, nSizeNHTSACreditReport);

            if (bRtn)
                NHTSAResults.AppendText("DLL Initialized\n");
            else
                NHTSAResults.AppendText("DLL Initialization Failed\n");
        }

        private void button8_Click_2(object sender, EventArgs e)
        {
            NHTSAResults.Text = "";
            // clearTextBox();
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // DLL init (to create objects in DLL)
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            string strNameplateData = textBoxNameplate.Text;
            string strEPACreditReport = textBoxEPA.Text;
            string strNHTSACreditReport = textBoxNHTSA.Text;

            int nSizeNameplateData = strNameplateData.Length;
            int nSizeEPACreditReport = strEPACreditReport.Length;
            int nSizeNHTSACreditReport = strNHTSACreditReport.Length;
            bool bRtn = dll_module_init(strNameplateData, nSizeNameplateData, strEPACreditReport, nSizeEPACreditReport, strNHTSACreditReport, nSizeNHTSACreditReport);

            if (bRtn)
                NHTSAResults.AppendText("DLL Initialized\n");
            else
                NHTSAResults.AppendText("DLL Initialization Failed\n");


            // DataTable dt = getDataTableFromCSVFile("C:\\Users\\JFANG3\\Documents\\Visual Studio 2012\\WarGamingGUI-v1\\WG_const_dll_test_20160210\\OEM_Vol_test.csv");
            DataTable dt = getDataTableFromCSVFile(textBoxOEM.Text);

            Dictionary<string, int[]> dict = new Dictionary<string, int[]>();
            Random rnd1 = new Random();
            foreach (DataRow row in dt.Rows)
            {
                string strNameplate = row["Nameplate"].ToString();
                int[] vol = Enumerable.Repeat(0, 32).ToArray();
                vol[0] = Int32.Parse(row["2015"].ToString());
                vol[1] = vol[0] * rnd1.Next(95, 105) / 100;
                vol[2] = vol[0] * rnd1.Next(95, 105) / 100;
                vol[3] = vol[0] * rnd1.Next(95, 105) / 100;
                vol[4] = vol[0] * rnd1.Next(95, 105) / 100;
                vol[5] = vol[0] * rnd1.Next(95, 105) / 100;
                vol[6] = vol[0] * rnd1.Next(95, 105) / 100;
                vol[7] = vol[0] * rnd1.Next(95, 105) / 100;
                vol[8] = vol[0] * rnd1.Next(95, 105) / 100;
                vol[9] = vol[0] * rnd1.Next(95, 105) / 100;
                vol[10] = vol[0] * rnd1.Next(95, 105) / 100;
                dict.Add(strNameplate, vol);
            }



            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // OEM init (to tell DLL for OEM to pull out a spectific OEM dataset
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            string strOEM = "FORD";
            int nSizeOEM = strOEM.Length;
            dll_oem_init(strOEM, nSizeOEM, 2015, 2025);

            var sw = Stopwatch.StartNew();
            double ttLHS = .0;
            for (int i = 0; i < 1000; i++)
            {
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                // OEM volume push
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                foreach (KeyValuePair<string, int[]> entry in dict)
                {
                    int nArray = 11;
                    int random = rnd1.Next(0, 10);
                    int orgVol = entry.Value[random];
                    entry.Value[random] = entry.Value[random] * rnd1.Next(95, 105) / 100;
                    string strNameplate = entry.Key.ToString();
                    dll_oem_push_vol(strOEM, nSizeOEM, strNameplate, strNameplate.Length, entry.Value, nArray);
                }


                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                // get the evaluation value (LHS): positive negative in your bank account over the time horizon
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                double lhs = dll_get_EPA_const_val();
                ttLHS += lhs;
            }

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Once OEM optimization gets an optimal volume for the particular year, update the dll
            // so that the dll remembers the updated volume and manage EPA & NHTSA credit banks
            // then we move to 2016 optimization....
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            foreach (KeyValuePair<string, int[]> entry in dict)
            {
                string strNameplate = entry.Key.ToString();
                dll_oem_finalize_vol(strOEM, nSizeOEM, strNameplate, strNameplate.Length, entry.Value[0], 2015);
            }

            sw.Stop();
            TimeSpan tmSpan = sw.Elapsed;
            // pushText("1000 evaluation function calls: " + tmSpan.TotalSeconds.ToString() + " (sec)");
            // pushText("This is avg. LHS val = " + (ttLHS/1000).ToString());
            NHTSAResults.AppendText("1000 evaluation function calls: " + tmSpan.TotalSeconds.ToString() + " (sec)\n");
            NHTSAResults.AppendText("This is LHS val = " + (ttLHS / 1000).ToString() + "\n");

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // DLL delete (to release all memory allocated in DLL)
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            dll_module_destroy();

        }

        private void button7_Click_1(object sender, EventArgs e)
        {
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // DLL delete (to release all memory allocated in DLL)
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            dll_module_destroy();

            NHTSAResults.AppendText("DLL Release from Memory\n");
        }

        private void button3_Click_1(object sender, EventArgs e)
        {

            // chartCreditNHTSA.ChartAreas[0].Area3DStyle.Enable3D = true;
            // save credit in array CarCredit
            CarCredit[0, 0] = 2009;
            CarCredit[0, 1] = 10123456;
            CarCredit[1, 0] = 2010;
            CarCredit[1, 1] = 8526532;

            CarCredit[2, 0] = 2011;
            CarCredit[2, 1] = 14029628;
            CarCredit[3, 0] = 2012;
            CarCredit[3, 1] = 14029796;

            CarCredit[4, 0] = 2013;
            CarCredit[4, 1] = 24503730;
            CarCredit[5, 0] = 2014;
            CarCredit[5, 1] = 28524032;

            CarCredit[6, 0] = 2015;
            CarCredit[6, 1] = -28293176;
            CarCredit[7, 0] = 2016;
            CarCredit[7, 1] = 24354888;

            CarCredit[8, 0] = 2017;
            CarCredit[8, 1] = 19458396;
            CarCredit[9, 0] = 2018;
            CarCredit[9, 1] = 15958110;

            CarCredit[10, 0] = 2019;
            CarCredit[10, 1] = -12767919;
            CarCredit[11, 0] = 2020;
            CarCredit[11, 1] = 29241544;

            CarCredit[12, 0] = 2021;
            CarCredit[12, 1] = 7143628;
            CarCredit[13, 0] = 2022;
            CarCredit[13, 1] = 7143628;

            CarCredit[14, 0] = 2023;
            CarCredit[14, 1] = 15958110;
            CarCredit[15, 0] = 2024;
            CarCredit[15, 1] = 15958110;

            CarCredit[16, 0] = 2025;
            CarCredit[16, 1] = 0;

            // save credit in array TruckCredit
            TruckCredit[0, 0] = 2009;
            TruckCredit[0, 1] = 8526532;
            TruckCredit[1, 0] = 2010;
            TruckCredit[1, 1] = 10925536;

            TruckCredit[2, 0] = 2011;
            TruckCredit[2, 1] = -12488768;
            TruckCredit[3, 0] = 2012;
            TruckCredit[3, 1] = -13531410;

            TruckCredit[4, 0] = 2013;
            TruckCredit[4, 1] = 15820439;
            TruckCredit[5, 0] = 2014;
            TruckCredit[5, 1] = 31020040;

            TruckCredit[6, 0] = 2015;
            TruckCredit[6, 1] = 10000065;
            TruckCredit[7, 0] = 2016;
            TruckCredit[7, 1] = 13364307;

            TruckCredit[8, 0] = 2017;
            TruckCredit[8, 1] = -28293176;
            TruckCredit[9, 0] = 2018;
            TruckCredit[9, 1] = 0;

            TruckCredit[10, 0] = 2019;
            TruckCredit[10, 1] = 29506208;
            TruckCredit[11, 0] = 2020;
            TruckCredit[11, 1] = 7143628;

            TruckCredit[12, 0] = 2021;
            TruckCredit[12, 1] = 15958110;
            TruckCredit[13, 0] = 2022;
            TruckCredit[13, 1] = 29506208;

            TruckCredit[14, 0] = 2023;
            TruckCredit[14, 1] = -28293176;
            TruckCredit[15, 0] = 2024;
            TruckCredit[15, 1] = 38859732;

            TruckCredit[16, 0] = 2025;
            TruckCredit[16, 1] = 18637576;


            foreach (var series in chartCreditNHTSA.Series)
            {
                series.Points.Clear();
            }


            for (int i = 0; i < plotNumYear.Value; i++)
            {
                chartCreditNHTSA.Series["Car Credit"].Points.Add(CarCredit[i, 1]);
                chartCreditNHTSA.Series["Car Credit"].Points[i].Color = Color.Blue;
                chartCreditNHTSA.Series["Car Credit"].Points[i].AxisLabel = CarCredit[i, 0].ToString();
                chartCreditNHTSA.Series["Car Credit"].Points[i].LegendText = CarCredit[i, 0].ToString(); ;
                chartCreditNHTSA.Series["Car Credit"].Points[i].Label = CarCredit[i, 1].ToString();

                chartCreditNHTSA.Series["Truck Credit"].Points.Add(TruckCredit[i, 1]);
                chartCreditNHTSA.Series["Truck Credit"].Points[i].Color = Color.Red;
                chartCreditNHTSA.Series["Truck Credit"].Points[i].AxisLabel = TruckCredit[i, 0].ToString();
                chartCreditNHTSA.Series["Truck Credit"].Points[i].LegendText = TruckCredit[i, 0].ToString(); ;
                chartCreditNHTSA.Series["Truck Credit"].Points[i].Label = TruckCredit[i, 1].ToString();
            }


        }

        private void check3D_CheckedChanged(object sender, EventArgs e)
        {
            if (check3D.Checked)
                chartCreditNHTSA.ChartAreas[0].Area3DStyle.Enable3D = true;
            else
                chartCreditNHTSA.ChartAreas[0].Area3DStyle.Enable3D = false;
        }

        private void checkEPA3D_CheckedChanged(object sender, EventArgs e)
        {
            if (checkEPA3D.Checked)
                chartCredit.ChartAreas[0].Area3DStyle.Enable3D = true;
            else
                chartCredit.ChartAreas[0].Area3DStyle.Enable3D = false;
        }

        private void textChoose_TextChanged(object sender, EventArgs e)
        {

        }

        private void butReset_Click(object sender, EventArgs e)
        {

            double margin, range, percent;

            margin = Convert.ToDouble(zEVSafetyMargin.Text);
            range = Convert.ToDouble(zEVRange.Text);
            percent = Convert.ToDouble(zEVSaleRange.Text);

            for (int i = 0; i < dataGridView3.RowCount - 1; i++)
            {
                dataGridView3.Rows[i].Cells["ZEV_States_Percentage"].Value = percent / 100;
                dataGridView3.Rows[i].Cells["2015 Range"].Value = range;
            }
            dataGridView3.Refresh();
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            Routines.timeHorizon = (int)numericUpDown1.Value;
            Routines.oeshare = new double[Routines.timeHorizon][];
            Routines.sgshare = new double[Routines.timeHorizon][];
            Routines.fushare = new double[Routines.timeHorizon][];
            Routines.tshare = new double[Routines.timeHorizon][];
            Routines.lshare = new double[Routines.timeHorizon][];
            Routines.bshare = new double[Routines.timeHorizon][];
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            Routines.beginYear = (int)BeginYear.Value;
        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {
            Routines.baseYear = (int)numericUpDown3.Value;
        }

        private void comboBoxOEM_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBoxOEM.Text = comboBoxOEM.SelectedItem.ToString();
        }

        private void butSelectInput_Click(object sender, EventArgs e)
        {
            txtInputFile.Text = "";
            Regex r = new Regex("(\\\\)");
            openImportFile.Title = "Select a excel input data file";
            openImportFile.InitialDirectory = Directory.GetCurrentDirectory();
            openImportFile.FileName = "Inputdata.xls";
            openImportFile.Filter = "excel files (*.xls)|*.xls";
            if (openImportFile.ShowDialog() == DialogResult.OK)
            {
                string str = Directory.GetCurrentDirectory() + "\\" + this.openImportFile.FileName;
                string[] tmp = r.Split(str);
                importFile = tmp[tmp.Length - 1];
                txtInputFile.Text = this.openImportFile.FileName;
            }
            // groupBox2.Enabled = false;
            groupBox3.Enabled = false;
            // button3.Enabled = false;
            // button4.Enabled = false;
            // button5.Enabled = false;
            OptimizationResults.Text = "";
            // radioButton1.Checked = false;
            // radioButton2.Checked = false;
        }

        private void butImport_Click(object sender, EventArgs e)
        {
            Routines.importFile(importFile, 1);
            BeginYear.Value = Routines.recentYear + 1;
            numericUpDown3.Value = Routines.recentYear;
            textImport.Text = "Imported";
            // groupBox2.Enabled = true;
            groupBox3.Enabled = true;
            comboBoxOEM.Items.Clear();
            comboBoxOEM.Items.Add("FORD");
            comboBoxOEM.Items.Add("GMC");
            comboBoxOEM.Items.Add("FIAT");
            comboBoxOEM.Items.Add("TOYOTA");
            comboBoxOEM.Items.Add("HONDA");
            comboBoxOEM.Items.Add("NISSAN");
            comboBoxOEM.Items.Add("HYUNDAI");
            comboBoxOEM.Items.Add("VOLKSWAGEN");
            comboBoxOEM.Items.Add("MAZDA");
            comboBoxOEM.Items.Add("VOLVO");
            comboBoxOEM.Items.Add("BMW");
            comboBoxOEM.Items.Add("DAIMLER");

        }

        private void butOptimization_Click(object sender, EventArgs e)
        {

            Routines.timeHorizon = Convert.ToInt32(numericUpDown1.Value);
            Routines.beginYear = Convert.ToInt32(BeginYear.Value);
            Routines.baseYear = Convert.ToInt32(numericUpDown3.Value);
            Routines.optOEM = comboBoxOEM.Text.ToString().ToLower();
            GlobeData.BeginYear = BeginYear.Value.ToString();
            GlobeData.OEM = comboBoxOEM.Text.ToString().ToUpper();


            if (Routines.baseYear > Routines.recentYear)
            {
                MessageBox.Show("No Historical Data of " + Routines.baseYear);
                return;
            }
            Szroutine szFun = new Szroutine();
            Routines.genPopu = new Macro[1];
            Routines.genPopu = szFun.getPopu(Routines.population);
            Routines.genVehData = new VehicleData[1][];

            Routines.genVehData = szFun.prefutureData(Routines.vData, Routines.timeHorizon, Routines.Scenario);
            szFun.CalMktShare(Routines.genVehData);

            Routines.optData = new VehicleData[0][];
            Routines.optData = szFun.preOptData(Routines.optOEM, Routines.timeHorizon, Routines.genVehData);
            Routines.listOpt = new ArrayList[0];
            Routines.listOpt = szFun.getOptIndexlist(Routines.optData);
            Routines.genFCons = new Constraint[0];
            Routines.genFCons = szFun.getOptConstraint();

            Routines.listNewFCons = new ArrayList();
            Routines.listNewVData = new listnewvdata[Routines.timeHorizon];
            for (int i = 0; i < Routines.timeHorizon; i++)  //index for
            {
                Routines.listNewVData[i] = new listnewvdata();
                Routines.listNewVData[i].yr = Routines.beginYear + i;
                Routines.listNewVData[i].idx = new ArrayList();
                for (int j = 0; j < Routines.genVehData[i].Length; j++)
                    Routines.listNewVData[i].idx.Add(Routines.genVehData[i][j].modelName.ToLower() +
                      Routines.genVehData[i][j].transYear);
            }
            for (int i = 0; i < Routines.genFCons.Length; i++)
                Routines.listNewFCons.Add(Routines.genFCons[i].ModelName.ToLower() + Routines.genFCons[i].ModelYear);

            //add for listBDelat by sz
            Routines.listBDelta = new ArrayList();
            for (int i = 0; i < Routines.baseDelta.Length; i++)
            {
                Routines.listBDelta.Add(Routines.baseDelta[i].modelName.ToLower());
            }

            //  Routines.optiType = 2;

            //      genVarInd(Routines.timeHorizon, Routines.optData);
            genNormalFactor(Routines.optData);

            Optimization dfFun = new Optimization();
            string NAGResult = dfFun.OptimalRun();
            OptimizationResults.Text = NAGResult;

            string file = Directory.GetCurrentDirectory() + "\\OptimalResult" + ".xls";
            /* DialogResult rs = MessageBox.Show("Do you want to open Excel to analyse profit",
                                              "Evaluation Price", MessageBoxButtons.YesNo,
                                              MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
            if (rs == DialogResult.Yes) Routines.viewEXCEL(file);  JF 5/16/16  */

            // MessageBox.Show("Popup new window");
            // Display Optimization on another Form
            OptimizationDemo form3 = new OptimizationDemo();
            form3.Show();
        }

        private void genNormalFactor(VehicleData[][] vehNameTM)
        {
            int optLength = 0;
            for (int i = 0; i < vehNameTM.Length; i++)
                optLength += vehNameTM[i].Length;
            double[] baseVol = new double[optLength];     //base year volume by vehicle

            int transyear = Routines.baseYear + Routines.timeHorizon;
            int id = 0;
            for (int i = 0; i < vehNameTM.Length; i++)         //get selected Ford vehicle list.
            {
                for (int j = 0; j < vehNameTM[i].Length; j++)
                {
                    string oem = vehNameTM[i][j].OEM;
                    string modelName = vehNameTM[i][j].modelName;


                    //Get volume and ad spendings from base data
                    //  double[] vehBase = new double[];

                    int idx1 = Routines.listNewVData[Routines.timeHorizon - 1].idx.IndexOf(modelName + transyear);

                    if (idx1 > -1)
                    {
                        baseVol[id] = Routines.genVehData[Routines.timeHorizon - 1][idx1].volume;

                    }
                    else
                        baseVol[id] = 1000.0;
                    id++;
                }

            }
            Routines.baseVol = baseVol;
        }

        private void butOEMQuery_Click(object sender, EventArgs e)
        {
            string connetionString = null;
            string OEM = selectOEM.SelectedItem.ToString();
            // MessageBox.Show("OEM= " + OEM);

            SqlConnection cnn;
            connetionString = "Data Source=Srl0ad79.srl.ford.com; Database = WarGaming; Integrated Security=SSPI;";
            cnn = new SqlConnection(connetionString);
            try
            {
                cnn.Open();
                // MessageBox.Show("Connection Open ! ");
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Can not open sql server database! ");
            }

            string sSQL = "select * from " + OEM + "_2016";


            SqlCommand cmd = new SqlCommand(sSQL, cnn);

            // cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();
            SqlDataAdapter sqlDataAdap = new SqlDataAdapter(cmd);
            sqlDataAdap.Fill(dt);
            dataGridViewOEM.DataSource = dt;
        }

        private void dataGridViewOEM_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void selectOEM_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void userListDataGridView_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox4_Enter_1(object sender, EventArgs e)
        {

        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {
            /*
            // set the current caret position to the end
            richTextBoxOEM.SelectionStart = richTextBoxOEM.Text.Length;
            // scroll it automatically
            richTextBoxOEM.ScrollToCaret();
            //HighlightWords();
             */
        }

        private void groupBox7_Enter(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void IELoad_Click(object sender, EventArgs e)
        {
            if (runningyear == 0)
            {
                MessageBox.Show("The central player has not announced a game. Please wait ...");
                return;
            }

            string[] runningfile = Directory.GetFiles(path + @"\announcement\", "running*.txt");
            if (runningfile.Length == 1)
                runningyear = Convert.ToInt32(runningfile[0].Substring(runningfile[0].Length - 8, 4));

            // Cursor.Current = Cursors.WaitCursor;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            if ( !local )
                importFile = @"\\" + machineID + @"\" + sharedOem + @"\WG_input_" + (runningyear - 1).ToString() + ".xls";
            else
                importFile = "C:/WG/OEM/" + sharedOem + "/WG_input_" + (runningyear - 1).ToString() + ".xls";
            // MessageBox.Show(importFile);           
            Routines.importFile(importFile, 2);

            richTextBoxOEM.Text += "Input file " + "WG_input_" + (runningyear - 1).ToString() + ".xls is loaded." + newline + newline;
            Routines.timeHorizon = 1;
            //Routines.beginYear = Routines.recentYear + 1;
            //Routines.baseYear = Routines.recentYear;
            GlobeData.BeginYear = runningyear.ToString();   //JF
            Routines.beginYear = runningyear;
            Routines.baseYear = runningyear - 1;


            //tabSVL.SelectedIndex = 0;
            btnSVL.Enabled = true;
            if (btnReadResults.Enabled) btnReadResults.Enabled = false;
            btnToMDM.Enabled = false;
            btnROOptimal.Enabled = false;
            butChart.Enabled = false;
        }

        // private void btnReset_Click(object sender, EventArgs e)
        private void btnReadResults_Click(object sender, EventArgs e)
        {
            /*
            IELoad.Enabled = false;
            btnSelectOem.Enabled = true;
            //listBox1.Enabled = true;
            btnSVL.Enabled = false;
            btnROOptimal.Enabled = false;
            dataGridView1.Rows.Clear();
            //backgroundWorker1.CancelAsync();
            btnSelectOem.BackgroundImage = null;
            btnSelectOem.Text = "Click to select OEM";
            btnToMDM.Enabled = false;
            try
            {
                var dirannouncement = new DirectoryInfo(path + @"\" + sharedOem);
                foreach (var file in dirannouncement.EnumerateFiles("*.*")) { file.Delete(); }
            }
            catch
            {
                richTextBoxOEM.Text += "Reset failed" + Environment.NewLine;
            } */

 
            // runningyear = Routines.beginYear;
            MessageBox.Show("runningyear=" + runningyear);
            // MessageBox.Show("sharedOem=" + sharedOem);

            // runningyear = 2017;
            // sharedOem = "FORD";

            string file1;
            if (!local)
                file1 = @"\\" + machineID + @"\" + sharedOem + @"\WG_input_" + (runningyear).ToString() + ".xls";
            else
                file1 = "C:/WG/OEM/" + sharedOem + "/WG_input_" + (runningyear).ToString() + ".xls";

            if (!File.Exists(file1))
            {
                MessageBox.Show("The MDM result for " + (runningyear).ToString() + " has not been published.");
                return;
            }
            string file2 = Directory.GetCurrentDirectory() + "\\WG_input_" + (runningyear).ToString() + ".xls";
            if (File.Exists(file2)) File.Delete(file2);
            File.Copy(file1, file2);

            DialogResult rs = MessageBox.Show("Do you want to open Excel to analyse profit",
                                  "Evaluation Price", MessageBoxButtons.YesNo,
                                  MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
            if (rs == DialogResult.Yes) Routines.viewEXCEL(file2);
        }

        private void btnROOptimal_Click(object sender, EventArgs e)
        {
            string file = Directory.GetCurrentDirectory() + "\\OptimalResult.xls";
            int nrow, ncol;
            int i;
            string[] strHeader = new string[9];
            strHeader[0] = "Group"; strHeader[1] = "Type";
            strHeader[2] = "Fuel_Type"; strHeader[3] = "Segment";
            strHeader[4] = "Oem"; strHeader[5] = "Model"; strHeader[6] = "year";
            strHeader[7] = "price"; strHeader[8] = "volume";

            // get row and colume numbers
            ncol = dataGridView1.Columns.Count - 2; //3rd to 11st
            nrow = 0;
            foreach (DataGridViewRow dgvRow in dataGridView1.Rows)
            {
                // Make sure it's not an empty row.
                if (!dgvRow.IsNewRow) nrow++;
            }

            object[,] objR = new object[nrow, strHeader.Length];
            i = 0;
            foreach (DataGridViewRow dgvRow in dataGridView1.Rows)
            {
                // Make sure it's not an empty row.
                if (!dgvRow.IsNewRow)
                {
                    for (int c = 2; c < dgvRow.Cells.Count; c++)
                    {
                        // Append the cells data followed by a comma to delimit.
                        objR[i, c - 2] = dgvRow.Cells[c].Value;
                    }
                    i++;
                }
            }
            Routines.WriteToEXCEL(objR, strHeader, 1, "optData", file);
            this.Text = "you change on optimization results has been saved";
            richTextBoxOEM.Text += "The content in " + file + " is updated with your change\r\n";
        }

        private void btnToMDM_Click(object sender, EventArgs e)
        {
            string newfile;
            if (!local)
                newfile = @"\\" + machineID + @"\" + sharedOem + @"\" + selectedOem.ToLower() + "_" + runningyear.ToString() + ".xls";
            else
                newfile = "C:/WG/OEM/" + sharedOem + "/" + selectedOem.ToLower() + "_" + runningyear.ToString() + ".xls";
            
            if (File.Exists(newfile))
            {
                string sss = "You have already submitted this year's optimization results to central player" + "\r\n" + "\r\n";
                sss += "Please don't submit it again" + "\r\n" + "\r\n";
                MessageBox.Show(sss);
                return;
            }

            if (sharedOem == null) return;

            // the selected OEM has not been claimed
            DialogResult dialogResult = MessageBox.Show("Are you ready to submit the decision to the central player? You decision cannot be changed after submission.", "Decision submittion", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                //send message to MDM
                // archive optimization results in another place
                // if (File.Exists(newfile)) File.Delete(newfile);

                string file = Directory.GetCurrentDirectory() + "\\OptimalResult.xls";
                if (File.Exists(file))
                {
                    try
                    {
                        File.Copy(file, newfile);
                    }
                    catch (System.Exception ex)
                    {
                       MessageBox.Show(ex.ToString());
                    }
                    this.Text = "Optimization results are submitted to central player";
                    richTextBoxOEM.Text += "Optimization results are submitted to central player" + "\r\n";
                    richTextBoxOEM.Text += "Please wait until the central player announce the order to run next year" + "\r\n";
                }
                else
                {
                    richTextBoxOEM.Text += "Submission failed. You haven't generate your Optimization results" + "\r\n";
                }
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
            return;
        }

        private void btnMdm_Click(object sender, EventArgs e)
        {
            if (sharedOem == null)
            {
                MessageBox.Show("Please select an OEM first");
                return;
            }
            runningyear = Routines.beginYear;
            string file1;
            if ( !local )
                file1 = @"\\" + machineID + @"\" + sharedOem + @"\WG_input_" + (runningyear - 1).ToString() + ".xls";
            else
                file1 = "C:/WG/OEM/" + sharedOem + "/WG_input_" + (runningyear - 1).ToString() + ".xls";

            if (!File.Exists(file1))
            {
                MessageBox.Show("The MDM result for " + (runningyear - 1).ToString() + " has not been published.");
                return;
            }
            string file2 = Directory.GetCurrentDirectory() + "\\WG_input_" + (runningyear - 1).ToString() + ".xls";
            if (File.Exists(file2)) File.Delete(file2);
            File.Copy(file1, file2);

            DialogResult rs = MessageBox.Show("Do you want to open Excel to analyse profit",
                                  "Evaluation Price", MessageBoxButtons.YesNo,
                                  MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
            if (rs == DialogResult.Yes) Routines.viewEXCEL(file2);
        }

        private void btnSelectOem_Click(object sender, EventArgs e)
        {

            if (!File.Exists(path + @"\announcement\OEM_list.txt"))
            {
                MessageBox.Show("The central has not sent out an OEM list for you to select. Please try it again later.");
                return;
            }

            //selectedOem = listBox1.GetItemText(listBox1.SelectedValue);
            //selectedOem = filenames.ElementAt(0).ToUpper();
            OemMachinName = Environment.MachineName;

            // string sss = "It is time to play year of " + runningyear.ToString() + newline + newline;
            // richTextBoxOEM.Invoke((MethodInvoker)delegate { richTextBoxOEM.Text += sss; });

            IELoad.Enabled = true;
            //loadOemInfo(selectedOem);
            //listBox1.Enabled = false;
            //backgroundWorker1.RunWorkerAsync();
            frm2 = new Form2();
            DialogResult dr = frm2.ShowDialog(this);
            if (dr == DialogResult.Cancel)
            {
                MessageBox.Show("You haven't selected an OEM. Please do it before playing the game.");
                return;
            }
            else if (dr == DialogResult.OK)
            {
                selectedOem = frm2.getText().ToUpper();
                int keyIndex = Array.FindIndex(filenames, w => w.Contains(frm2.getText()));
                sharedOem = oemList[keyIndex];
                // MessageBox.Show("image file= " + "C:\\WG\\Resources\\" + frm2.getText() + "1.jpg" );
                Image image = Image.FromFile("C:\\WG\\Resources\\" + frm2.getText() + "1.jpg" );
                // btnSelectOem.BackgroundImage = (Image)Properties.Resources.ResourceManager.GetObject(frm2.getText() + Convert.ToString(1));
                btnSelectOem.BackgroundImage = image;
                btnSelectOem.BackgroundImageLayout = ImageLayout.Zoom;
                btnSelectOem.Text = null;
                requestTicketName = frm2.returnRequestFile;
                frm2.Close();
                btnSelectOem.Enabled = false;

                // added by JF
                GlobeData.BeginYear = BeginYear.Value.ToString();
                GlobeData.OEM = selectedOem;
            }
        }

        /*
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {

        } */

        private void richTextBoxOEM_TextChanged(object sender, EventArgs e)
        {
            // set the current caret position to the end
            richTextBoxOEM.SelectionStart = richTextBoxOEM.Text.Length;
            // scroll it automatically
            richTextBoxOEM.ScrollToCaret();
            //HighlightWords();
        }

        private void btnSelectOem_Click_1(object sender, EventArgs e)
        {
            //selectedOem = listBox1.GetItemText(listBox1.SelectedValue);
            //selectedOem = filenames.ElementAt(0).ToUpper();
            OemMachinName = Environment.MachineName;
            IELoad.Enabled = true;
            //loadOemInfo(selectedOem);
            //listBox1.Enabled = false;
            //backgroundWorker1.RunWorkerAsync();
            frm2 = new Form2();
            DialogResult dr = frm2.ShowDialog(this);
            if (dr == DialogResult.Cancel)
            {
                MessageBox.Show("You haven't selected an OEM. Please do it before playing the game.");
                return;
                frm2.Close();
                selectedOem = frm2.getText().ToUpper();
                btnSelectOem.BackgroundImage = (Image)Properties.Resources.ResourceManager.GetObject(frm2.getText() + Convert.ToString(1));
                btnSelectOem.BackgroundImageLayout = ImageLayout.Zoom;
                btnSelectOem.Text = null;
                requestTicketName = frm2.returnRequestFile;
                frm2.Close();
                btnSelectOem.Enabled = false;
            }
            else if (dr == DialogResult.OK)
            {
                selectedOem = frm2.getText().ToUpper();
                int keyIndex = Array.FindIndex(filenames, w => w.Contains(frm2.getText()));
                sharedOem = oemList[keyIndex];
                btnSelectOem.BackgroundImage = (Image)Properties.Resources.ResourceManager.GetObject(frm2.getText() + Convert.ToString(1));
                btnSelectOem.BackgroundImageLayout = ImageLayout.Zoom;
                btnSelectOem.Text = null;
                requestTicketName = frm2.returnRequestFile;
                frm2.Close();
                btnSelectOem.Enabled = false;
            }
        }

        private void btnSVL_Click(object sender, EventArgs e)
        {
            if (Routines.population == null) return;
            // Routines.timeHorizon = Convert.ToInt32(timeHorizonUpDown.Value);
            // Routines.beginYear = Convert.ToInt32(beginYearUpDown.Value);
            // Routines.baseYear = Convert.ToInt32(baseYearUpDown.Value);
            int keyIndex = Array.FindIndex(filenames, w => w.Contains(selectedOem.ToLower()));
            Routines.timeHorizon = 1;
            //Routines.beginYear = Routines.recentYear+1;
            //Routines.baseYear = Routines.recentYear;
            Routines.beginYear = runningyear;
            Routines.baseYear = runningyear - 1;

            //Routines.optOEM = oemList[keyIndex].ToLower();
            Routines.optOEM = sharedOem.ToLower();
            //MessageBox.Show(Routines.optOEM);

            if (Routines.baseYear > Routines.recentYear)
            {
                MessageBox.Show("No Historical Data of " + Routines.baseYear);
                return;
            }

            //Cursor.Current = Cursors.WaitCursor;

            btnROOptimal.Enabled = true;
            btnSVL.Enabled = false;

            Szroutine szFun = new Szroutine();
            Routines.genPopu = new Macro[1];
            Routines.genPopu = szFun.getPopu(Routines.population);
            Routines.genVehData = new VehicleData[1][];

            Routines.genVehData = szFun.prefutureData(Routines.vData, Routines.timeHorizon, Routines.Scenario);
            szFun.CalMktShare(Routines.genVehData);

            Routines.optData = new VehicleData[0][];
            Routines.optData = szFun.preOptData(Routines.optOEM, Routines.timeHorizon, Routines.genVehData);
            Routines.listOpt = new ArrayList[0];
            Routines.listOpt = szFun.getOptIndexlist(Routines.optData);
            Routines.genFCons = new Constraint[0];
            Routines.genFCons = szFun.getOptConstraint();

            Routines.listNewFCons = new ArrayList();
            Routines.listNewVData = new listnewvdata[Routines.timeHorizon];
            for (int i = 0; i < Routines.timeHorizon; i++)  //index for
            {
                Routines.listNewVData[i] = new listnewvdata();
                Routines.listNewVData[i].yr = Routines.beginYear + i;
                Routines.listNewVData[i].idx = new ArrayList();
                for (int j = 0; j < Routines.genVehData[i].Length; j++)
                    Routines.listNewVData[i].idx.Add(Routines.genVehData[i][j].modelName.ToLower() +
                      Routines.genVehData[i][j].transYear);
            }
            for (int i = 0; i < Routines.genFCons.Length; i++)
                Routines.listNewFCons.Add(Routines.genFCons[i].ModelName.ToLower() + Routines.genFCons[i].ModelYear);

            //add for listBDelat by sz
            Routines.listBDelta = new ArrayList();
            for (int i = 0; i < Routines.baseDelta.Length; i++)
            {
                Routines.listBDelta.Add(Routines.baseDelta[i].modelName.ToLower());
            }

            //  Routines.optiType = 2;

            //      genVarInd(Routines.timeHorizon, Routines.optData);
            genNormalFactor(Routines.optData);

            Optimization dfFun = new Optimization();
            string NAGResult = dfFun.OptimalRun();
            //richTextBox1.Text += NAGResult;
            btnToMDM.Enabled = true;
            butChart.Enabled = true;
            btnROOptimal.Enabled = true;

            string file = Directory.GetCurrentDirectory() + "\\OptimalResult.csv";

            //show data in datagridview
            importOemOpt(file);
        }

        private void importOemOpt(string fileName)
        {
            dataGridView1.Rows.Clear();
           
            using (StreamReader reader = new StreamReader(fileName))
            {
                string line;
                int i = 0;
                while ((line = reader.ReadLine()) != null)
                {
                    DataGridViewRow row = (DataGridViewRow)dataGridView1.Rows[0].Clone();
                    string[] info = line.Split(',');
                    int ncol = info.Length;
                    row.Cells[1].Value = i;
                    for (int j = 0; j < ncol; j++)
                    {
                        row.Cells[j + 2].Value = info[j];
                    }
                    dataGridView1.Rows.Add(row);
                    row.Cells[0].Value = true;
                    if (row.Cells[0].Value == null) MessageBox.Show("unchecked");
                    i++;
                }
            }
        }

        private void butChart_Click(object sender, EventArgs e)
        {
            // MessageBox.Show("Popup new window");
            // Display Optimization on another Form
            OptimizationDemo form3 = new OptimizationDemo();
            form3.Show();
        }

        private void goNHTSA_Click(object sender, EventArgs e)
        {
            webBrowser1.Navigate("http://www.nhtsa.gov/fuel-economy");
        }

        private void goEPA_Click(object sender, EventArgs e)
        {
            webBrowser1.Navigate("https://www3.epa.gov");
        }

        private void goENERGY_Click(object sender, EventArgs e)
        {
            webBrowser1.Navigate("http://www.fueleconomy.gov/feg/download.shtml");
        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            // MessageBox.Show("Popup new window");
            // Display Optimization on another Form
            selectedCompare = true;
            OptimizationDemo form3 = new OptimizationDemo();
            form3.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // Display EPA Qurry Form on another Form
            EPAquery = true;
            NHTSAquery = false;
            DatabaseQurry form3 = new DatabaseQurry();
            form3.Show();
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            // Display NHITA Qurry Form on another Form
            NHTSAquery = true;
            EPAquery = false;
            DatabaseQurry form3 = new DatabaseQurry();
            form3.Show();
        }

        private void butFordCredit_Click(object sender, EventArgs e)
        {
            // Display Ford 2015 Credit Qurry Form on another Form
            FordCredit = true;
            EPAquery = false;
            NHTSAquery = false;
            DatabaseQurry form3 = new DatabaseQurry();
            form3.Show();
        }      
 
    }
}
