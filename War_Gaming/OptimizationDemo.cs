using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using Microsoft.VisualBasic;
using System.Collections;
using System.Diagnostics;

using System.Data.SqlClient;
using System.Windows.Forms.DataVisualization.Charting;

namespace WindowsFormsApplication2
{
    public partial class OptimizationDemo : Form
    {
        DataTable ResultsHeader = new DataTable();

        public OptimizationDemo()
        {
            InitializeComponent();
            DisplayData();
            DisplayChart();
            // DataTable ResultsHeader = new DataTable();
            if (Form1.selectedCompare)
            {
                // MessageBox.Show("selected tab2");
                tabControl2.SelectedIndex = 2;
            }
        }

        public void DisplayChart()
        {
           OrigChart.ChartAreas[0].AxisX.Interval = 1;
           OrigChart.Series["Price"].Points.DataBindXY((DataView)OptiDataView.DataSource, "Model",(DataView)OptiDataView.DataSource, "price" );
           // Invalidate Chart
           OrigChart.Invalidate();

           OrigChart.Series["Volume"].Points.DataBindXY((DataView)OptiDataView.DataSource, "Model", (DataView)OptiDataView.DataSource, "volume");
           // Invalidate Chart
           OrigChart.Invalidate(); 

           OptiPie.Series[0].IsValueShownAsLabel = false;

           int i = 0;
           foreach (DataRow row in ResultsHeader.Rows)
           {
               var year = row["year"];
               if (Convert.ToString(year) == "2016")
               {
                   var volume1 = row["volume"];
                   var model = row["Model"];
                   OptiPie.Series[0].Points.AddXY(Convert.ToString(model), Convert.ToDouble(volume1));
                   i++;
               }

               // chartCAFEPie.Series[0].Points.AddXY("ICE", ice);
 
           }

 
        }

        public void DisplayData()
        {
            string dataDir = Directory.GetCurrentDirectory();
            // Excel.ApplicationClass ExcelObj = new Excel.ApplicationClass();
            Excel.Application ExcelObj = new Excel.Application();
            // MessageBox.Show("dataDir=" + dataDir);

            Excel.Workbook theWorkbook = ExcelObj.Workbooks.Open(dataDir + "\\OptimalResult.xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            Excel.Sheets sheets = theWorkbook.Worksheets;

            Excel.Worksheet worksheet = (Excel.Worksheet)sheets.get_Item(1);

            Excel.Range range = worksheet.UsedRange;

            System.Array myvalues = (System.Array)range.Cells.Value2;

            int vertical = myvalues.GetLength(0);
            int horizontal = myvalues.GetLength(1);
            horizontal =9;  // only read the first 6 columns


            string[] headers = new string[horizontal];
            string[] data = new string[horizontal];

            // DataTable ResultsHeader = new DataTable();
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
            OptiDataView.DataSource = myDataView;
            // dataGridView3.DataBind();
            // dataGridView3.Dispose();
            theWorkbook.Close();
            ExcelObj.Quit();
        }

        private void OptiPie_Click(object sender, EventArgs e)
        {

        }

        private void checkVol3D_CheckedChanged(object sender, EventArgs e)
        {
            if (checkVol3D.Checked) {
                OrigChart.ChartAreas[0].Area3DStyle.Enable3D = true;
                OptiPie.ChartAreas[0].Area3DStyle.Enable3D = true;
            } else {
                OrigChart.ChartAreas[0].Area3DStyle.Enable3D = false;
                OptiPie.ChartAreas[0].Area3DStyle.Enable3D = false;
            }
        }

        private void OrigChart_Click(object sender, EventArgs e)
        {

        }

        private void OptiDataView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void OptiGroup2_Enter(object sender, EventArgs e)
        {

        }

        private void OptiGroup1_Enter(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void OptimizationDemo_Load(object sender, EventArgs e)
        {

        }

        private void comboXaxleScale_SelectedIndexChanged(object sender, EventArgs e)
        {
            Opti_Year.ChartAreas[0].AxisX.Interval = comboXaxleScale.SelectedIndex + 1;
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked)
            {
                Opti_Year.ChartAreas[0].Area3DStyle.Enable3D = true;
            }
            else
            {
                Opti_Year.ChartAreas[0].Area3DStyle.Enable3D = false;
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            {
                Opti_Year.Series[0].Enabled = true;
                Opti_Year.Series[3].Enabled = true;
                Opti_Year.Series[1].Enabled = false;
                Opti_Year.Series[2].Enabled = false;
                Opti_Year.Series[4].Enabled = false;
                Opti_Year.Series[5].Enabled = false;
                checkBox2.Checked = false;
                checkBox1.Checked = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            foreach (var series in Opti_Year.Series)
            {
                series.Points.Clear();
            }

            comboXaxleScale.SelectedIndex = 0;

            tabControl2.SelectedIndex = 1;
            // Show_Query_Info.Text = "OEM: " + comboOEM.Text + ";  Brand: " + comboBrand.Text + ";  Segment: " + comboSegment.Text + ";  PT Type: " + comboCatalog.Text;

  
            int i, j, k;
            i = j = k = 0;
            Opti_Year.ChartAreas[0].AxisX.Interval = 1;

            foreach (DataRow row in ResultsHeader.Rows)
            {
                var year = row["year"];
                if (Convert.ToString(year) == "2016")
                {
                    var price1 = row["price"];
                    Opti_Year.Series["Price2016"].Points.Add(Convert.ToDouble(price1));
                    Opti_Year.Series["Price2016"].Points[i].Color = Color.Green;
                    Opti_Year.Series["Price2016"].Points[i].AxisLabel = row["Model"].ToString();
                    Opti_Year.Series["Price2016"].Points[i].LegendText = row["Model"].ToString();

                    var volume1 = row["volume"];
                    Opti_Year.Series["Vol_2016"].Points.Add(Convert.ToDouble(volume1));
                    Opti_Year.Series["Vol_2016"].Points[i].Color = Color.Blue;
                    Opti_Year.Series["Vol_2016"].Points[i].AxisLabel = row["Model"].ToString();
                    Opti_Year.Series["Vol_2016"].Points[i].LegendText = row["Model"].ToString();
                    i++;
                }
                else if (Convert.ToString(year) == "2017")
                {
                    var price1 = row["price"];
                    Opti_Year.Series["Price2017"].Points.Add(Convert.ToDouble(price1));
                    Opti_Year.Series["Price2017"].Points[j].Color = Color.Orange;
                    Opti_Year.Series["Price2017"].Points[j].AxisLabel = row["Model"].ToString();
                    Opti_Year.Series["Price2017"].Points[j].LegendText = row["Model"].ToString();

                    var volume1 = row["volume"];
                    Opti_Year.Series["Vol_2017"].Points.Add(Convert.ToDouble(volume1));
                    Opti_Year.Series["Vol_2017"].Points[j].Color = Color.Red;
                    Opti_Year.Series["Vol_2017"].Points[j].AxisLabel = row["Model"].ToString();
                    Opti_Year.Series["Vol_2017"].Points[j].LegendText = row["Model"].ToString();
                    j++;
                }
                else if (Convert.ToString(year) == "2018")
                {
                    var price1 = row["price"];
                    Opti_Year.Series["Price2018"].Points.Add(Convert.ToDouble(price1));
                    Opti_Year.Series["Price2018"].Points[k].Color = Color.Cyan;
                    Opti_Year.Series["Price2018"].Points[k].AxisLabel = row["Model"].ToString();
                    Opti_Year.Series["Price2018"].Points[k].LegendText = row["Model"].ToString();

                    var volume1 = row["volume"];
                    Opti_Year.Series["Vol_2018"].Points.Add(Convert.ToDouble(volume1));
                    Opti_Year.Series["Vol_2018"].Points[k].Color = Color.Yellow;
                    Opti_Year.Series["Vol_2018"].Points[k].AxisLabel = row["Model"].ToString();
                    Opti_Year.Series["Vol_2018"].Points[k].LegendText = row["Model"].ToString();
                    k++;
                }

            } 
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                Opti_Year.Series[0].Enabled = true;
                Opti_Year.Series[3].Enabled = true;
                Opti_Year.Series[1].Enabled = true;
                Opti_Year.Series[4].Enabled = true;
                Opti_Year.Series[2].Enabled = false;
                Opti_Year.Series[5].Enabled = false;
                checkBox1.Checked = false;
                checkBox3.Checked = false;
            }
        }

        private void checkBox1_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                Opti_Year.Series[0].Enabled = true;
                Opti_Year.Series[1].Enabled = true;
                Opti_Year.Series[2].Enabled = true;
                Opti_Year.Series[3].Enabled = true;
                Opti_Year.Series[4].Enabled = true;
                Opti_Year.Series[5].Enabled = true;
                checkBox2.Checked = false; 
                checkBox3.Checked = false;

            }
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            OrigChart.ChartAreas[0].AxisX.Interval = (int)numericUpDown1.Value;
        }

        private void butToDatabase_Click(object sender, EventArgs e)
        {
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


            string sql1 = GlobeData.OEM + "_" + GlobeData.BeginYear;
            string sql = "Create Table " + GlobeData.OEM + "_" + GlobeData.BeginYear + " (";
            // MessageBox.Show(sql1);
            sql1 = "SELECT * FROM " + sql1;
            MessageBox.Show(sql1);
            
            foreach (DataColumn column in ResultsHeader.Columns)
            {
                sql += "[" + column.ColumnName + "] " + "nvarchar(50)" + ",";
            }
            sql = sql.TrimEnd(new char[] { ',' }) + ")";

            SqlCommand cmd = new SqlCommand(sql, cnn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            cmd.ExecuteNonQuery();

            using (var adapter = new SqlDataAdapter(sql1, cnn))
            using (var builder = new SqlCommandBuilder(adapter))
            {
                adapter.InsertCommand = builder.GetInsertCommand();
                adapter.Update(ResultsHeader);
            }

            cnn.Close();
        }

        private void butPlot_Click(object sender, EventArgs e)
        {
            string oem = comboOEM.Text;
           //  MessageBox.Show("OEM=" + oem);
            // retrive data from Database
            string connetionString = null;
            string sSQL = "";
            // bool where = false;

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

            // string str1 = "'" + cDSIDTextBox.Text + "'," + "'" + nAMETextBox.Text + "'," + "'" + sURNAMETextBox.Text + "'," + "'" + pASSWORDTextBox.Text + "'";

            if ( comboOEM.SelectedIndex == 0 )
                sSQL = "select oem, modname, vol2015, price2015, vol2016, price2016, vol2017, price2017, vol2015 * 0.001 * price2015 * 0.001 as [Sales2015(M)], vol2016 * 0.001 * price2016 * 0.001 as [Sales2016(M)], vol2017 * 0.001 * price2017 * 0.001 as [Sales2017(M)] from dbo.VEHDATA2015_2017";
            else
                sSQL = "select modname, vol2015, price2015, vol2016, price2016, vol2017, price2017, vol2015 * 0.001 * price2015 * 0.001 as [Sales2015(M)], vol2016 * 0.001 * price2016 * 0.001 as [Sales2016(M)], vol2017 * 0.001 * price2017 * 0.001 as [Sales2017(M)] from dbo.VEHDATA2015_2017 where oem =" + "'" + oem + "'";

            if (comboOEM.SelectedIndex > 0 && comboBoxGroup.SelectedIndex > 0)
            {
                sSQL = sSQL + " and [group] = " + "'" + comboBoxGroup.Text + "'";
            }
            else if (comboOEM.SelectedIndex == 0 && comboBoxGroup.SelectedIndex > 0)
            {
                sSQL = sSQL + " where [group] = " + "'" + comboBoxGroup.Text + "'";
            }

            SqlCommand cmd = new SqlCommand(sSQL, cnn);

            // cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();
            SqlDataAdapter sqlDataAdap = new SqlDataAdapter(cmd);
            sqlDataAdap.Fill(dt);
            dataGridViewOEM.DataSource = dt; 

            //  Draw Chart here
            foreach (var series in CompareChart.Series)
            {
                series.Points.Clear();
            }

            // DataRow[] row;
            int i = 0;
            CompareChart.Series[6].Enabled = false;
            checkBox6.Checked = false;
            checkBox2017.Checked = true;

            foreach (DataRow row in dt.Rows)
            {
                var v2015 = row["vol2015"];
                CompareChart.Series["Vol2015"].Points.Add(Convert.ToInt32(v2015));
                CompareChart.Series["Vol2015"].Points[i].Color = Color.Yellow;
                CompareChart.Series["Vol2015"].Points[i].AxisLabel = row["modname"].ToString();
                CompareChart.Series["Vol2015"].Points[i].LegendText = row["modname"].ToString();

                var p2015 = row["price2015"];
                CompareChart.Series["Price2015"].Points.Add(Convert.ToInt32(p2015));
                CompareChart.Series["Price2015"].Points[i].Color = Color.Green;
                CompareChart.Series["Price2015"].Points[i].AxisLabel = row["modname"].ToString();
                CompareChart.Series["Price2015"].Points[i].LegendText = row["modname"].ToString();
                CompareChart.Series["Price2015"].YAxisType = AxisType.Secondary;
                
                var v2016 = row["vol2016"];
                CompareChart.Series["Vol2016"].Points.Add(Convert.ToInt32(v2016));
                CompareChart.Series["Vol2016"].Points[i].Color = Color.Blue;
                CompareChart.Series["Vol2016"].Points[i].AxisLabel = row["modname"].ToString();
                CompareChart.Series["Vol2016"].Points[i].LegendText = row["modname"].ToString();

                var p2016 = row["price2016"];
                CompareChart.Series["Price2016"].Points.Add(Convert.ToInt32(p2016));
                CompareChart.Series["Price2016"].Points[i].Color = Color.Red;
                CompareChart.Series["Price2016"].Points[i].AxisLabel = row["modname"].ToString();
                CompareChart.Series["Price2016"].Points[i].LegendText = row["modname"].ToString();
                CompareChart.Series["Price2016"].YAxisType = AxisType.Secondary;

                var v2017 = row["vol2017"];
                CompareChart.Series["Vol2017"].Points.Add(Convert.ToInt32(v2017));
                CompareChart.Series["Vol2017"].Points[i].Color = Color.Tan;
                CompareChart.Series["Vol2017"].Points[i].AxisLabel = row["modname"].ToString();
                CompareChart.Series["Vol2017"].Points[i].LegendText = row["modname"].ToString();
                CompareChart.ChartAreas[0].AxisY.Title = "Volume";

                var p2017 = row["price2017"];
                CompareChart.Series["Price2017"].Points.Add(Convert.ToInt32(p2017));
                CompareChart.Series["Price2017"].Points[i].Color = Color.Orange;
                CompareChart.Series["Price2017"].Points[i].AxisLabel = row["modname"].ToString();
                CompareChart.Series["Price2017"].Points[i].LegendText = row["modname"].ToString();
                CompareChart.Series["Price2017"].YAxisType = AxisType.Secondary;
                CompareChart.ChartAreas[0].AxisY2.Title = "Price";

                var sales = row["Sales2016(M)"];
                CompareChart.Series[6].Points.AddXY(row["modname"], Convert.ToInt32(v2016),  Convert.ToDouble(sales) );

                i++;
            }
        }

        private void checkBox2017_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2017.Checked)
            {
                CompareChart.Series[0].Enabled = true;
                CompareChart.Series[1].Enabled = true;
                CompareChart.Series[2].Enabled = true;
                CompareChart.Series[3].Enabled = true;
                CompareChart.Series[4].Enabled = true;
                CompareChart.Series[5].Enabled = true;
                CompareChart.Series[6].Enabled = false;
                checkBox2015.Checked = false;
                checkBox2016.Checked = false;
                checkBoxPrice.Checked = false;
                checkBoxVolume.Checked = false;
                checkBox6.Checked = false;
            }
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked)
            {
                CompareChart.ChartAreas[0].Area3DStyle.Enable3D = true;
            }
            else
            {
                CompareChart.ChartAreas[0].Area3DStyle.Enable3D = false;
            }

        }

        private void comboBoxXaxleScale_SelectedIndexChanged(object sender, EventArgs e)
        {
            CompareChart.ChartAreas[0].AxisX.Interval = comboBoxXaxleScale.SelectedIndex + 1;
        }

        private void checkBox2015_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2015.Checked)
            {
                CompareChart.Series[0].Enabled = true;
                CompareChart.Series[1].Enabled = false;
                CompareChart.Series[2].Enabled = false;
                CompareChart.Series[3].Enabled = true;
                CompareChart.Series[4].Enabled = false;
                CompareChart.Series[5].Enabled = false;
                CompareChart.Series[6].Enabled = false;
                checkBox2016.Checked = false;
                checkBoxVolume.Checked = false;
                checkBoxPrice.Checked = false;
                checkBox2017.Checked = false; 
                checkBox6.Checked = false;

            }
        }

        private void checkBox2016_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2016.Checked)
            {
                CompareChart.Series[0].Enabled = true;
                CompareChart.Series[1].Enabled = true;
                CompareChart.Series[2].Enabled = false;
                CompareChart.Series[3].Enabled = true;
                CompareChart.Series[4].Enabled = true;
                CompareChart.Series[5].Enabled = false;
                CompareChart.Series[6].Enabled = false;
                checkBox2015.Checked = false;
                checkBox2017.Checked = false;
                checkBoxPrice.Checked = false;
                checkBoxVolume.Checked = false;
                checkBox6.Checked = false;
            }
        }

        private void dataGridViewOEM_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void checkBoxVolume_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxVolume.Checked)
            {
                CompareChart.Series[0].Enabled = true;
                CompareChart.Series[1].Enabled = true;
                CompareChart.Series[2].Enabled = true;
                CompareChart.Series[3].Enabled = false;
                CompareChart.Series[4].Enabled = false;
                CompareChart.Series[5].Enabled = false;
                CompareChart.Series[6].Enabled = false;
                checkBoxPrice.Checked = false;
                checkBox2015.Checked = false;
                checkBox2016.Checked = false;
                checkBox2017.Checked = false;
                checkBox6.Checked = false;
            }
        }

        private void checkBoxPrice_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxPrice.Checked)
            {
                CompareChart.Series[0].Enabled = false;
                CompareChart.Series[1].Enabled = false;
                CompareChart.Series[2].Enabled = false;
                CompareChart.Series[3].Enabled = true;
                CompareChart.Series[4].Enabled = true;
                CompareChart.Series[5].Enabled = true;
                CompareChart.Series[6].Enabled = false;
                checkBoxVolume.Checked = false;
                checkBox2015.Checked = false;
                checkBox2016.Checked = false;
                checkBox2017.Checked = false;
                checkBox6.Checked = false;
            }
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked)
            {
                CompareChart.Series[0].Enabled = false;
                CompareChart.Series[1].Enabled = false;
                CompareChart.Series[2].Enabled = false;
                CompareChart.Series[3].Enabled = false;
                CompareChart.Series[4].Enabled = false;
                CompareChart.Series[5].Enabled = false;
                CompareChart.Series[6].Enabled = true;
                checkBoxVolume.Checked = false;
                checkBoxPrice.Checked = false;
                checkBox2015.Checked = false;
                checkBox2016.Checked = false;
                checkBox2017.Checked = false;

                // CompareChart.Series[6].YAxisType = AxisType.Secondary;
                // CompareChart.ChartAreas[0].AxisY2.Title = "Price";
            }
        }

        private void comboOEM_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void butVolume_Click(object sender, EventArgs e)
        {
            string oem = comboOEM.Text;
            //  MessageBox.Show("OEM=" + oem);
            // retrive data from Database
            string connetionString = null;
            string sSQL = "";
            // bool where = false;

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

            // string str1 = "'" + cDSIDTextBox.Text + "'," + "'" + nAMETextBox.Text + "'," + "'" + sURNAMETextBox.Text + "'," + "'" + pASSWORDTextBox.Text + "'";

            sSQL = "select oem, sum(vol2015*0.001) as v2015K, sum( vol2015*0.001*price2015*0.001 ) as vp2015M, sum(vol2016*0.001) as v2016K, sum( vol2016*0.001*price2016*0.001 ) as vp2016M, sum(vol2017*0.001) as v2017K, sum( vol2017*0.001*price2017*0.001 ) as vp2017M  from dbo.VEHDATA2015_2017 group by oem";

            SqlCommand cmd = new SqlCommand(sSQL, cnn);

            // cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();
            SqlDataAdapter sqlDataAdap = new SqlDataAdapter(cmd);
            sqlDataAdap.Fill(dt);
            dataGridViewOEM.DataSource = dt;

            //  Draw Chart here
            foreach (var series in CompareChart.Series)
            {
                series.Points.Clear();
            }

            // DataRow[] row;
            int i = 0;
            CompareChart.Series[6].Enabled = false;
            checkBox6.Checked = false;
            checkBox2017.Checked = true;

            foreach (DataRow row in dt.Rows)
            {
                var v2015 = row["vol2015"];
                CompareChart.Series["Vol2015"].Points.Add(Convert.ToInt32(v2015));
                CompareChart.Series["Vol2015"].Points[i].Color = Color.Yellow;
                CompareChart.Series["Vol2015"].Points[i].AxisLabel = row["modname"].ToString();
                CompareChart.Series["Vol2015"].Points[i].LegendText = row["modname"].ToString();

                var p2015 = row["price2015"];
                CompareChart.Series["Price2015"].Points.Add(Convert.ToInt32(p2015));
                CompareChart.Series["Price2015"].Points[i].Color = Color.Green;
                CompareChart.Series["Price2015"].Points[i].AxisLabel = row["modname"].ToString();
                CompareChart.Series["Price2015"].Points[i].LegendText = row["modname"].ToString();
                CompareChart.Series["Price2015"].YAxisType = AxisType.Secondary;

                var v2016 = row["vol2016"];
                CompareChart.Series["Vol2016"].Points.Add(Convert.ToInt32(v2016));
                CompareChart.Series["Vol2016"].Points[i].Color = Color.Blue;
                CompareChart.Series["Vol2016"].Points[i].AxisLabel = row["modname"].ToString();
                CompareChart.Series["Vol2016"].Points[i].LegendText = row["modname"].ToString();

                var p2016 = row["price2016"];
                CompareChart.Series["Price2016"].Points.Add(Convert.ToInt32(p2016));
                CompareChart.Series["Price2016"].Points[i].Color = Color.Red;
                CompareChart.Series["Price2016"].Points[i].AxisLabel = row["modname"].ToString();
                CompareChart.Series["Price2016"].Points[i].LegendText = row["modname"].ToString();
                CompareChart.Series["Price2016"].YAxisType = AxisType.Secondary;

                var v2017 = row["vol2017"];
                CompareChart.Series["Vol2017"].Points.Add(Convert.ToInt32(v2017));
                CompareChart.Series["Vol2017"].Points[i].Color = Color.Tan;
                CompareChart.Series["Vol2017"].Points[i].AxisLabel = row["modname"].ToString();
                CompareChart.Series["Vol2017"].Points[i].LegendText = row["modname"].ToString();
                CompareChart.ChartAreas[0].AxisY.Title = "Volume";

                var p2017 = row["price2017"];
                CompareChart.Series["Price2017"].Points.Add(Convert.ToInt32(p2017));
                CompareChart.Series["Price2017"].Points[i].Color = Color.Orange;
                CompareChart.Series["Price2017"].Points[i].AxisLabel = row["modname"].ToString();
                CompareChart.Series["Price2017"].Points[i].LegendText = row["modname"].ToString();
                CompareChart.Series["Price2017"].YAxisType = AxisType.Secondary;
                CompareChart.ChartAreas[0].AxisY2.Title = "Price";

                var sales = row["Sales2016(M)"];
                CompareChart.Series[6].Points.AddXY(row["modname"], Convert.ToInt32(v2016), Convert.ToDouble(sales));

                i++;
            }
        }
    }
}
