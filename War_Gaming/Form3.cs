using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Data.SqlClient;

using System.Windows.Forms.DataVisualization.Charting;


namespace WindowsFormsApplication2
{
    public partial class DatabaseQurry : Form
    {
        int bev, cng, hev, ice, lpg, phev;
        DataTable globaldt = new DataTable();

        public DatabaseQurry()
        {
            InitializeComponent();
            if (Form1.EPAquery)
            {
                tabControl1.SelectedIndex = 2;
            }
            else if (Form1.NHTSAquery)
            {
                tabControl1.SelectedIndex = 3;
            }
            else if (Form1.FordCredit)
            {
                tabControl1.SelectedIndex = 4;
            }
        }

        private void _2013_CAFEBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this._2013_CAFEBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.warGamingDataSet3);

        }

        private void DatabaseQurry_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'warGamingDataSet3._2013_CAFE' table. You can move, or remove it, as needed.
            // this._2013_CAFETableAdapter.Fill(this.warGamingDataSet3._2013_CAFE);
            comboOEM.SelectedIndex = 0;
            comboBrand.SelectedIndex = 0;
            comboSegment.SelectedIndex = 0;
            comboCatalog.SelectedIndex = 0;
        }

        private void butUpdate_Click(object sender, EventArgs e)
        {
            string connetionString = null;
            bool where = false;
            string sSQL, group=null;

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

            if (!checkGroup.Checked)
                sSQL = "select Manufacturer, Brand, [Segment (FEV)], [ICE/xEV], Nameplate, Liters, Sales, MPG, 8887/MPG as CO2, Footprint, Sales/MPG as [Vol/MPG], Sales*Footprint as [Vol*Footprint] from dbo.[2013_CAFE-S]";
            else
            {
                sSQL = "select Manufacturer, Brand, [Segment (FEV)], [ICE/xEV], Nameplate, sum(Sales) Sales, sum(MPG*Sales)/sum(Sales) MPG, 8887/sum(MPG*Sales)*sum(Sales) as CO2, Footprint, sum(Sales)/sum(MPG*Sales)*sum(Sales) as [Vol/MPG], sum(Sales)*Footprint as [Vol*Footprint] from dbo.[2013_CAFE-S] ";
                group = " group by Nameplate, [ICE/xEV],[Segment (FEV)],Brand, Manufacturer, Footprint";
            }

 
            if (comboOEM.Text != "All")
            {
                    sSQL = sSQL + " where Manufacturer = " + "'" + comboOEM.Text + "'";
                    where = true;
                    // MessageBox.Show(sSQL);
            }
     

            if (comboBrand.SelectedIndex > 0)
            {
                if (where)
                    sSQL = sSQL + " and Brand = " + "'" + comboBrand.Text + "'";
                else
                {
                    sSQL = sSQL + " Where Brand = " + "'" + comboBrand.Text + "'";
                    where = true;
                }
            }

            if (comboSegment.SelectedIndex > 0)
            {
                if (where)
                    sSQL = sSQL + " and [Segment (FEV)] = " + "'" + comboSegment.Text + "'";
                else
                {
                    sSQL = sSQL + " Where [Segment (FEV)] = " + "'" + comboSegment.Text + "'";
                    where = true;
                }
            }

            if (comboCatalog.SelectedIndex > 0)
            {
                if (where)
                    sSQL = sSQL + " and [ICE/xEV] = " + "'" + comboCatalog.Text + "'";
                else
                {
                    sSQL = sSQL + " Where [ICE/xEV] = " + "'" + comboCatalog.Text + "'";
                    where = true;
                }
            }

            if ( checkGroup.Checked )
                sSQL = sSQL + group;

            // MessageBox.Show("sSQL=" + sSQL);
            SqlCommand cmd = new SqlCommand(sSQL, cnn);

            // cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();
            SqlDataAdapter sqlDataAdap = new SqlDataAdapter(cmd);
            sqlDataAdap.Fill(dt);
            dataCAFE.DataSource = dt;
            globaldt = dt;

 


            //  Draw Chart here

            foreach (var series in chartCAFEBar.Series)
            {
                series.Points.Clear();
            }

            // DataRow[] row;
            int i = 0;
            int sale;

            bev = cng = hev = ice = lpg = phev = 0;
            foreach (DataRow row in dt.Rows)
            {
                var cellCO2 = row["CO2"];
                chartCAFEBar.Series["CO2/10"].Points.Add( Convert.ToDouble(cellCO2) / 10 );
                chartCAFEBar.Series["CO2/10"].Points[i].Color = Color.Red;
                chartCAFEBar.Series["CO2/10"].Points[i].AxisLabel = row["Nameplate"].ToString();
                chartCAFEBar.Series["CO2/10"].Points[i].LegendText = row["Nameplate"].ToString();

                var cellMPG = row["MPG"];
                chartCAFEBar.Series["MPG"].Points.Add(Convert.ToDouble(cellMPG));
                chartCAFEBar.Series["MPG"].Points[i].Color = Color.Green;
                chartCAFEBar.Series["MPG"].Points[i].AxisLabel = row["Nameplate"].ToString();
                chartCAFEBar.Series["MPG"].Points[i].LegendText = row["Nameplate"].ToString();

                var cellVol = row["Sales"];
                chartCAFEBar.Series["VOL(K)"].Points.Add(Convert.ToDouble(cellVol)/1000.0);
                chartCAFEBar.Series["VOL(K)"].Points[i].Color = Color.Blue;
                chartCAFEBar.Series["VOL(K)"].Points[i].AxisLabel = row["Nameplate"].ToString();
                chartCAFEBar.Series["VOL(K)"].Points[i].LegendText = row["Nameplate"].ToString();

                // chartDisplay.Series[0].Points.AddXY(row["Nameplate"].ToString(), Convert.ToDouble(cellVol) / 1000.0, Convert.ToDouble(cellCO2) / 10);

                var cellICE = row["ICE/xEV"];
                var cellSale = row["Sales"];
                try
                {
                    sale = Convert.ToInt32(cellSale);
                    if (cellICE.ToString() == "ICE")
                    {
                        ice += sale;
                    }
                    else if (cellICE.ToString() == "BEV") {
                        bev += sale;
                    }
                    else if (cellICE.ToString() == "CNG")
                    {
                        cng += sale;
                    }
                    else if (cellICE.ToString() == "HEV")
                    {
                        hev += sale;
                    }
                    else if (cellICE.ToString() == "LPG")
                    {
                        lpg += sale;
                    }
                    else if (cellICE.ToString() == "PHEV")
                    {
                        phev += sale;
                    }

                }
                catch (Exception)
                {
                }

 
                i++;
            }

            object sum;
            // sum = dt.Compute(“Sum("Sales")”, “”); //DT is the DataTable and Amount is the Column to SUM in DataTable.
            sum = dt.Compute("Sum(Sales)", "");
            // int sumInt = Convert.ToInt32(dt.Compute("SUM(Sales)", string.Empty));
            // MessageBox.Show("sumInt= " + sumInt);

            object vol_MPG;
            vol_MPG = dt.Compute("Sum([Vol/MPG])", "");

            object volxFootprint;
            volxFootprint = dt.Compute("Sum([Vol*Footprint])", "");

            object avg;
            avg = dt.Compute("AVG(MPG)", "");


            richTextDataQurry.AppendText("Total Sales: " + sum + "\n");
            richTextDataQurry.AppendText("ICE Sales: " + ice + "\n");
            richTextDataQurry.AppendText("HEV Sales: " + hev + "\n");
            richTextDataQurry.AppendText("BEV Sales: " + bev + "\n");

            try
            {
                richTextDataQurry.AppendText("MPG Harmonic Average: " + float.Parse(sum.ToString()) / float.Parse(vol_MPG.ToString()) + "\n");
                richTextDataQurry.AppendText("Sales Weighted Average of Footprint: " + float.Parse(volxFootprint.ToString()) / float.Parse(sum.ToString()) + "\n");
            }
            catch (Exception)
            {
            }
            richTextDataQurry.AppendText("\n");

            // draw Pie chart here
            foreach (var series in chartCAFEPie.Series)
            {
                series.Points.Clear();
            }
            /* Take Colors To Display Pie In That Colors Of Taken Five Values.
            Color[] myPieColors = { Color.Red, Color.Black, Color.Blue, Color.Green, Color.Maroon };
            OptiPie.Series[0]
            chartCAFEPie.Series[0].Points.DataBindXY((DataView)OptiDataView.DataSource, "modelName", (DataView)OptiDataView.DataSource, "OrigVolume");
            chartCAFEPie.Series[0].IsValueShownAsLabel = false;
            */

            chartCAFEPie.Series[0].Points.AddXY("ICE", ice);
            chartCAFEPie.Series[0].Points.AddXY("HEV", hev);
            chartCAFEPie.Series[0].Points.AddXY("BEV", bev);
            chartCAFEPie.Series[0].Points.AddXY("PHEV", phev);
            chartCAFEPie.Series[0].Points.AddXY("LPG", lpg);
            chartCAFEPie.Series[0].Points.AddXY("CNG", cng); 

            cnn.Close();

            // Show meassage in textBox Show_Query_Info
            butPiePT.Enabled = true;
        }

        private void butClean_Click(object sender, EventArgs e)
        {
            richTextDataQurry.Text = "";
        }

        private void comboOEM_SelectedIndexChanged(object sender, EventArgs e)
        {   
            if ( comboOEM.SelectedIndex >= 0 )
            {
                comboBrand.DataSource = null;              
                comboBrand.Items.Clear();
                string strClass = string.Empty;
                strClass = (string)comboOEM.Text;
                List<string> list = null;
            
                // Bind Names dropdownlist based on Class value
                list = GetNamesByClass(strClass);
                comboBrand.DataSource = list;
            }

        }

        private List<string> GetNamesByClass(string clsss)
        {
            string[] Ford = {"All","FORD", "LINCOLN", "Roush Industries" };
            string[] BMW = { "All","BMW", "MINI" };
            string[] Daimler = {"All","MERCEDES BENZ", "Mercedes Benz" };
            string[] FCA = { "All","chryler", "Dodge", "Fiat", "Jeep", "RAM"};
            string[] GM = { "All","BUICK", "CADILIAC", "CHEVROLET", "GMC" };
            string[] Honda = {"All", "ACURA", "HONDA" };
            string[] Hyundai = {"HYUNDAI"};
            string[] Kia = { "KIA"};
            string[] Mazda = { "ALL","MAZDA", "Mazda" };
            string[] Nissan = {"All", "INFINITI", "NISSAN" };
            string[] Tesla = { "TESLA"};
            string[] Toyota = {"All", "LEXUS", "SCION", "TOYOTA" };
            string[] Volkswagen = {"All", "AUDI", "Porsche", "Volkswagen" };
            string[] Volvo = { "VOLVO" };
            
            string[] All = {"All","ACURA","AUDI","BMW","BUICK","CADILIAC","CHEVROLET","Chrysler","Dodge","Fiat","FORD","GMC","HONDA","HYUNDAI","INIFINITI","Jeep","KIA","LEXUS","LINCOLN","MAZDA","Mazda","MERCEDES BENZ","Mercedes Benz","MINI","NISSAN","Porsch","RAM","Poush Industries", "SCION","TESAL","TOYOTA","Volkswagen","VOLVO"};

            List<string> listRange = new List<string>();
            if (clsss == "All")
            {
                listRange = new List<string>(All);
            }
            else if (clsss == "Ford")
            {
                listRange = new List<string>(Ford);
            } else if ( clsss == "BMW") {
                //listRange.Add("BMW");
                listRange = new List<string>(BMW);
            }
            else if (clsss == "Daimler")
            {
                listRange = new List<string>(Daimler);
            }
            else if (clsss == "Fiat Chrysler")
            {
                listRange = new List<string>(FCA);
            }
            else if (clsss == "General Motors")
            {
                listRange = new List<string>(GM);
            }
            else if (clsss == "Honda")
            {
                listRange = new List<string>(Honda);
            }
            else if (clsss == "Hyundai")
            {
                listRange = new List<string>(Hyundai);
            }
            else if (clsss == "Kia")
            {
                listRange = new List<string>(Kia);
            }
            else if (clsss == "Mazda")
            {
                listRange = new List<string>(Mazda);
            }
            else if (clsss == "Nissan")
            {
                listRange = new List<string>(Nissan);
            }
            else if (clsss == "Tesla")
            {
                listRange = new List<string>(Tesla);
            }
            else if (clsss == "Toyota")
            {
                listRange = new List<string>(Toyota);
            }
            else if (clsss == "Volkswagen")
            {
                listRange = new List<string>(Volkswagen);
            }
            else if (clsss == "Volvo")
            {
                listRange = new List<string>(Volvo);
            }
 

            return listRange;
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void chartCAFEBar_Click(object sender, EventArgs e)
        {

        }

        private void butPiePT_Click(object sender, EventArgs e)
        {
            /*
            Chart Chart0 = new Chart();
            Chart0 = chartDisplay;
            ChartArea ChartArea0 = new ChartArea("name");
            Chart0.ChartAreas.Add(ChartArea0);
            Series Series0 = new Series();
            Chart0.Series.Add(Series0);
            // link series to area here
            Series0.ChartArea = "name"; 
             
            chart1.ChartAreas.Clear();
            chart1.ChartAreas.Add(chartarea1);
            chart1.ChartAreas.Add(chartarea2);
             
            */



            foreach (var series in chartDisplay.Series)
            {
                series.Points.Clear();
            }

            comboBoxXaxleScale.SelectedIndex = 0;

            int i = 0;

            foreach (DataRow row in globaldt.Rows)
            {
                var cellCO2 = row["CO2"];

                var cellMPG = row["MPG"];

                var cellVol = row["Sales"];

                chartDisplay.Series[0].Points.AddXY(row["Nameplate"].ToString(), Convert.ToDouble(cellVol), Convert.ToDouble(cellCO2) / 10);
                chartDisplay.Series[1].Points.AddXY(row["Nameplate"].ToString(), Convert.ToDouble(cellVol));
                chartDisplay.ChartAreas[0].AxisY.Title = "Sales";

                /*
                chartDisplay.Series["Bar"].Points.Add(Convert.ToDouble(cellVol) / 1000.0);
                chartDisplay.Series["Bar"].Points[i].Color = Color.Blue;
                chartDisplay.Series["Bar"].Points[i].AxisLabel = row["Nameplate"].ToString();
                chartDisplay.Series["Bar"].Points[i].LegendText = row["Nameplate"].ToString(); */

                i++;
            }

            tabControl1.SelectedIndex = 1;
            Show_Query_Info.Text = "OEM: " + comboOEM.Text + ";  Brand: " + comboBrand.Text + ";  Segment: " + comboSegment.Text + ";  PT Type: " + comboCatalog.Text;

            // Draw StackedArea100 
            /*
            chartDisplay.Series[1].Points.AddXY("N1", 3);
            chartDisplay.Series[1].Points.AddXY("N2", 5);
            chartDisplay.Series[1].Points.AddXY("N3", 7);
            chartDisplay.Series[1].Points.AddXY("N4", 9); */

            // Draw Line for Sales




            // draw bubble chart in another tab
            // Create a new chart.
            // ChartControl bubbleChart = new ChartControl();

            // Create two bubble series.
            // Series series1 = new Series("Series 1", ViewType.Bubble);
            // Series series2 = new Series("Series 2", ViewType.Bubble);

            /* Add points to them.
            chartDisplay.Series[0].Points.AddXY(1, 2, 5);
            chartDisplay.Series[0].Points.AddXY(2, 1, 5);
            chartDisplay.Series[0].Points.AddXY(3, 3, 6);
            chartDisplay.Series[0].Points.AddXY(4, 2, 6);
            chartDisplay.Series[0].Points.AddXY(5, 5, 8);
            // chartDisplay.Series[0].Points. */


            // Add both series to the chart.
            // bubbleChart.Series.AddRange(new Series[] { series1, series2 });

            // Set the numerical argument scale types for the series,
            // as it is qualitative, by default.
            // series1.ArgumentScaleType = ScaleType.Numerical;
            // series2.ArgumentScaleType = ScaleType.Numerical;

            // Access the view-type-specific options of the series.
            // ((BubbleSeriesView)series1.View).MaxSize = 2;
            // ((BubbleSeriesView)series1.View).MinSize = 1;
            // ((BubbleSeriesView)series1.View).BubbleMarkerOptions.Kind = MarkerKind.Circle;

            // Access the type-specific options of the diagram.
            // ((XYDiagram)bubbleChart.Diagram).EnableAxisXZooming = true;

            // Hide the legend (if necessary).
            // bubbleChart.Legend.Visible = false;

            // Add the chart to the form.
            // bubbleChart.Dock = DockStyle.Fill;
            // this.Controls.Add(bubbleChart);

        }

        private void checkDataQuery3D_CheckedChanged(object sender, EventArgs e)
        {
            if (checkDataQuery3D.Checked)
            {
                chartCAFEBar.ChartAreas[0].Area3DStyle.Enable3D = true;
                chartCAFEPie.ChartAreas[0].Area3DStyle.Enable3D = true;
            } else
            {
                chartCAFEBar.ChartAreas[0].Area3DStyle.Enable3D = false;
                chartCAFEPie.ChartAreas[0].Area3DStyle.Enable3D = false;
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                chartDisplay.ChartAreas[0].Area3DStyle.Enable3D = true;
            }
            else
            {
                chartDisplay.ChartAreas[0].Area3DStyle.Enable3D = false;
            }
        }

        private void comboBoxXaxleScale_SelectedIndexChanged(object sender, EventArgs e)
        {
            chartDisplay.ChartAreas[0].AxisX.Interval = comboBoxXaxleScale.SelectedIndex + 1;
        }

        private void checkBoxCO2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxCO2.Checked)
            {
                chartDisplay.Series[0].Enabled = true;
                chartDisplay.Series[1].Enabled = false;
                checkBoxLine.Checked = false;
                checkBoxAll.Checked = false;
            }
        }

        private void checkBoxLine_CheckedChanged(object sender, EventArgs e)
        {
            chartDisplay.Series[1].Enabled = true;
            chartDisplay.Series[0].Enabled = false;
            checkBoxCO2.Checked = false;
            checkBoxAll.Checked = false;
        }

        private void checkBoxAll_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxAll.Checked)
            {
                chartDisplay.Series[0].Enabled = true;
                chartDisplay.Series[1].Enabled = true;
                checkBoxCO2.Checked = false;
                checkBoxLine.Checked = false;
            }
            else
            {
                chartDisplay.Series[0].Enabled = false;
                chartDisplay.Series[1].Enabled = false;
                // checkBoxCO2.Checked = false;
                // checkBoxLine.Checked = false;
            }
        }

        private void butPlotEPA_Click(object sender, EventArgs e)
        {
            double[] car   = new double[16] { 2698609,  5312639,  3587117,  7174385, 12953128, 17086473, 10386975,  484461,        0, -3417378, -6334719, -1.1e+07, -1.5e+07, -1.9e+07, -2.3e+07, -2.8e+07 };
            double[] truck = new double[16] { 5859831, 10462767, 12488771, 13542503, 15593307, 16304588,  9591025, 4802936, -1.2e+07, -2.4e+07, -3.6e+07, -4.7e+07, -5.7e+07, -7.5e+07, -8.7e+07, -1.0e+08 };
            int[]     year = new int[16] { 2009, 2010, 2011, 2012, 2013, 2014, 2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025 };
 
            int length = Convert.ToInt16(comboYear.Text) + 2008;
            int length1 = length - 2008;

            if (length1 > 6) length1--;


            // retrive data from Database
            string connetionString = null;
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

            string sSQL = "select YEAR, CREDIT, DEFICIT from dbo.Result_EPA_credit_before where YEAR <= " +  "'" + length + "'";

            /*
            if (comboBoxGroup.SelectedIndex > 0)
            {
                sSQL = sSQL + " and [group] = " + "'" + comboBoxGroup.Text + "'";
            } */

            SqlCommand cmd = new SqlCommand(sSQL, cnn);

            // cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();
            SqlDataAdapter sqlDataAdap = new SqlDataAdapter(cmd);
            sqlDataAdap.Fill(dt);
            // dataGridViewOEM.DataSource = dt; 

            //  Draw Chart here
            foreach (var series in chartEPA.Series)
            {
                series.Points.Clear();
            }

            foreach (var series in chartBank.Series)
            {
                series.Points.Clear();
            }

            // DataRow[] row;
            int i = 0;
            int j = 0;

            chartEPA.ChartAreas[0].AxisX.Interval = 1;
            chartBank.ChartAreas[0].AxisX.Interval = 1;


            foreach (DataRow row in dt.Rows)
            {
                if ( j % 3 == 0 ) { 
                    var creditC = row["CREDIT"];
                    var deficitC = row["DEFICIT"];

                    chartEPA.Series["CarCredit"].Points.Add(Convert.ToInt32(creditC) + Convert.ToInt32(deficitC));

                    if ( Convert.ToInt32(creditC) + Convert.ToInt32(deficitC) >= 0 )
                        chartEPA.Series["CarCredit"].Points[i].Color = Color.Green;
                    else
                        chartEPA.Series["CarCredit"].Points[i].Color = Color.Red;

                    chartEPA.Series["CarCredit"].Points[i].AxisLabel = row["YEAR"].ToString();
                    chartEPA.Series["CarCredit"].Points[i].LegendText = row["YEAR"].ToString();

                }
                else if ( j % 3 == 1) {
                    var creditT = row["CREDIT"];
                    var deficitT = row["DEFICIT"];
                    chartEPA.Series["TruckCredit"].Points.Add(Convert.ToInt32(creditT) + Convert.ToInt32(deficitT));

                    if (Convert.ToInt32(creditT) + Convert.ToInt32(deficitT) >= 0)
                        chartEPA.Series["TruckCredit"].Points[i].Color = Color.Blue;
                    else
                        chartEPA.Series["TruckCredit"].Points[i].Color = Color.Brown;

                    chartEPA.Series["TruckCredit"].Points[i].AxisLabel = row["YEAR"].ToString();
                    chartEPA.Series["TruckCredit"].Points[i].LegendText = row["YEAR"].ToString();

                    i++;
                } else 
                {
                }

                j++;
            }

            for (i = 0; i < length1; i++)
            {
                chartBank.Series["CarBank"].Points.Add(car[i]);

                if ( car[i] >= 0)
                    chartBank.Series["CarBank"].Points[i].Color = Color.Green;
                else
                    chartBank.Series["CarBank"].Points[i].Color = Color.Red;
 
                chartBank.Series["CarBank"].Points[i].AxisLabel = year[i].ToString();
                chartBank.Series["CarBank"].Points[i].LegendText = year[i].ToString();

                chartBank.Series["TruckBank"].Points.Add(truck[i]);

                if (truck[i] >= 0)
                    chartBank.Series["TruckBank"].Points[i].Color = Color.Blue;
                else
                    chartBank.Series["TruckBank"].Points[i].Color = Color.Brown;

                chartBank.Series["TruckBank"].Points[i].AxisLabel = year[i].ToString();
                chartBank.Series["TruckBank"].Points[i].LegendText = year[i].ToString();
            }
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void comboXaxleScale_SelectedIndexChanged(object sender, EventArgs e)
        {
            chartEPA.ChartAreas[0].AxisX.Interval = comboBoxXaxleScale.SelectedIndex + 1;
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked)
            {
                chartEPA.ChartAreas[0].Area3DStyle.Enable3D = true;
            }
            else
            {
                chartEPA.ChartAreas[0].Area3DStyle.Enable3D = false;
            }
        }

        private void checkBoxCar_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxCar.Checked)
            {
                chartEPA.Series[0].Enabled = true;
                chartEPA.Series[1].Enabled = false;
                checkBoxTruck.Checked = false;
                checkBoxShowAll.Checked = false;
            }
        }

        private void checkBoxTruck_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxTruck.Checked)
            {
                chartEPA.Series[0].Enabled = false;
                chartEPA.Series[1].Enabled = true;
                checkBoxCar.Checked = false;
                checkBoxShowAll.Checked = false;
            }
        }

        private void checkBoxShowAll_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxShowAll.Checked)
            {
                chartEPA.Series[0].Enabled = true;
                chartEPA.Series[1].Enabled = true;
                checkBoxCar.Checked = false;
                checkBoxTruck.Checked = false;
            }
        }

        private void tabPage4_Click(object sender, EventArgs e)
        {

        }

        private void butNHTSAPlot_Click(object sender, EventArgs e)
        {
            double[] dp = new double[17] { 18198882, 35046802, 71422450, 95546303, 1.22e+08, 1.62e+08, 1.11e+08, 85269685, 69344277, 35689421,  6388304, 6388305,         0, -4504895, -1.4e+07,   -3e+07, -6.4e+07 };
            double[] lt = new double[17] { 13928915, 22696550, 30284389, 36113884, 36815111, 40514897, 30990454, 14634980,  8805485, 12956466, 13467292, 24118571, 26209363, 26288329, 28108484, 15359847,       -1 };
            double[] ip = new double[17] {  7301196,  7332814,  7332814,  7332814,  7306574,  7250333,   103066,    55870,     2699,   -49317,   -69351,   -77672,   -77673,  -100152,  -140881,  -218488,  -350579 };
            int[] year = new int[17] { 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025 };

            int length = Convert.ToInt16(comboBoxNHTSA.Text) + 2007;
            int length1 = length - 2007;

            if (length1 > 7) length1--;


            // retrive data from Database
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


            string sSQL = "select YEAR, CREDIT, DEFICIT from dbo.Result_NHTSA_credit_before where YEAR <= " + "'" + length + "'";

            SqlCommand cmd = new SqlCommand(sSQL, cnn);

            // cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();
            SqlDataAdapter sqlDataAdap = new SqlDataAdapter(cmd);
            sqlDataAdap.Fill(dt);
            // dataGridViewOEM.DataSource = dt; 

            //  Draw Chart here
            foreach (var series in chartNHTSA.Series)
            {
                series.Points.Clear();
            }

            foreach (var series in chartNHTSABank.Series)
            {
                series.Points.Clear();
            }

            // DataRow[] row;
            int i = 0;
            int j = 0;

            chartNHTSA.ChartAreas[0].AxisX.Interval = 1;
            chartNHTSABank.ChartAreas[0].AxisX.Interval = 1;


            foreach (DataRow row in dt.Rows)
            {
                if (j % 6 == 0)
                {
                    var creditC = row["CREDIT"];
                    var deficitC = row["DEFICIT"];

                    chartNHTSA.Series["DPCredit"].Points.Add(Convert.ToInt32(creditC) + Convert.ToInt32(deficitC));

                    if (Convert.ToInt32(creditC) + Convert.ToInt32(deficitC) >= 0)
                        chartNHTSA.Series["DPCredit"].Points[i].Color = Color.Green;
                    else
                        chartNHTSA.Series["DPCredit"].Points[i].Color = Color.Red;

                    chartNHTSA.Series["DPCredit"].Points[i].AxisLabel = row["YEAR"].ToString();
                    chartNHTSA.Series["DPCredit"].Points[i].LegendText = row["YEAR"].ToString();

                }
                else if (j % 6 == 1)
                {
                    var creditT = row["CREDIT"];
                    var deficitT = row["DEFICIT"];
                    chartNHTSA.Series["LTCredit"].Points.Add(Convert.ToInt32(creditT) + Convert.ToInt32(deficitT));

                    if (Convert.ToInt32(creditT) + Convert.ToInt32(deficitT) >= 0)
                        chartNHTSA.Series["LTCredit"].Points[i].Color = Color.Blue;
                    else
                        chartNHTSA.Series["LTCredit"].Points[i].Color = Color.Brown;

                    chartNHTSA.Series["LTCredit"].Points[i].AxisLabel = row["YEAR"].ToString();
                    chartNHTSA.Series["LTCredit"].Points[i].LegendText = row["YEAR"].ToString();
                }
                else if (j % 6 == 2)
                {
                    var creditT = row["CREDIT"];
                    var deficitT = row["DEFICIT"];
                    chartNHTSA.Series["IPCredit"].Points.Add(Convert.ToInt32(creditT) + Convert.ToInt32(deficitT));

                    if (Convert.ToInt32(creditT) + Convert.ToInt32(deficitT) >= 0)
                        chartNHTSA.Series["IPCredit"].Points[i].Color = Color.DarkGreen;
                    else
                        chartNHTSA.Series["IPCredit"].Points[i].Color = Color.DarkRed;

                    chartNHTSA.Series["IPCredit"].Points[i].AxisLabel = row["YEAR"].ToString();
                    chartNHTSA.Series["IPCredit"].Points[i].LegendText = row["YEAR"].ToString();

                    i++;
                } else
                {
                }

                j++;
            }

            for (i = 0; i < length1; i++)
            {
                chartNHTSABank.Series["DPBank"].Points.Add(dp[i]);

                if (dp[i] >= 0)
                    chartNHTSABank.Series["DPBank"].Points[i].Color = Color.Green;
                else
                    chartNHTSABank.Series["DPBank"].Points[i].Color = Color.Red;

                chartNHTSABank.Series["DPBank"].Points[i].AxisLabel = year[i].ToString();
                chartNHTSABank.Series["DPBank"].Points[i].LegendText = year[i].ToString();

                chartNHTSABank.Series["LTBank"].Points.Add(lt[i]);

                if (lt[i] >= 0)
                    chartNHTSABank.Series["LTBank"].Points[i].Color = Color.Blue;
                else
                    chartNHTSABank.Series["LTBank"].Points[i].Color = Color.Brown;

                chartNHTSABank.Series["LTBank"].Points[i].AxisLabel = year[i].ToString();
                chartNHTSABank.Series["LTBank"].Points[i].LegendText = year[i].ToString();

                chartNHTSABank.Series["IPBank"].Points.Add(ip[i]);

                if (ip[i] >= 0)
                    chartNHTSABank.Series["IPBank"].Points[i].Color = Color.DarkGreen;
                else
                    chartNHTSABank.Series["IPBank"].Points[i].Color = Color.DarkRed;

                chartNHTSABank.Series["IPBank"].Points[i].AxisLabel = year[i].ToString();
                chartNHTSABank.Series["IPBank"].Points[i].LegendText = year[i].ToString();
            }
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked)
            {
                chartNHTSA.ChartAreas[0].Area3DStyle.Enable3D = true;
            }
            else
            {
                chartNHTSA.ChartAreas[0].Area3DStyle.Enable3D = false;
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked)
            {
                chartNHTSA.Series[0].Enabled = true;
                chartNHTSA.Series[1].Enabled = false;
                chartNHTSA.Series[2].Enabled = false;
                checkBox3.Checked = false;
                checkBox7.Checked = false;
                checkBox2.Checked = false;
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            {
                chartNHTSA.Series[0].Enabled = false;
                chartNHTSA.Series[1].Enabled = true;
                chartNHTSA.Series[2].Enabled = false;
                checkBox2.Checked = false;
                checkBox7.Checked = false;
                checkBox4.Checked = false;
            }
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox7.Checked)
            {
                chartNHTSA.Series[0].Enabled = false;
                chartNHTSA.Series[1].Enabled = false;
                chartNHTSA.Series[2].Enabled = true;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                chartNHTSA.Series[0].Enabled = true;
                chartNHTSA.Series[1].Enabled = true;
                chartNHTSA.Series[2].Enabled = true;
                checkBox7.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
            }
        }

        private void comboBrand_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void richTextDataQurry_TextChanged(object sender, EventArgs e)
        {

        }

        private void checkGroup_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            // retrive data from Database
            string connetionString = null;
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

            string sSQL = "select NAMEPLATE, CAFE_DP, CAFE_LT, CAFE_IP, TARGET_DP, TARGET_LT, TARGET_IP, SPLIT_DP, SPLIT_LT, SPLIT_IP, vol2015, ((CAFE_DP-TARGET_DP)*SPLIT_DP + (CAFE_LT-TARGET_LT)*SPLIT_LT + (CAFE_IP-TARGET_IP)*SPLIT_IP )*vol2015 as credit from dbo.CAFE_FORD_2015 Order by credit";

            SqlCommand cmd = new SqlCommand(sSQL, cnn);

            // cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();
            SqlDataAdapter sqlDataAdap = new SqlDataAdapter(cmd);
            sqlDataAdap.Fill(dt);
            // dataGrid2015.DataSource = dt;
             
            //  Draw Chart here
            foreach (var series in chartFord.Series)
            {
                series.Points.Clear();
            }

            // DataRow[] row;
            int i = 0;

            chartFord.ChartAreas[0].AxisX.Interval = 1;
            chartCredit_Vol.ChartAreas[0].AxisX.Interval = 1;

            foreach (DataRow row in dt.Rows)
            {
                    var creditDP = row["CAFE_DP"];
                    var targetDP = row["TARGET_DP"];

                    var creditLT = row["CAFE_LT"];
                    var targetLT = row["TARGET_LT"];

                    var creditIP = row["CAFE_IP"];
                    var targetIP = row["TARGET_IP"];

                    var vol2015 = row["vol2015"];
                    var splitDP = row["SPLIT_DP"];
                    var splitLT = row["SPLIT_LT"];
                    var splitIP = row["SPLIT_IP"];

                    if (Convert.ToDouble(creditDP) >= 0)
                    {
                        chartFord.Series["DPCredit"].Points.Add(Convert.ToDouble(creditDP) - Convert.ToDouble(targetDP));
                        chartFord.Series["DPCredit"].Points[i].Color = Color.Green;

                        chartFord.Series["DPCredit"].Points[i].AxisLabel = row["NAMEPLATE"].ToString();
                        chartFord.Series["DPCredit"].Points[i].LegendText = row["NAMEPLATE"].ToString();
                    }
                    if (Convert.ToDouble(creditLT) >= 0)
                    {
                        chartFord.Series["LTCredit"].Points.Add(Convert.ToDouble(creditLT) - Convert.ToDouble(targetLT));
                        chartFord.Series["LTCredit"].Points[i].Color = Color.Blue;

                        chartFord.Series["LTCredit"].Points[i].AxisLabel = row["NAMEPLATE"].ToString();
                        chartFord.Series["LTCredit"].Points[i].LegendText = row["NAMEPLATE"].ToString();
                    }
                    if (Convert.ToDouble(creditIP) >= 0)
                    {
                        chartFord.Series["IPCredit"].Points.Add(Convert.ToDouble(creditIP) - Convert.ToDouble(targetIP));
                        chartFord.Series["IPCredit"].Points[i].Color = Color.Red;

                        chartFord.Series["IPCredit"].Points[i].AxisLabel = row["NAMEPLATE"].ToString();
                        chartFord.Series["IPCredit"].Points[i].LegendText = row["NAMEPLATE"].ToString();
                    }

                    if (Convert.ToDouble(creditDP) >= 0)
                    {
                        chartCredit_Vol.Series["DPCredit_Vol"].Points.Add((Convert.ToDouble(creditDP) - Convert.ToDouble(targetDP)) * Convert.ToDouble(vol2015) * Convert.ToDouble(splitDP));
                        chartCredit_Vol.Series["DPCredit_Vol"].Points[i].Color = Color.Green;

                        chartCredit_Vol.Series["DPCredit_Vol"].Points[i].AxisLabel = row["NAMEPLATE"].ToString();
                        chartCredit_Vol.Series["DPCredit_Vol"].Points[i].LegendText = row["NAMEPLATE"].ToString();
                    }
                    if (Convert.ToDouble(creditLT) >= 0)
                    {
                        chartCredit_Vol.Series["LTCredit_Vol"].Points.Add((Convert.ToDouble(creditLT) - Convert.ToDouble(targetLT)) * Convert.ToDouble(vol2015) * Convert.ToDouble(splitLT));
                        chartCredit_Vol.Series["LTCredit_Vol"].Points[i].Color = Color.Blue;

                        chartCredit_Vol.Series["LTCredit_Vol"].Points[i].AxisLabel = row["NAMEPLATE"].ToString();
                        chartCredit_Vol.Series["LTCredit_Vol"].Points[i].LegendText = row["NAMEPLATE"].ToString();
                    }
                    if (Convert.ToDouble(creditIP) >= 0)
                    {
                        chartCredit_Vol.Series["IPCredit_Vol"].Points.Add((Convert.ToDouble(creditIP) - Convert.ToDouble(targetIP)) * Convert.ToDouble(vol2015) * Convert.ToDouble(splitIP));
                        chartCredit_Vol.Series["IPCredit_Vol"].Points[i].Color = Color.Red;

                        chartCredit_Vol.Series["IPCredit_Vol"].Points[i].AxisLabel = row["NAMEPLATE"].ToString();
                        chartCredit_Vol.Series["IPCredit_Vol"].Points[i].LegendText = row["NAMEPLATE"].ToString();
                    }

                    i++;
                
            }

        }

        private void checkDP_CheckedChanged(object sender, EventArgs e)
        {
            if (checkDP.Checked)
            {
                chartFord.Series[0].Enabled = true;
                chartFord.Series[1].Enabled = false;
                chartFord.Series[2].Enabled = false;
                chartCredit_Vol.Series[0].Enabled = true;
                chartCredit_Vol.Series[1].Enabled = false;
                chartCredit_Vol.Series[2].Enabled = false;
                checkLT.Checked = false;
                checkIP.Checked = false;
                checkAll.Checked = false;
            }
        }

        private void checkLT_CheckedChanged(object sender, EventArgs e)
        {
            if (checkLT.Checked)
            {
                chartFord.Series[0].Enabled = false;
                chartFord.Series[1].Enabled = true;
                chartFord.Series[2].Enabled = false;
                chartCredit_Vol.Series[0].Enabled = false;
                chartCredit_Vol.Series[1].Enabled = true;
                chartCredit_Vol.Series[2].Enabled = false;
                checkDP.Checked = false;
                checkIP.Checked = false;
                checkAll.Checked = false;
            }
        }

        private void checkIP_CheckedChanged(object sender, EventArgs e)
        {
            if (checkIP.Checked)
            {
                chartFord.Series[0].Enabled = false;
                chartFord.Series[1].Enabled = false;
                chartFord.Series[2].Enabled = true;
                chartCredit_Vol.Series[0].Enabled = false;
                chartCredit_Vol.Series[1].Enabled = false;
                chartCredit_Vol.Series[2].Enabled = true;
                checkLT.Checked = false;
                checkDP.Checked = false;
                checkAll.Checked = false;
            }
        }

        private void checkAll_CheckedChanged(object sender, EventArgs e)
        {
            if (checkAll.Checked)
            {
                chartFord.Series[0].Enabled = true;
                chartFord.Series[1].Enabled = true;
                chartFord.Series[2].Enabled = true;
                chartCredit_Vol.Series[0].Enabled = true;
                chartCredit_Vol.Series[1].Enabled = true;
                chartCredit_Vol.Series[2].Enabled = true;
                checkLT.Checked = false;
                checkIP.Checked = false;
                checkDP.Checked = false;
            }
        }

        private void checkBox11_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox11.Checked)
            {
                chartFord.ChartAreas[0].Area3DStyle.Enable3D = true;
                chartCredit_Vol.ChartAreas[0].Area3DStyle.Enable3D = true;
            }
            else
            {
                chartFord.ChartAreas[0].Area3DStyle.Enable3D = false;
                chartCredit_Vol.ChartAreas[0].Area3DStyle.Enable3D = false;
            }
        }
    }
}
