using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Reflection;
using System.Text.RegularExpressions;


namespace WindowsFormsApplication2
{
    public class Routines
    {
        //imported data
        public static Macro[] population;
        public static oeGroupData[] oeGroup;
        public static sgGroupData[] sgGroup;
        public static fuGroupData[] fuGroup;
        public static tGroupData[] tGroup;
        public static lGroupData[] lGroup;
        public static bGroupData[] bGroup;
        public static VehicleData[] vData;
        public static VehicleData[] Scenario;
        public static BaseDelta[] baseDelta;
        public static Constraint[] Constraints;
        public static Refresh[] refresh;


        //gen data
        public static VehicleData[][] genVehData;
        //      public static VehicleData[][] genVehData_Fix;
        public static VehicleData[][] optData; // add by sz on 28/07

        public static Constraint[] genFCons;
        public static Macro[] genPopu;

        //selected data
        //     public static int[] selTransYear;
        public static string optOEM;

        public static double omiga = 0.93;           //discount rate for all
        //  public static int numSeloptVeh = 0;  //add by sz for number of selected ford vehicle lines



        //init default
        private static int _baseYear = 2015;
        private static int _timeHorizon = 1;
        private static int _beginYear = 2016;



        public static double[][] oeshare = new double[Routines.timeHorizon][];
        public static double[][] sgshare = new double[Routines.timeHorizon][];
        public static double[][] fushare = new double[Routines.timeHorizon][];
        public static double[][] tshare = new double[Routines.timeHorizon][];
        public static double[][] lshare = new double[Routines.timeHorizon][];
        public static double[][] bshare = new double[Routines.timeHorizon][];



        public static int recentYear = 1000;
        public static int baseYear { get { return _baseYear; } set { _baseYear = value; } }
        public static int timeHorizon { get { return _timeHorizon; } set { _timeHorizon = value; } }
        public static int beginYear { get { return _beginYear; } set { _beginYear = value; } }



        //excel sheet's name
        public static string[] xlsSheets ={ "macro","bgroup",
            "lgroup","tgroup", "fgroup","sgroup","vehdata","oegroup","basedelta","constraints","refresh","scenario"};//,"planvolume","fordtranstypelist","alltranstypelist", };
        public static string[][] shtNames = new string[xlsSheets.Length][];

        //search key list
        public static ArrayList listVData = new ArrayList();   //oem+name+transyear+type    //vdata
        //  public static ArrayList listPlanVol = new ArrayList();   //oem+name+transyear+type    //volume
        public static ArrayList listScenario = new ArrayList();   //oem+name+transyear+type    
        public static ArrayList listFConstraints = new ArrayList();   //oem+name+transyear         //constraints
        public static ArrayList listRefresh = new ArrayList();   //oem+name+refreshYear         //refresh
        public static ArrayList listLaunch = new ArrayList();   //oem+name+LaunchYear         //refresh
        public static ArrayList listBDelta = new ArrayList();   //oem+name+type              //basedelta
        // public static ArrayList listFMCCData = new ArrayList();   //oem+name                   //FMCCData
        public static ArrayList listPOPU = new ArrayList();   //transyear                  //population  

        //new generated search keylist
        public static listnewvdata[] listNewVData;                  //oem+name+transyear    //genvehdata
        public static ArrayList listNewFCons = new ArrayList();      //oem+name+transyear         //genfCons

        public static ArrayList[] listOpt;    //add by sz


        public static int[] vehTypeLen;
        // public static string[] vehName;
        public static int vehTypeLenTot;

        public const double normalFactor = 3000.0 * 1000.0;
        public const int numVolumeVar = 1;
        //   public const int numVariable = numVolumeVar;// + numAdVar + numBonusVar;
        public static double[] baseVol;







        //export data to excel
        public static void WriteToEXCEL(object[,] rawdata, string[] colField, int sheetNum, string shtName, string fName)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Workbooks books = app.Workbooks;
            Range rng;
            _Workbook book;
            Sheets sheets;
            _Worksheet sheet;

            if (!File.Exists(@fName))
            {
                book = books.Add(Type.Missing);
                book.SaveAs(@fName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            book = app.Workbooks.Open(@fName, Type.Missing,
                false, true, Type.Missing, Type.Missing, true, Type.Missing,
                Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            sheets = book.Worksheets;
            if (sheetNum > sheets.Count)
            {
                sheet = (Microsoft.Office.Interop.Excel.Worksheet)
                       book.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                sheet.Name = shtName;
            }
            else
            {
                sheet = (_Worksheet)sheets.get_Item(sheetNum);
                sheet.Name = shtName;
            }
            //   ((Microsoft.Office.Interop.Excel._Worksheet)books[1]).Activate();
            sheet.Columns.ClearContents();

            object[] Headers = colField;
            string lastCol = GetRangeLetter(colField.Length) + 1;
            rng = sheet.get_Range("A1", lastCol);
            rng.Value2 = Headers;

            rng = sheet.get_Range("A2", Type.Missing);
            rng = rng.get_Resize(rawdata.Length / colField.Length, colField.Length);
            rng.Value2 = rawdata;

            books[1].RefreshAll();
            foreach (Microsoft.Office.Interop.Excel.Workbook b in app.Workbooks) { b.Save(); b.Close(true, Type.Missing, Type.Missing); }
            books.Close();
            app.Quit();
        }

        //export profit analyse

        //import data from Excel file
        public static void importFile(string infile, int flag)
        {
            // ApplicationClass app = new ApplicationClass();
            // flag = 1:  infile has directory path;   flag = 2:   infile has n o directory path
            string fileName;

            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            if (flag == 1)
                fileName = Directory.GetCurrentDirectory() + "\\" + infile;
            else
                fileName = infile;

            if (!File.Exists(fileName))
            {
                MessageBox.Show("The input Excel data file does not exist! Please check.", "Warning");
                return;
            }
            Workbook wb = app.Workbooks.Open(@fileName, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Sheets shts = wb.Worksheets;



            listPOPU = new ArrayList();
            listVData = new ArrayList();
            listBDelta = new ArrayList();
            listFConstraints = new ArrayList();
            listRefresh = new ArrayList();
            listLaunch = new ArrayList();
            foreach (Worksheet sht in shts)
            {
                int nRow, nCol;
                string endRange;
                ((Microsoft.Office.Interop.Excel._Worksheet)sht).Activate();
                string shtName = sht.Name.ToString();
                nRow = 1;
                nCol = 156;
                endRange = GetRangeLetter(nCol) + nRow.ToString();
                Range rng = app.get_Range("A1", endRange);
                object[,] rdata = (object[,])rng.get_Value(Missing.Value);

                // column headings and nCol_
                int nCol_ = 0;
                ArrayList flds = new ArrayList();
                for (int i = 0; i < nCol; i++)
                {
                    string str;
                    try
                    {
                        str = Convert.ToString(rdata[1, i + 1]);
                        if (str.Replace(" ", "") == "") break;
                    }
                    catch { break; }
                    nCol_ = i + 1;
                    str = str.Trim().ToLower().Replace(" ", "");
                    flds.Add(str);
                }

                nRow = 65535;
                nCol = nCol_;
                endRange = GetRangeLetter(nCol) + nRow.ToString();
                rng = app.get_Range("A2", endRange);
                rdata = (System.Object[,])rng.get_Value(Missing.Value);

                // nRow_
                int nRow_ = 0;
                for (int i = 1; i < nRow; i++)
                {
                    try
                    {
                        string str = Convert.ToString(rdata[i, 1]);
                        if (str.Replace(" ", "") == "") break;
                    }
                    catch { break; }
                    nRow_ = i;
                }
                // ConfigName[idx]
                int idx = Array.IndexOf(xlsSheets, shtName.ToLower());
                if ((nCol_ > 0 && nRow_ > 0) && idx > -1)
                {
                    shtNames[idx] = new string[nCol];
                    flds.CopyTo(shtNames[idx]);
                }

                switch (shtName.ToLower())
                {
                    case "macro":
                        population = new Macro[nRow_];
                        listPOPU = new ArrayList();
                        for (int i = 0; i < nRow_; i++)
                        {
                            population[i] = new Macro();
                            population[i].transYear = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("transactionyear")]);
                            population[i].population = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("population")]);
                            listPOPU.Add(population[i].transYear);
                        }
                        break;
                    case "bgroup":
                        bGroup = new bGroupData[nRow_];
                        for (int i = 0; i < nRow_; i++)
                        {
                            bGroup[i] = new bGroupData();
                            bGroup[i].bGroupID = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("bgroupid")]);
                            bGroup[i].sigma = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("sigma")]);

                        }
                        break;
                    case "lgroup":
                        lGroup = new lGroupData[nRow_];
                        for (int i = 0; i < nRow_; i++)
                        {
                            lGroup[i] = new lGroupData();
                            lGroup[i].lGroupID = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("lgroupid")]);
                            lGroup[i].rho = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("rho")]);
                        }
                        break;
                    case "tgroup":
                        tGroup = new tGroupData[nRow_];
                        for (int i = 0; i < nRow_; i++)
                        {
                            tGroup[i] = new tGroupData();
                            tGroup[i].tGroupID = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("tgroupid")]);
                            tGroup[i].phi = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("phi")]);
                        }
                        break;

                    case "fgroup":
                        fuGroup = new fuGroupData[nRow_];
                        for (int i = 0; i < nRow_; i++)
                        {
                            fuGroup[i] = new fuGroupData();
                            fuGroup[i].fuGroupID = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("fugroupid")]);
                            fuGroup[i].theta = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("theta")]);
                        }
                        break;
                    case "sgroup":
                        sgGroup = new sgGroupData[nRow_];
                        for (int i = 0; i < nRow_; i++)
                        {
                            sgGroup[i] = new sgGroupData();

                            sgGroup[i].sgroupID = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("sggroupid")]);
                            sgGroup[i].alpha = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("alpha")]);
                            sgGroup[i].tau = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("tau")]);
                        }
                        break;
                    case "oegroup":
                        oeGroup = new oeGroupData[nRow_];
                        for (int i = 0; i < nRow_; i++)
                        {
                            oeGroup[i] = new oeGroupData();
                            oeGroup[i].oeGroupID = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("oegroupid")]);
                            oeGroup[i].lambda = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("lambda")]);
                        }
                        break;
                    case "vehdata":
                        vData = new VehicleData[nRow_];
                        listVData = new ArrayList();
                        for (int i = 0; i < nRow_; i++)
                        {
                            vData[i] = new VehicleData();
                            vData[i].group = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("group")]).ToLower();
                            vData[i].type = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("type")]).ToLower();
                            vData[i].fuel_type = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("fuel_type")]).ToLower();
                            vData[i].segment = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("segment")]).ToLower();
                            vData[i].OEM = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("oem")]).ToLower();
                            vData[i].modelName = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("modname")]).ToLower();
                            vData[i].modelYear = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("modelyear")]);
                            vData[i].transYear = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("year")]);
                            vData[i].oeGroupID = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("oegroupid")]);
                            vData[i].sgGroupID = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("sggroupid")]);
                            vData[i].fuGroupID = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("fugroupid")]);
                            vData[i].tGroupID = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("tgroupid")]);
                            vData[i].lGroupID = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("lgroupid")]);
                            vData[i].bGroupID = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("bgroupid")]);
                            vData[i].price = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("price")]) / 10000.0;
                            vData[i].volume = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("volume")]);
                            if (vData[i].transYear > recentYear) recentYear = vData[i].transYear;
                            listVData.Add(vData[i].modelName + vData[i].transYear);
                        }
                        break;
                    case "scenario":
                        Scenario = new VehicleData[nRow_];
                        listScenario = new ArrayList();
                        for (int i = 0; i < nRow_; i++)
                        {
                            Scenario[i] = new VehicleData();
                            Scenario[i].group = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("group")]).ToLower();
                            Scenario[i].type = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("type")]).ToLower();
                            Scenario[i].fuel_type = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("fuel_type")]).ToLower();
                            Scenario[i].segment = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("segment")]).ToLower();
                            Scenario[i].OEM = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("oem")]).ToLower();
                            Scenario[i].modelName = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("modname")]).ToLower();
                            Scenario[i].modelYear = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("modelyear")]);
                            Scenario[i].transYear = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("year")]);
                            Scenario[i].oeGroupID = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("oegroupid")]);
                            Scenario[i].sgGroupID = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("sggroupid")]);
                            Scenario[i].fuGroupID = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("fugroupid")]);
                            Scenario[i].tGroupID = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("tgroupid")]);
                            Scenario[i].lGroupID = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("lgroupid")]);
                            Scenario[i].bGroupID = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("bgroupid")]);
                            Scenario[i].price = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("price")]) / 10000.0;
                            Scenario[i].volume = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("volume")]);
                            listScenario.Add(Scenario[i].modelName + Scenario[i].transYear);
                        }
                        break;
                    case "basedelta":
                        baseDelta = new BaseDelta[nRow_];
                        listBDelta = new ArrayList();
                        for (int i = 0; i < nRow_; i++)
                        {
                            baseDelta[i] = new BaseDelta();
                            baseDelta[i].OEM = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("oem")]);
                            // baseDelta[i].segment = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("segment")]);
                            baseDelta[i].modelName = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("modname")]).ToLower();
                            baseDelta[i].styleAgeDep = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("styleagedep")]);
                            baseDelta[i].majImpact = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("majimpact")]);
                            baseDelta[i].majStd = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("majstd")]);
                            baseDelta[i].adjFactor = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("adjfactor")]);
                            //    baseDelta[i].ddelta = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("ddelta")]);


                            listBDelta.Add(baseDelta[i].modelName);
                        }
                        break;


                    case "refresh":
                        refresh = new Refresh[nRow_];
                        listRefresh = new ArrayList();
                        listLaunch = new ArrayList();
                        for (int i = 0; i < nRow_; i++)
                        {
                            refresh[i] = new Refresh();
                            refresh[i].OEM = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("oem")]);
                            refresh[i].modelName = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("name")]);
                            refresh[i].refreshYear = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("refreshyear")]);
                            refresh[i].launchYear = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("launchyear")]);
                            refresh[i].successRate = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("successrate")]);
                            refresh[i].incVolRate = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("increasedvolrate")]);
                            listRefresh.Add(refresh[i].modelName + refresh[i].refreshYear);
                            listLaunch.Add(refresh[i].modelName + refresh[i].launchYear);
                        }
                        break;

                    case "constraints":
                        Constraints = new Constraint[nRow_];
                        listFConstraints = new ArrayList();
                        for (int i = 0; i < nRow_; i++)
                        {
                            Constraints[i] = new Constraint();
                            Constraints[i].OEM = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("oem")]).ToLower();
                            Constraints[i].ModelName = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("modname")]).ToLower();
                            //Constraints[i].ModelYear = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("year")]);
                            Constraints[i].ModelYear = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("modelyear")]);
                            Constraints[i].productionMin = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("prodmin")]);
                            Constraints[i].productionMax = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("prodmax")]);
                            Constraints[i].variableCost = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("variablecost")]) / 10000.0;
                            Constraints[i].footprint = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("footprint")]);
                            Constraints[i].mpg = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("mpg")]);
                            Constraints[i].target = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("target")]);
                            Constraints[i].category = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("category")]);
                            Constraints[i].engType = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("engtype")]);
                            Constraints[i].VMT = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("vmt")]);
                            listFConstraints.Add(Constraints[i].ModelName + Constraints[i].ModelYear);
                        }
                        break;

                    default:
                        break;
                }
            }
            wb.Close(false, Type.Missing, Type.Missing);
            app.Workbooks.Close();
            app.Quit();

            //    segmentList = new string[segmentlist.Count]; segmentlist.Sort(); segmentlist.CopyTo(segmentList);
            //    oemList = new string[oemlist.Count]; oemlist.Sort(); oemlist.CopyTo(oemList);
            //    fordVehList = new string[fordvehlist.Count]; fordvehlist.Sort(); fordvehlist.CopyTo(fordVehList);
            //    otherVehList = new string[othervehlist.Count]; othervehlist.Sort(); othervehlist.CopyTo(otherVehList);
            //segOemName = new string[segoemname.Count]; segoemname.Sort(); segoemname.CopyTo(segOemName);
        }

        //import edited data from Excel file, All arraylists are headings

        public static string GetRangeLetter(int range)
        {
            string[] letterList ={"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M",
                                  "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
                                  "AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM",
                                  "AN","AO","AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ",
                                  "BA","BB","BC","BD","BE","BF","BG","BH","BI","BJ","BK","BL","BM",
                                  "BN","BO","BP","BQ","BR","BS","BT","BU","BV","BW","BX","BY","BZ",
                                  "BA","CB","CC","CD","CE","CF","CG","CH","CI","CJ","CK","CL","CM",
                                  "CN","CO","CP","CQ","CR","CS","CT","CU","CV","CW","CX","CY","CZ",
                                  "DA","DB","DC","DD","DE","DF","DG","DH","DI","DJ","DK","DL","DM",
                                  "DN","DO","DP","DQ","DR","DS","DT","DU","DV","DW","DX","DY","DZ",
                                  "EA","EB","EC","ED","EE","EF","EG","EH","EI","EJ","EK","EL","EM",
                                  "EN","EO","EP","EQ","ER","ES","ET","EU","EV","EW","EX","EY","EZ"};
            if (range < 1) return null;
            else return letterList[range - 1];
        }
        //view excel file
        public static void viewEXCEL(string filename)
        {
            string str = Directory.GetCurrentDirectory();
            if (!filename.Contains(str)) filename = Directory.GetCurrentDirectory() + "\\" + filename;
            System.Diagnostics.Process proc = new System.Diagnostics.Process();
            proc.EnableRaisingEvents = false;
            proc.StartInfo.FileName = "excel";
            proc.StartInfo.FileName = filename;
            proc.Start();
        }

        //export updated input data

        /*    public static void loadUPdatedDec(string fileName, Decision[] decision)
            {
                if (decision == null)
                {
                    MessageBox.Show("Decision does not exit!, Please Check!");
                    return;
                }
                Microsoft.Office.Interop.Excel.ApplicationClass app = new Microsoft.Office.Interop.Excel.ApplicationClass();
                if (!File.Exists(fileName))
                {
                    MessageBox.Show("The input Excel data file does not exist! Please check.", "Warning");
                    return;
                }
                Workbook wb = app.Workbooks.Open(@fileName, Type.Missing,
                    true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                Sheets shts = wb.Worksheets;

                foreach (Worksheet sht in shts)
                {
                    if (sht.Name.ToLower() != "decision") continue;
                    int nRow, nCol;
                    string endRange;
                    ((Microsoft.Office.Interop.Excel._Worksheet)sht).Activate();
                    string shtName = sht.Name.ToString();
                    nRow = 1;
                    nCol = 156;
                    endRange = GetRangeLetter(nCol) + nRow.ToString();
                    Range rng = app.get_Range("A1", endRange);
                    object[,] rdata = (object[,])rng.get_Value(Missing.Value);

                    // column headings and nCol_
                    int nCol_ = 0;
                    ArrayList flds = new ArrayList();
                    for (int i = 0; i < nCol; i++)
                    {
                        string str;
                        try
                        {
                            str = Convert.ToString(rdata[1, i + 1]);
                            if (str.Replace(" ", "") == "") break;
                        }
                        catch { break; }
                        nCol_ = i + 1;
                        str = str.Trim().ToLower().Replace(" ", "");
                        flds.Add(str);
                    }
                    // get data

                    nRow = 65535;
                    nCol = nCol_;
                    endRange = GetRangeLetter(nCol) + nRow.ToString();
                    rng = app.get_Range("A2", endRange);
                    rdata = (System.Object[,])rng.get_Value(Missing.Value);

                    // get nRow_
                    int nRow_ = 0;
                    for (int i = 1; i < nRow; i++)
                    {
                        try
                        {
                            string str = Convert.ToString(rdata[i, 1]);
                            if (str.Replace(" ", "") == "") break;
                        }
                        catch { break; }
                        nRow_ = i;
                    }

                    switch (shtName.ToLower())
                    {

                        case "decision":
                            if (decision.Length != (nRow_))
                            {
                                MessageBox.Show("something is wrong!");
                                return;
                            }
                            for (int i = 0; i < nRow_; i++)
                            {

                                // decision[i] = new Decision();
                                decision[i].OEM = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("oem")]);
                                decision[i].modelName = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("modelname")]); ;
                                decision[i].transYear = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("transyear")]);
                                decision[i].retailVolume = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("retailvolume")]);
                                decision[i].rentalVolume = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("rentalvolume")]);
                                decision[i].rentalriskVolume = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("rentalriskvolume")]);
                                decision[i].fleetVolume = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("fleetvolume")]);
                                decision[i].adspBrand = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("adspbrand")]);
                                decision[i].adspRetail = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("adspretail")]);


                            }
                            break;

                        default:
                            break;
                    }
                }
                wb.Close(false, Type.Missing, Type.Missing);
                app.Workbooks.Close();
                app.Quit();

            }
    */
        //export profit analyse
        public static void exportDecision(string filename, Decision[] dec)
        {
            string dataDir = Directory.GetCurrentDirectory();
            if (File.Exists(filename))
            {
                Regex sl = new Regex("\\\\");
                string[] tmp = sl.Split(filename);
                File.Move(filename, dataDir + "\\Data\\" + tmp[tmp.Length - 1].Remove(tmp[tmp.Length - 1].Length - 4) +
                    DateTime.Now.Day + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second + ".xls");
                File.Copy(dataDir + "\\empty.xls", filename);
            }
            else File.Copy(dataDir + "\\empty.xls", filename);

            string[] strHeader = new string[9];
            strHeader[0] = "OEM"; strHeader[1] = "modelName";
            strHeader[2] = "transYear"; strHeader[3] = "retailVolume";


            int totrow = dec.Length;
            //for (int i = 0; i < getProfits.Length; i++) totrow += getProfits[i].Length;

            object[,] objR = new object[totrow, strHeader.Length];

            for (int i = 0; i < dec.Length; i++)
            {
                objR[i, 0] = dec[i].OEM;
                objR[i, 1] = dec[i].modelName;
                objR[i, 2] = dec[i].transYear;
                objR[i, 3] = dec[i].Volume;
            }

            WriteToEXCEL(objR, strHeader, 1, "Decision", filename);
        }


        public static void exportOptdata(string filename, VehicleData[][] optData)
        {
            //MessageBox.Show(filename);
            string str = "";
            string dataDir = Directory.GetCurrentDirectory();
            string[] strHeader = new string[10];
            strHeader[0] = "Group"; strHeader[1] = "Type";
            strHeader[2] = "Fuel_Type"; strHeader[3] = "Segment";
            strHeader[4] = "Oem"; strHeader[5] = "Model"; strHeader[6] = "year";
            strHeader[7] = "volume"; strHeader[8] = "price";
            strHeader[9] = "Elasticity";

            int totrow = 0;
            for (int i = 0; i < optData.Length; i++)
            {
                for (int j = 0; j < optData[i].Length; j++)
                {
                    totrow++;
                }
            }

            object[,] objR = new object[totrow, strHeader.Length];
            int k = 0;

            for (int i = 0; i < optData.Length; i++)
            {
                for (int j = 0; j < optData[i].Length; j++)
                {
                    objR[k, 0] = optData[i][j].group.ToUpper();
                    objR[k, 1] = optData[i][j].type.ToUpper();
                    objR[k, 2] = optData[i][j].fuel_type.ToUpper();
                    objR[k, 3] = optData[i][j].segment.ToUpper();
                    objR[k, 4] = optData[i][j].OEM.ToUpper();
                    objR[k, 5] = optData[i][j].modelName.ToUpper();
                    objR[k, 6] = optData[i][j].transYear;
                    objR[k, 7] = optData[i][j].volume;
                    objR[k, 8] = optData[i][j].price;
                    objR[k, 9] = optData[i][j].Elasticity;
                    k++;
                }
            }
            //MessageBox.Show("OK");
            WriteToEXCEL(objR, strHeader, 1, "optData", filename);

            //write to a temporary csv file
            for (int i = 0; i < k; i++)
            {
                for (int j = 0; j < strHeader.Length; j++)
                {
                    if (j < strHeader.Length - 1)
                    {
                        str += objR[i, j].ToString() + ",";
                    }
                    else
                    {
                        str += objR[i, j].ToString() + Environment.NewLine;
                    }
                }
            }
            string csvfname = dataDir + @"\OptimalResult.csv";
            // MessageBox.Show(csvfname);
            using (StreamWriter writer = new StreamWriter(csvfname))
            {
                writer.Write(str);
            }

        }

        public static void loadUPdatedOpt(string fileName, VehicleData[][] optData, ArrayList[] listOpt)
        {
            if (optData == null)
            {
                MessageBox.Show("Decision does not exit!, Please Check!");
                return;
            }
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            if (!File.Exists(fileName))
            {
                MessageBox.Show("The input Excel data file does not exist! Please check.", "Warning");
                return;
            }
            Workbook wb = app.Workbooks.Open(@fileName, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Sheets shts = wb.Worksheets;

            foreach (Worksheet sht in shts)
            {
                if (sht.Name.ToLower() != "optdata" && sht.Name.ToLower() != "vehdata") continue;
                int nRow, nCol;
                string endRange;
                //   ((Microsoft.Office.Interop.Excel._Worksheet)sht).Activate();
                string shtName = sht.Name.ToString();
                nRow = 1;
                nCol = 156;
                endRange = GetRangeLetter(nCol) + nRow.ToString();
                Range rng = app.get_Range("A1", endRange);
                object[,] rdata = (object[,])rng.get_Value(Missing.Value);

                // column headings and nCol_
                int nCol_ = 0;
                ArrayList flds = new ArrayList();
                for (int i = 0; i < nCol; i++)
                {
                    string str;
                    try
                    {
                        str = Convert.ToString(rdata[1, i + 1]);
                        if (str.Replace(" ", "") == "") break;
                    }
                    catch { break; }
                    nCol_ = i + 1;
                    str = str.Trim().ToLower().Replace(" ", "");
                    flds.Add(str);
                }
                // get data

                nRow = 65535;
                nCol = nCol_;
                endRange = GetRangeLetter(nCol) + nRow.ToString();
                rng = app.get_Range("A2", endRange);
                rdata = (System.Object[,])rng.get_Value(Missing.Value);

                // get nRow_
                int nRow_ = 0;
                for (int i = 1; i < nRow; i++)
                {
                    try
                    {
                        string str = Convert.ToString(rdata[i, 1]);
                        if (str.Replace(" ", "") == "") break;
                    }
                    catch { break; }
                    nRow_ = i;
                }

                switch (shtName.ToLower())
                {

                    case "optdata":
                    case "vehdata":

                        for (int i = 0; i < nRow_; i++)
                        {

                            int k = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("year")]) - Routines.baseYear - 1;
                            if (k < 0) k = 0;
                            string oem = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("oem")]).ToLower();
                            string modelname = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("modname")]).ToLower();
                            int j = listOpt[k].IndexOf(modelname);

                            optData[k][j].group = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("group")]).ToLower();
                            optData[k][j].type = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("type")]).ToLower();
                            optData[k][j].fuel_type = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("fuel_type")]).ToLower();
                            optData[k][j].segment = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("segment")]).ToLower();
                            optData[k][j].OEM = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("oem")]).ToLower();
                            optData[k][j].modelName = Convert.ToString(rdata[i + 1, 1 + flds.IndexOf("modname")]).ToLower();
                            optData[k][j].modelYear = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("modelyear")]);
                            optData[k][j].transYear = Convert.ToInt32(rdata[i + 1, 1 + flds.IndexOf("year")]);
                            optData[k][j].price = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("price")]) / 10000.0;
                            optData[k][j].volume = Convert.ToDouble(rdata[i + 1, 1 + flds.IndexOf("volume")]);
                        }
                        break;

                    default:
                        break;
                }
            }
            wb.Close(false, Type.Missing, Type.Missing);
            app.Workbooks.Close();
            app.Quit();

        }
        public static void exportVehdata(string filename, VehicleData[][] optData)
        {
            string dataDir = Directory.GetCurrentDirectory();
            if (File.Exists(filename))
            {
                Regex sl = new Regex("\\\\");
                string[] tmp = sl.Split(filename);
                File.Move(filename, dataDir + "\\Data\\" + tmp[tmp.Length - 1].Remove(tmp[tmp.Length - 1].Length - 4) +
                    DateTime.Now.Day + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second + ".xls");
                File.Copy(dataDir + "\\empty.xls", filename);
            }
            else File.Copy(dataDir + "\\empty.xls", filename);

            string[] strHeader = new string[17];
            strHeader[0] = "Group"; strHeader[1] = "Type";
            strHeader[2] = "Fuel_Type"; strHeader[3] = "Segment";
            strHeader[4] = "Oem"; strHeader[5] = "Model"; strHeader[6] = "year";
            strHeader[7] = "volume"; strHeader[8] = "price";
            strHeader[9] = "lGroupId";
            strHeader[10] = "tGroupId";
            strHeader[11] = "fuGroupId";
            strHeader[12] = "sgGroupId";
            strHeader[13] = "oeGroupId";
            strHeader[14] = "bGroupId";
            strHeader[15] = "modelyear";
            strHeader[16] = "Elasticity";

            int totrow = 0;
            for (int i = 0; i < optData.Length; i++)
            {
                for (int j = 0; j < optData[i].Length; j++)
                {
                    totrow++;
                }
            }

            object[,] objR = new object[totrow, strHeader.Length];
            int k = 0;
            for (int i = 0; i < optData.Length; i++)
            {
                for (int j = 0; j < optData[i].Length; j++)
                {
                    objR[k, 0] = optData[i][j].group.ToUpper();
                    objR[k, 1] = optData[i][j].type.ToUpper();
                    objR[k, 2] = optData[i][j].fuel_type.ToUpper();
                    objR[k, 3] = optData[i][j].segment.ToUpper();
                    objR[k, 4] = optData[i][j].OEM.ToUpper();
                    objR[k, 5] = optData[i][j].modelName.ToUpper();
                    objR[k, 6] = optData[i][j].transYear;
                    objR[k, 7] = optData[i][j].volume;
                    objR[k, 8] = optData[i][j].price;
                    objR[k, 9] = optData[i][j].lGroupID;
                    objR[k, 10] = optData[i][j].tGroupID;
                    objR[k, 11] = optData[i][j].fuGroupID;
                    objR[k, 12] = optData[i][j].sgGroupID;
                    objR[k, 13] = optData[i][j].oeGroupID;
                    objR[k, 14] = optData[i][j].bGroupID;
                    objR[k, 15] = optData[i][j].modelYear;
                    objR[k, 16] = optData[i][j].Elasticity;
                    k++;
                }
            }
            WriteToEXCEL(objR, strHeader, 1, "optData", filename);
        }
        /////////////////////////////////////////////////////////////////////////////////
    }
}
