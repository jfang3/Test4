using System;
using System.IO;
using System.Data;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Reflection;
using System.Diagnostics;

namespace WindowsFormsApplication2
{
    // Optimization class to optimize the aggegate Ford profit for selected segments
    // Need NAG library
    public class Optimization
    {

        [DllImport("DTW6402DA.dll")]
        public static extern void E04UEF(string string_, int length_string_);

        [DllImport("DTW6402DA.dll")]
        public static extern void E04UFF(ref int irevcm, ref int n, ref int nclin, ref int ncnln,
            ref int lda, ref int ldcj, ref int ldr, double[,] a, double[] bl, double[] bu,
            ref int iter, int[] istate, double[] c, double[,] cjac, double[] clamda, ref double objf,
            double[] objgrd, double[,] r, double[] x, int[] needc, int[] iwork, ref int liwork,
            double[] work, ref int lwork, ref int ifail);

        //      string strNHTSACreditReport = DIR_PATH + "NHTSA_MY_NHTSA_ld_cafe_credit_2008_2014.csv";



        //  private const double epsilon = 0.0; //1.0E-6;
        //  private const double lowerB = 0.0; //100.0;

        private const int numVolumeVar = Routines.numVolumeVar; // # of decision variables for Volumes by transaction type:

        private const int numVariable = numVolumeVar;// = 5;
        //private const double normalFactor0 = 1.0E+6; //10000000.0;	// objective normalization constatnt
        //private  double normalFactor = Routines.timeHorizon * normalFactor0;
        double[] baseVol = Routines.baseVol;


        double normalFactorL = Routines.normalFactor * Routines.timeHorizon; //* 10;

        private static int _displayNAG = 1;//0;	// 1=NAG results display; 0= no display
        public static int displayNAG { get { return _displayNAG; } set { _displayNAG = value; } }

        //Get data from other routines
        int baseYear = Routines.baseYear;
        int timeHorizon = Routines.timeHorizon;


        Constraint[] newFCons = Routines.genFCons;
        Macro[] newMacro = Routines.genPopu;
        //BaseDelta[] baseDelta = Routines.baseDelta;


        int vehTypeLenTot = Routines.vehTypeLenTot;  //=n1+n2+....+nk
        //

        public static Decision[] DecisionOptim;

        //   public static Decision[] shadowPriceEvl;

        int numProdCstr; //the number of production constraints:
        double[] x;      // decison vaiables - retail/lease/rental/fleet's market shares
        //    public static double[] xzero_evl;
        // # of variable, in order of "retail rental, fleet" in volume,
        //    and "adspending1 adspending2 " in variable, 
        // vehicle line 1(year1, year2, ...),vehicle line 2(...), ..., in time
        public Optimization()
        {
            // get the number of selected Ford vehicles
            //int selFordVehLen = Routines.numSelFordVeh;

            //count the number of production constraints:
            numProdCstr = 0;
            for (int i = 0; i < Routines.listOpt.Length; i++)
            {
                numProdCstr += Routines.listOpt[i].Count;
                vehTypeLenTot += Routines.listOpt[i].Count;
            }
            //

            //   int numVar = timeHorizon * selFordVehLen * numVariable;  	// double the variable space on purpose(?)
            x = new double[vehTypeLenTot];							    // init decison vaiables
            for (int i = 0; i < x.Length; i++) x[i] = 1.0;


        }

        // main optimal sub-routine to do the optimaization
        // transaction:		data containing the parameter of nested logit model for each vehicle line
        // fordConstraint:	ford vehicle line constraints
        // hData:			historical data
        // rData:			trading-in data for each vehicle line
        // fData:			data related to FMCC 
        public string OptimalRun()
        {
            Szroutine szFun = new Szroutine();
            VehicleData[][] optdata = new VehicleData[Routines.optData.Length][];//Routines.optData;
            for (int i = 0; i < Routines.optData.Length; i++)
            {
                optdata[i] = szFun.CopyVehData(Routines.optData[i]);
            }

            ArrayList[] listopt = szFun.getOptIndexlist(optdata);

            // set NAG constants
            int n = x.Length;
            // int nclin = (Config.fordVehicle+1)*Config.timeHorizon;	// # of linear constraints, product + unitary
            int nclin = timeHorizon;// numProdCstr;// +timeHorizon * 2 + Routines.AdBgtFlag * 2 * timeHorizon + timeHorizon * 5;		// # of linear constraints, product + CAFE + Ad budget
            int ncnln = 0;// 1;// 0;				// # of nonlinear constraints, none


            int nmax = 2 * n;
            int nclmax = 2 * nclin;
            int ncnmax = 10;


            int iter = 0;			// # of major iterations
            double objf = 0.0;		// value of objective function
            int lda = nclmax;		// 1st dimension of array A(linear constraint), >= max(1,NCLIN)
            int ldcj = ncnmax;		// 1st dimension of array CJAC(nonlinear constraint), >= max(1,NCNLN)
            int ldr = nmax;			// 1st dimension of array R(Cholesky factor), >= N

            double[,] a = new double[nmax, lda];					// linear constraint matrix for variables
            double[] bl = new double[nmax + nclmax + ncnmax];		// low boundary array for constraints
            double[] bu = new double[nmax + nclmax + ncnmax];       // high boundary array for constraints

            double[,] cjac = new double[nmax, ldcj];				// gradient array of nonlinear constraints( d(c[i])/d(x[j]) )
            double[] c = new double[ncnmax];					// values of nonlinear constraints at certain x(i)
            double[] clamda = new double[nmax + nclmax + ncnmax];	// on exit, the values of QP multipliers from the last QP
            double[] objgrd = new double[nmax];					// gradient of objective function at x(i)
            double[,] r = new double[nmax, ldr];					// on exit, contains the upper triangular Cholesky factor
            int[] needc = new int[ncnmax];						// indices of C/cjac that need to be filled
            int liwork = 3 * n + nclin + 2 * ncnln;                 // 500; 8000; WORKING SPACE, >= 3*N + NCLIN + 2*NCNLN
            int lwork = 21 * n + 2;                                   // DEFAULT LWORK
            if (ncnln == 0 && nclin > 0)
            {
                lwork = 2 * n * n + 21 * n + 11 * nclin + 2;        // >=2*N*N + 21*N + 11*NCLIN + 2, if NCLIN>0 and NCNLN=0
            }
            else if (ncnln > 0 && nclin >= 0)
            {
                lwork = 2 * n * n + n * nclin + 2 * n * ncnln + 21 * n + 11 * nclin + 22 * ncnln + 1;   // 2*N*N + N*NCLIN + 2*N*NCNLN + 21*N + 11*NCLIN + 22*NCNLN +1 if NCNLN > 0 and NCLIN >= 0
            }
            int ifail;											// flag for subroutine
            int irevcm;											// indicator for entry/re-entry/intermediate exit/final exit 
            double[] work = new double[lwork];					// work space
            int[] istate = new int[nmax + nclmax + ncnmax];		// states of constraints /may change to n+nclin+ncnln
            int[] iwork = new int[liwork];						// work space


            if ((n > nmax) || (nclin > nclmax) || (ncnln > ncnmax))
            {
                return "Something wrong with NAG dimensions: n/nclin/ncnln.\r\n";
            }

            // Input:   newFCons -- Ford constraints for variables and production limits
            //          baseDelta       -- lifetime 
            // Output:  variable boundary, bu and bl, in volume

            SetBoundary(newFCons, n, bl, bu);


            // nonlinear constraints, none


            // linear constraint coefficient, production
            SetConstraintCoeff(newFCons, n, a);

            //for(int i=0; i<ncnln; i++)		// initialize gradients array
            //	for(int j=0; j<n; j++)	
            //		cjac[i,j]=0.0; 

            // set objective function normalization constant
            //Config.normalFactor =1000000;

            ifail = -1;
            irevcm = 0;

            if (displayNAG == 0)
            {
                E04UEF("Nolist", 6);			// if the NAG window is displayed
                E04UEF("Print Level=0", 13);	//display result details or not
            }
            E04UEF("Derivative Level=0", 18);
            E04UEF("Verify Level=0", 14);
            int major_iter = (int)Math.Max(800, 3 * (n + nclin) + 10 * ncnln);
            E04UEF("Major Iteration Limit = " + major_iter + " ", 28);

            //E04UEF("Major Iteration Limit = 200", 27);
            //E04UEF("Function Precision = 1.0E-12", 28);
            //E04UEF("Optimality Tolerance = 1.0E-10", 30);
            //E04UEF("Line Search Tolerance = 0.5", 27);    //0.1, 0.9

            string stringText = "";			// return text info for display
            int lastIter = -1;              // to detect major iteration

            string oem = Routines.optOEM.ToUpper();
            //     bool bRtnOEMInit = dll_oem_init(oem, oem.Length, Routines.beginYear, Routines.beginYear + Routines.timeHorizon - 1);
            //  int nEvalEPA = 0;               // EPA Evaluation Count
            //   int nEvalEPANeg = 0;            // EPA Evaluation Count (negative returns)
            //    double bestBL = -1.0E+25;
            //    double worstBL = 1.0E+25;
            //     double finalBL = 0;
            var sw = new Stopwatch();

            //  TimeSpan tmSpan = new TimeSpan();

            DateTime startTD = DateTime.Now;

            stringText += "Optimization starts at " + DateTime.Now.ToLongTimeString() + ", " + DateTime.Now.ToShortDateString() + "\r\n";
            do
            {
                E04UFF(ref irevcm, ref n, ref nclin, ref ncnln, ref lda, ref ldcj, ref ldr, a, bl, bu, ref iter, istate,
                    c, cjac, clamda, ref objf, objgrd, r, x, needc, iwork, ref liwork, work, ref lwork, ref ifail);

                if (iter != lastIter)					// every major iteration
                {
                    //stringText += "Iter="+iter+" ";
                    lastIter = iter;
                    //CalculatePopulation(x);
                }

                if ((irevcm > 0) && (ifail == -1))
                {
                    if ((irevcm == 1) || (irevcm == 3))		// calculate value of objective function
                        objf = -objFunction(optdata, x, listopt);	// this is a minimization programmin
                    if ((irevcm == 2) || (irevcm == 3))		// objective gradient
                    { }
                    if ((irevcm == 4) || (irevcm == 6))     // nonlinear constraint
                    {
                        /* 
                         sw.Start();
                         int id = 0;
                         string ptType = "ICE";
                         for (int i = 0; i < Routines.listOpt.Length; i++)
                         {
                             string[] vehName = new string[Routines.listOpt[i].Count];
                             Routines.listOpt[i].CopyTo(vehName);
                             for (int j = 0; j < vehName.Length; j++)
                             {
                                 Regex sl = new Regex(" ");
                                 string[] tmp = sl.Split(vehName[j]);
                                 string brand = tmp[0];
                                 string nameplate = tmp[1];

                                 int nArray = 1;
                                 int[] vol = new Int32[nArray];
                                 for (int k = 0; k < nArray; k++) vol[k] = Convert.ToInt32(x[id] * baseVol[id]);

                                 brand = brand.ToUpper();
                                 nameplate = nameplate.ToUpper();
                                 dll_oem_push_vol(oem, oem.Length, brand, brand.Length, nameplate, nameplate.Length, ptType, ptType.Length, vol, nArray);
                                 id++;
                             }
                           
                         }

                         // GET EPA COMPLIANCE VALUE
                    
                         double cVal = dll_get_EPA_const_val();
                         if (bestBL < cVal) bestBL = cVal;
                         if (worstBL > cVal) worstBL = cVal;
                         finalBL = cVal;
                         c[0] = cVal;
                         nEvalEPA++;
                         if (cVal < 0) nEvalEPANeg++;
                         sw.Stop();
                          */
                    }
                    if ((irevcm == 5) || (irevcm == 6))		// constraint Jacobian
                    {
                    }
                }
            }
            while ((irevcm > 0) && (ifail == -1));



            DateTime stopTD = DateTime.Now;
            stringText += "\r\nOptimization ends at " + DateTime.Now.ToLongTimeString() + ", " + DateTime.Now.ToShortDateString() + "\r\n";
            double duration = (stopTD.Ticks - startTD.Ticks) / 10000000.0;
            stringText += "Duration = " + duration.ToString() + "s.\r\n\r\n";
            stringText += sw.ElapsedMilliseconds / 1000.0 + "s.\r\n\r\n"; ;
            stringText += "Flag i-fail = " + ifail + "\r\n";
            //stringText += "Objective function = " +(-1)*objf +"\r\n";
            stringText += "Total iterations = " + iter + "\r\n";

            double actualProfit = (-1.0) * objf;
            //MessageBox.Show("Total Profit = " + actualProfit);
            stringText += "Objective normalization constant=" + normalFactorL + "\r\n";
            stringText += "The vaule of objective function = " + actualProfit + "\r\n";
            stringText += "Average profit = " + actualProfit * normalFactorL / timeHorizon + "\r\n\r\n";

            DecisionOptim = xToDec(x);

            for (int i = 0; i < x.Length; i++)
                if (Math.Abs(x[i]) < 1.0E-10) x[i] = 0.0;

            stringText += "Decision variables:\r\n";

            stringText += szFun.PrntDec(DecisionOptim);

            string file = Directory.GetCurrentDirectory() + "\\OptimalResult" + ".xls";
            objf = -objFunction(optdata, x, listopt);

            VehicleData[][] result = szFun.preOutData(optdata);
            Routines.exportOptdata(file, result);
            //MessageBox.Show(stringText);
            return stringText;
        }

        //Set lower and upper bound for decision variables and linear constraints
        //  	DSVar -- dimension = Config.timeHorizon*Config.VehLineList * Config.numVariable;
        //							in order:	vehicle line 1(year1, year2, ...), vehicle line 2(...), ...
        //							in each year: retail(Volume), rental(Volume), fleet, 
        //                                      Ad spending (2 variables)
        //      Linear Constraints -- annual production limits
        //      n -- number of total decision variables
        //      fordConstraint -- only include vehicle lines in the selected segment
        //          SAME ORDER as decision variables, sorted by vehicle line, year for the record list


        private double objFunction(VehicleData[][] optData, double[] x, ArrayList[] listOpt)
        {
            double objValue = 0.0;

            // get the number of selected Ford vehicles
            //int selFordVehLen = Routines.numSelFordVeh;

            Decision[] decision = xToDec(x);

            Szroutine szFun = new Szroutine();

            ArrayList listDecision = szFun.getDecIndexlist(decision);
            objValue = szFun.TestCalProfit(optData, decision, listDecision);

            //MessageBox.Show("Total Profit = " + objValue);

            return objValue / normalFactorL;
        }


        private Decision[] xToDec(double[] x)
        {

            Decision[] decision = new Decision[vehTypeLenTot];
            for (int i = 0; i < x.Length; i++)
                if (Math.Abs(x[i]) < 1.0E-10) x[i] = 0.0;

            //Move x[] to class: decision
            int id = 0;

            for (int i = 0; i < Routines.listOpt.Length; i++)
            {
                string[] vehName = new string[Routines.listOpt[i].Count];
                Routines.listOpt[i].CopyTo(vehName);
                for (int j = 0; j < vehName.Length; j++)
                {
                    int tt = i + Routines.beginYear;

                    decision[id] = new Decision();
                    //  decision[id].OEM = oem;
                    decision[id].modelName = vehName[j];
                    decision[id].transYear = tt;
                    decision[id].Volume = x[id] * baseVol[id];
                    id += 1;
                }

            }
            return decision;
        }



        private void SetBoundary(Constraint[] fConstraint, int n, double[] bl, double[] bu)
        {

            // int[] prodgroup_id = new int[numProdCstr];

            int id = 0;
            for (int t = 0; t < timeHorizon; t++)
            {
                //  string[] vehName = new string[Routines.listOpt[t].Count];
                //   Routines.listOpt[t].CopyTo(vehName);
                for (int i = 0; i < Routines.listOpt[t].Count; i++)         //get selected Ford vehicle list.
                {
                    //   bl[id] = 0.0;
                    // bu[id] = 1.0e20;
                    //    int tt = t + Routines.beginYear;
                    //     int cid = Routines.listNewFCons.IndexOf(vehName[i] + tt);

                    //     double minProd = Routines.genFCons[cid].productionMin;//cid].productionMin;
                    //      double maxProd = Routines.genFCons[cid].productionMax;

                    double minProd = Routines.genFCons[i].productionMin;
                    double maxProd = Routines.genFCons[i].productionMax;
                    // if (vehName[i] == "ford f150") MessageBox.Show("f150 boundary");
                    bl[id] = minProd / baseVol[i];
                    bu[id] = maxProd / baseVol[i];
                    id++;
                }


            }
            //set boundary for linear constraint, production max/min
            for (int t = 0; t < timeHorizon; t++)
            {

                bl[id] = 0.0;
                bu[id] = 1.0E20;
                id += 1;
            }

            //set  boundary fo rnonlinear constraint, production max/min
            /*
                        for (int t = 0; t < timeHorizon; t++)
                        {


                            for (int i = 0; i <2 ; i++)         //get selected Ford vehicle list.
                            {



                                bl[id] = 0;
                                bu[id] = 1.0e20;

                                id += 1;
                            }

                        }
                        */

        }



        private void SetConstraintCoeff(Constraint[] fConstraint, int n, double[,] a)
        {
            // linear constraint coefficient, production
            // int numProdCstr = 0;
            //  int idlo = 0;
            // int icol = 0;
            //    for (int i = 0; i < numProdCstr; i++)
            //    {

            //            for (int j = 0; j < n; j++)
            //            {
            //                if (j == i)
            //                    a[j, i] = 1.0;
            ///                else
            //                    a[j, i] = 0.0;
            //           }
            //     icol += 1;

            // idlo += vehTypeLen[i];
            //       }
            //At the end, # of product constraints: numProdCstr = icol 

            //4*timeHorizon linear constraints' coefficients, car & truck CAFE, Ford & L/M AdBudget
            for (int t = 0; t < timeHorizon; t++)
            {
                string[] vehName = new string[Routines.listOpt[t].Count];
                Routines.listOpt[t].CopyTo(vehName);
                int cumCount = 0;
                for (int i = 0; i < cumCount; i++)
                {
                    a[i, t] = 0.0;
                }
                for (int i = cumCount; i < cumCount + Routines.listOpt[t].Count; i++)
                {

                    int tt = t + Routines.beginYear;
                    int cid = Routines.listNewFCons.IndexOf(vehName[i] + tt);
                    double vmt = Routines.genFCons[cid].VMT;
                    double mpg = Routines.genFCons[cid].mpg;
                    double target = Routines.genFCons[cid].target;
                    a[i, t] = -vmt * baseVol[i] * (1.0 / mpg - 1.0 / target);

                }
                for (int i = cumCount + Routines.listOpt[t].Count; i < n; i++)
                {
                    a[i, t] = 0.0;
                }

                cumCount += Routines.listOpt[t].Count;
            }
        }







        /////////////////////////
    }
}