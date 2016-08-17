using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Data;
using System.Collections;
using System.Windows.Forms;
using System.Reflection;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;


namespace WindowsFormsApplication2
{
    public class Szroutine
    {

        public VehicleData[][] prefutureData(VehicleData[] histdata, int timeHorizon, VehicleData[] scenario)
        {
            // generate n years data in turn
            VehicleData[][] fData = new VehicleData[timeHorizon][];
            VehicleData[] basedata = getBYData(histdata, Routines.baseYear);
            getDelta_new(basedata, Routines.baseYear, Routines.baseDelta);
            fData[0] = getNextYearData(basedata, scenario);
            for (int i = 1; i < Routines.timeHorizon; i++)
                fData[i] = getNextYearData(fData[i - 1], scenario);
            return fData;
        }

        private VehicleData[] getBYData(VehicleData[] vdata, int transYear)
        {

            int datalength = 0;
            for (int i = 0; i < vdata.Length; i++)
            {
                if (vdata[i].transYear == Routines.baseYear)
                {
                    datalength++;
                }
            }
            VehicleData[] BaseData = new VehicleData[datalength];
            int j = 0;
            for (int i = 0; i < vdata.Length; i++)
            {
                if (vdata[i].transYear == Routines.baseYear)
                {
                    BaseData[j] = new VehicleData();
                    BaseData[j].group = vdata[i].group;
                    BaseData[j].type = vdata[i].type;
                    BaseData[j].fuel_type = vdata[i].fuel_type;
                    BaseData[j].segment = vdata[i].segment;
                    BaseData[j].OEM = vdata[i].OEM;
                    BaseData[j].modelName = vdata[i].modelName;
                    BaseData[j].modelYear = vdata[i].modelYear;
                    BaseData[j].oeGroupID = vdata[i].oeGroupID;
                    BaseData[j].sgGroupID = vdata[i].sgGroupID;
                    BaseData[j].fuGroupID = vdata[i].fuGroupID;
                    BaseData[j].tGroupID = vdata[i].tGroupID;
                    BaseData[j].lGroupID = vdata[i].lGroupID;
                    BaseData[j].bGroupID = vdata[i].bGroupID;
                    BaseData[j].transYear = transYear;
                    BaseData[j].volume = vdata[i].volume;
                    BaseData[j].price = vdata[i].price;
                    j++;
                }
            }
            return BaseData;
        }

        //generate vehicle data acording to last year data
        public VehicleData[] getNextYearData(VehicleData[] baseData, VehicleData[] scenario)
        {
            int transyear = baseData[0].transYear + 1;
            ArrayList listbsdata = new ArrayList();
            int baselength = baseData.Length;
            for (int i = 0; i < baseData.Length; i++)
            {
                listbsdata.Add(baseData[i].modelName + Convert.ToString(baseData[i].transYear + 1));
            }
            for (int i = 0; i < scenario.Length; i++)
            {
                if (scenario[i].transYear == transyear)
                {
                    int idx = listbsdata.IndexOf(scenario[i].modelName + scenario[i].transYear);
                    if (idx != -1)
                    {
                        if (scenario[i].volume <= 0)
                            baselength--;
                    }
                    else
                    {
                        if (scenario[i].transYear == transyear && scenario[i].volume > 0)
                        {
                            baselength++;

                        }
                    }
                }
            }

            VehicleData[] rData = new VehicleData[baselength];
            int k = 0;
            for (int i = 0; i < baseData.Length; i++)
            {
                int idxscenario = Routines.listScenario.IndexOf(baseData[i].modelName + transyear);
                if (idxscenario == -1)
                {
                    rData[k] = new VehicleData();
                    rData[k].group = baseData[i].group;
                    rData[k].type = baseData[i].type;
                    rData[k].fuel_type = baseData[i].fuel_type;
                    rData[k].segment = baseData[i].segment;
                    rData[k].OEM = baseData[i].OEM;
                    rData[k].modelName = baseData[i].modelName;
                    rData[k].transYear = baseData[i].transYear + 1;
                    rData[k].modelYear = baseData[i].modelYear + 1;
                    rData[k].volume = baseData[i].volume;
                    rData[k].oeGroupID = baseData[i].oeGroupID;
                    rData[k].sgGroupID = baseData[i].sgGroupID;
                    rData[k].fuGroupID = baseData[i].fuGroupID;
                    rData[k].tGroupID = baseData[i].tGroupID;
                    rData[k].lGroupID = baseData[i].lGroupID;
                    rData[k].bGroupID = baseData[i].bGroupID;
                    rData[k].price = baseData[i].price;

                    string key = rData[k].OEM + rData[k].modelName + rData[k].modelYear;
                    int idR = Routines.listRefresh.IndexOf(key);
                    if (idR != -1)
                        rData[k].volume *= (1 + Routines.refresh[idR].incVolRate);
                    k++;
                }
                else
                {
                    if (scenario[idxscenario].volume > 0)
                    {
                        rData[k] = new VehicleData();
                        rData[k].group = baseData[i].group;
                        rData[k].type = baseData[i].type;
                        rData[k].fuel_type = baseData[i].fuel_type;
                        rData[k].segment = baseData[i].segment;
                        rData[k].OEM = baseData[i].OEM;
                        rData[k].modelName = baseData[i].modelName;
                        rData[k].transYear = baseData[i].transYear + 1;
                        rData[k].modelYear = baseData[i].modelYear + 1;
                        rData[k].oeGroupID = baseData[i].oeGroupID;
                        rData[k].sgGroupID = baseData[i].sgGroupID;
                        rData[k].fuGroupID = baseData[i].fuGroupID;
                        rData[k].tGroupID = baseData[i].tGroupID;
                        rData[k].lGroupID = baseData[i].lGroupID;
                        rData[k].bGroupID = baseData[i].bGroupID;
                        rData[k].price = scenario[idxscenario].price;
                        rData[k].volume = scenario[idxscenario].volume;
                        k++;
                    }
                }
            }
            for (int i = 0; i < scenario.Length; i++)
            {
                int idxbsdata = listbsdata.IndexOf(scenario[i].modelName + scenario[i].transYear);
                if (idxbsdata == -1 && scenario[i].transYear == transyear && scenario[i].volume > 0)
                {

                    rData[k] = new VehicleData();
                    rData[k].OEM = scenario[i].OEM;
                    rData[k].segment = scenario[i].segment;
                    rData[k].modelName = scenario[i].modelName;
                    rData[k].type = scenario[i].type;
                    rData[k].transYear = scenario[i].transYear;
                    rData[k].modelYear = scenario[i].modelYear;
                    rData[k].volume = scenario[i].volume;
                    rData[k].oeGroupID = scenario[i].oeGroupID;
                    rData[k].sgGroupID = scenario[i].sgGroupID;
                    rData[k].fuGroupID = scenario[i].fuGroupID;
                    rData[k].tGroupID = scenario[i].tGroupID;
                    rData[k].lGroupID = scenario[i].lGroupID;
                    rData[k].bGroupID = scenario[i].bGroupID;
                    rData[k].price = scenario[i].price;

                    k++;
                }
            }

            return rData;
        }

        //Calculate other vehicle lines marketing share 
        public void CalMktShare(VehicleData[][] futuredata)
        {
            for (int i = 0; i < Routines.timeHorizon; i++)
            {

                int Pkey = Routines.beginYear + i;
                int idp = Routines.listPOPU.IndexOf(Pkey);
                //calculate non-Optimal vehicles (competitors') mGroup base share 
                Routines.oeshare[i] = new double[Routines.oeGroup.Length];
                Routines.sgshare[i] = new double[Routines.sgGroup.Length];
                Routines.fushare[i] = new double[Routines.fuGroup.Length];
                Routines.tshare[i] = new double[Routines.tGroup.Length];
                Routines.lshare[i] = new double[Routines.lGroup.Length];
                Routines.bshare[i] = new double[Routines.bGroup.Length];
                for (int j = 0; j < futuredata[i].Length; j++)
                {
                    futuredata[i][j].share = futuredata[i][j].volume / Routines.genPopu[idp].population;
                    if (futuredata[i][j].OEM.ToLower() != Routines.optOEM)
                    {
                        Routines.oeshare[i][futuredata[i][j].oeGroupID] += futuredata[i][j].share;
                        Routines.sgshare[i][futuredata[i][j].sgGroupID] += futuredata[i][j].share;
                        Routines.fushare[i][futuredata[i][j].fuGroupID] += futuredata[i][j].share;
                        Routines.tshare[i][futuredata[i][j].tGroupID] += futuredata[i][j].share;
                        Routines.lshare[i][futuredata[i][j].lGroupID] += futuredata[i][j].share;
                        Routines.bshare[i][futuredata[i][j].bGroupID] += futuredata[i][j].share;
                    }

                }
            }
        }



        //prepare optimal vehicle lines data
        public VehicleData[][] preOptData(string oemlist, int timeHorizon, VehicleData[][] genVehData)
        {
            VehicleData[][] optData = new VehicleData[timeHorizon][];

            for (int i = 0; i < timeHorizon; i++)
            {
                int transyear = Routines.beginYear + i;
                int ilength = 0;
                for (int j = 0; j < genVehData[i].Length; j++)
                {
                    if (genVehData[i][j].OEM == oemlist)
                        ilength++;
                }

                optData[i] = new VehicleData[ilength];
                int jid = 0;
                for (int j = 0; j < genVehData[i].Length; j++)
                {

                    if (genVehData[i][j].OEM == oemlist)
                    {

                        optData[i][jid] = new VehicleData();
                        optData[i][jid].group = genVehData[i][j].group;
                        optData[i][jid].type = genVehData[i][j].type;
                        optData[i][jid].fuel_type = genVehData[i][j].fuel_type;
                        optData[i][jid].segment = genVehData[i][j].segment;
                        optData[i][jid].OEM = genVehData[i][j].OEM;
                        optData[i][jid].modelName = genVehData[i][j].modelName;
                        optData[i][jid].modelYear = genVehData[i][j].modelYear;
                        optData[i][jid].transYear = genVehData[i][j].transYear;
                        optData[i][jid].oeGroupID = genVehData[i][j].oeGroupID;
                        optData[i][jid].sgGroupID = genVehData[i][j].sgGroupID;
                        optData[i][jid].fuGroupID = genVehData[i][j].fuGroupID;
                        optData[i][jid].tGroupID = genVehData[i][j].tGroupID;
                        optData[i][jid].lGroupID = genVehData[i][j].lGroupID;
                        optData[i][jid].bGroupID = genVehData[i][j].bGroupID;
                        optData[i][jid].price = genVehData[i][j].price;
                        optData[i][jid].volume = genVehData[i][j].volume;
                        optData[i][jid].share = genVehData[i][j].share;
                        optData[i][jid].delta = getDelta(genVehData[i][j].modelName, genVehData[i][j].modelYear);   //genVehData[i][j].delta;
                        optData[i][jid].profit = genVehData[i][j].profit;
                        optData[i][jid].Elasticity = genVehData[i][j].Elasticity;
                        optData[i][jid].DelsubP = genVehData[i][j].DelsubP;


                        jid++;

                    }

                }
                // Routines += ilength; 
            }


            return optData;
        }

        // Calculate selected vehicle lines' price
        public double getDelta(string modName, int MY)
        {
            double delta = 0;
            string key = modName;
            int idd = Routines.listBDelta.IndexOf(key);
            string keyR = modName + MY;
            int idR = Routines.listRefresh.IndexOf(keyR);


            if (idd > -1)
            {
                if (Routines.baseDelta[idd].modelYear < MY)
                {
                    if (idR > -1)
                        delta = getDelta(modName, MY - 1) + Routines.baseDelta[idd].styleAgeDep + Routines.baseDelta[idd].majImpact + Routines.refresh[idR].successRate * Routines.baseDelta[idd].majStd;
                    else
                        delta = getDelta(modName, MY - 1) + Routines.baseDelta[idd].styleAgeDep;
                }
                else
                    delta = Routines.baseDelta[idd].delta0;
            }
            else
                delta = 0.0;

            return delta;
        }

        private void CaloptPrice(VehicleData[][] optvehicle)//, ArrayList listopt)
        {

            for (int i = 0; i < optvehicle.Length; i++)
            {
                double[] oeshare = new double[Routines.oeGroup.Length];
                double[] sgshare = new double[Routines.sgGroup.Length];
                double[] fushare = new double[Routines.fuGroup.Length];
                double[] tshare = new double[Routines.tGroup.Length];
                double[] lshare = new double[Routines.lGroup.Length];
                double[] bshare = new double[Routines.bGroup.Length];
                double szero = 1;
                int indTY = optvehicle[i][0].transYear - Routines.beginYear;
                //int indTY_macro = optvehicle[i][0].transYear - Routines.baseYear;

                Array.Copy(Routines.oeshare[indTY], 0, oeshare, 0, oeshare.Length);
                Array.Copy(Routines.sgshare[indTY], 0, sgshare, 0, sgshare.Length);
                Array.Copy(Routines.fushare[indTY], 0, fushare, 0, fushare.Length);
                Array.Copy(Routines.tshare[indTY], 0, tshare, 0, tshare.Length);
                Array.Copy(Routines.lshare[indTY], 0, lshare, 0, lshare.Length);
                Array.Copy(Routines.bshare[indTY], 0, bshare, 0, bshare.Length);

                for (int j = 0; j < optvehicle[i].Length; j++)
                {


                    int oeId = optvehicle[i][j].oeGroupID;
                    int sgId = optvehicle[i][j].sgGroupID;
                    int fuId = optvehicle[i][j].fuGroupID;
                    int cId = optvehicle[i][j].tGroupID;
                    int lId = optvehicle[i][j].lGroupID;
                    int bId = optvehicle[i][j].bGroupID;
                    int Pkey = optvehicle[i][j].transYear;
                    int idp = Routines.listPOPU.IndexOf(Pkey);
                    double volume = optvehicle[i][j].volume;
                    optvehicle[i][j].share = volume / Routines.genPopu[idp].population;

                    oeshare[oeId] += optvehicle[i][j].share;
                    sgshare[sgId] += optvehicle[i][j].share;
                    fushare[fuId] += optvehicle[i][j].share;
                    tshare[cId] += optvehicle[i][j].share;
                    lshare[lId] += optvehicle[i][j].share;
                    bshare[bId] += optvehicle[i][j].share;

                }

                for (int j = 0; j < bshare.Length; j++)
                    szero -= bshare[j];




                for (int j = 0; j < optvehicle[i].Length; j++)
                {

                    string key = optvehicle[i][j].modelName;
                    int idx = Routines.listBDelta.IndexOf(key);
                    if (idx != -1)
                    {

                        int oeId = optvehicle[i][j].oeGroupID;
                        int sgId = optvehicle[i][j].sgGroupID;
                        int fuId = optvehicle[i][j].fuGroupID;
                        int tId = optvehicle[i][j].tGroupID;
                        int lId = optvehicle[i][j].lGroupID;
                        int bId = optvehicle[i][j].bGroupID;
                        double deltaj = optvehicle[i][j].delta;// getDelta(optvehicle[i][j].modelName, optvehicle[i][j].modelYear);//optvehicle[i][j].DelsubP;
                        double lambda = Routines.oeGroup[oeId].lambda;
                        double tau = Routines.sgGroup[sgId].tau;
                        double theta = Routines.fuGroup[fuId].theta;
                        double phi = Routines.tGroup[tId].phi;
                        double rho = Routines.lGroup[lId].rho;
                        double sigma = Routines.bGroup[bId].sigma;
                        double alpha = Routines.sgGroup[sgId].alpha * Routines.baseDelta[idx].adjFactor;
                        double ddelta = Routines.baseDelta[idx].ddelta;
                        double sj = optvehicle[i][j].share;

                        double Soe = oeshare[oeId];
                        double Ssg = sgshare[sgId];
                        double Sfu = fushare[fuId];
                        double Sc = tshare[tId];
                        double Slu = lshare[lId];
                        double Sb = bshare[bId];
                        double price = -(deltaj - lambda * Math.Log(sj + ddelta) - (tau - lambda) * Math.Log(Soe) - (theta - tau) * Math.Log(Ssg)
                                    - (phi - theta) * Math.Log(Sfu) - (rho - phi) * Math.Log(Sc)
                                    - (sigma - rho) * Math.Log(Slu) - (1 - sigma) * Math.Log(Sb) + Math.Log(szero)) / alpha;

                        double price1 = -(deltaj - lambda * Math.Log(sj * (1.0 + 1.0e-6) + ddelta) - (tau - lambda) * Math.Log(Soe + sj * 1.0e-6) - (theta - tau) * Math.Log(Ssg + sj * 1.0e-6)
                              - (phi - theta) * Math.Log(Sfu + sj * 1.0e-6) - (rho - phi) * Math.Log(Sc + sj * 1.0e-6)
                              - (sigma - rho) * Math.Log(Slu + sj * 1.0e-6) - (1 - sigma) * Math.Log(Sb + sj * 1.0e-6) + Math.Log(szero - sj * 1.0e-6)) / alpha;

                        optvehicle[i][j].price = price;
                        optvehicle[i][j].Elasticity = 1.0e-6 / (price1 / price - 1);//Routines.baseDelta[idx].PriceEla;



                    }
                }



            }


        }


        public double TestCalProfit(VehicleData[][] optData, Decision[] Decision, ArrayList listDecision)
        {

            getfuData(optData, Decision, listDecision);

            double TotalProfit = 0.0;
            CalannuProfit(optData);
            double[] profit = CalAnnToPro(optData);
            TotalProfit = CalTotalProfit(profit);
            return TotalProfit;
        }

        public double[] CalAnnToPro(VehicleData[][] optData)
        {
            double[] profit = new double[optData.Length];
            for (int i = 0; i < optData.Length; i++)
            {
                for (int j = 0; j < optData[i].Length; j++)
                {
                    // if (optData[i][j].modelName.ToLower() == "ford focus")
                    profit[i] += optData[i][j].profit;

                }
            }
            return profit;
        }

        public double CalTotalProfit(double[] profit)
        {
            double Total = 0.0;
            for (int i = 0; i < profit.Length; i++)
                Total += profit[i] * Math.Pow(Routines.omiga, i);
            return Total;
        }

        public void CalannuProfit(VehicleData[][] optData)
        {
            CaloptPrice(optData);
            for (int i = 0; i < Routines.timeHorizon; i++)
            {
                for (int j = 0; j < optData[i].Length; j++)
                {
                    string searchkeyONTY = optData[i][j].modelName + Convert.ToString(optData[i][j].transYear);
                    string searchkeyON = optData[i][j].modelName;
                    int idxONTY = Routines.listNewFCons.IndexOf(searchkeyONTY);
                    double volume = optData[i][j].volume;
                    double Price = optData[i][j].price;
                    double cost = Routines.genFCons[idxONTY].variableCost;

                    if (volume >= 0)
                        optData[i][j].profit = volume * (Price - cost);
                    else
                        optData[i][j].profit = 0.0;

                }
            }

        }

        public Constraint[] getOptConstraint()
        {

            int n = 0;
            for (int i = 0; i < Routines.optData.Length; i++)
            {
                n += Routines.optData[i].Length;
            }
            Constraint[] optCons = new Constraint[n];
            int conid = 0;
            for (int i = 0; i < Routines.optData.Length; i++)
            {
                for (int j = 0; j < Routines.optData[i].Length; j++)
                {

                    string oem = Routines.optData[i][j].OEM;
                    string modname = Routines.optData[i][j].modelName;
                    int year = Routines.optData[i][j].transYear;
                    string keyCon = modname + Routines.baseYear;
                    string keyConCur = modname + year;

                    int idxBase = Routines.listFConstraints.IndexOf(keyCon);
                    int idxC = Routines.listFConstraints.IndexOf(keyConCur);

                    if (idxC == -1)
                    {
                        if (idxBase > -1) idxC = idxBase;
                        else
                        {
                            MessageBox.Show("The constraints are not completed");
                            return Routines.Constraints;
                        }

                    }
                    optCons[conid] = new Constraint();
                    optCons[conid].OEM = oem;
                    optCons[conid].ModelName = modname;
                    optCons[conid].ModelYear = year;
                    optCons[conid].productionMin = Routines.Constraints[idxC].productionMin;
                    optCons[conid].productionMax = Routines.Constraints[idxC].productionMax;
                    optCons[conid].variableCost = Routines.Constraints[idxC].variableCost;
                    optCons[conid].footprint = Routines.Constraints[idxC].footprint;
                    optCons[conid].mpg = Routines.Constraints[idxC].mpg;
                    optCons[conid].target = Routines.Constraints[idxC].target;
                    optCons[conid].category = Routines.Constraints[idxC].category;
                    optCons[conid].engType = Routines.Constraints[idxC].engType;
                    optCons[conid].VMT = Routines.Constraints[idxC].VMT;




                    conid++;


                }

            }
            return optCons;
        }



        public ArrayList getDecIndexlist(Decision[] decision)
        {
            ArrayList listDecision = new ArrayList();
            for (int i = 0; i < decision.Length; i++)
            {
                string key = decision[i].modelName + Convert.ToString(decision[i].transYear);
                listDecision.Add(key);
            }
            return listDecision;
        }

        public ArrayList[] getOptIndexlist(VehicleData[][] optdata)
        {
            ArrayList[] listopt = new ArrayList[optdata.Length];
            for (int j = 0; j < optdata.Length; j++)
            {
                listopt[j] = new ArrayList();
                for (int i = 0; i < optdata[j].Length; i++)
                {
                    string key = optdata[j][i].modelName;
                    listopt[j].Add(key);
                }
            }
            return listopt;
        }

        /*        public double getDelta(string oem, string modName, int MY, string tranType)
                {
                    double delta = 0;
                    string key = oem + modName + tranType;
                    int idd = Routines.listBDelta.IndexOf(key);
                    string keyR = oem + modName + MY;
                    int idR = Routines.listRefresh.IndexOf(keyR);

                    if (MY > Routines.baseYear + 5)
                    {
                        //MessageBox.Show("model year");
                        int tmp = (MY - Routines.baseYear) % 5;
                        if (tmp == 0) tmp = 5;
                        MY = Routines.baseYear + tmp;
                    }

                    if (idd > -1)
                    {
                        if (Routines.baseDelta[idd].modelYear < MY)
                        {
                            if (idR > -1)
                                delta = getDelta(oem, modName, MY - 1, tranType) + Routines.baseDelta[idd].styleAgeDep + Routines.baseDelta[idd].majImpact + Routines.refresh[idR].successRate * Routines.baseDelta[idd].majStd;
                            else
                                delta = getDelta(oem, modName, MY - 1, tranType) + Routines.baseDelta[idd].styleAgeDep;
                        }
                        else
                            delta = Routines.baseDelta[idd].delta0;
                    }
                    else
                        delta = 0.0;

                    return delta;
                }
                */
        private void getfuData(VehicleData[][] futureData, Decision[] Decision, ArrayList listDecision)
        {
            for (int i = 0; i < futureData.Length; i++)
            {
                for (int j = 0; j < futureData[i].Length; j++)
                {   //add volume decision on future data

                    string key = futureData[i][j].modelName + Convert.ToString(futureData[i][j].transYear);
                    int idx = listDecision.IndexOf(key);
                    if (idx != -1)
                    {
                        futureData[i][j].volume = Decision[idx].Volume;

                    }
                    else
                    {
                        MessageBox.Show("Something is wrong! Decision doesn't include selected vehicle");
                        futureData[i][j].volume = 0.0;
                    }


                }

            }
            // return futureData;
        }




        public Macro[] getPopu(Macro[] Pop)
        {
            int lenPop = Routines.population.Length;
            int lastYear = Routines.population[lenPop - 1].transYear;
            if (lastYear >= Routines.beginYear + Routines.timeHorizon)
                return Pop;
            int lenNew = Routines.beginYear + Routines.timeHorizon - 1 - lastYear;
            Macro[] genPop = new Macro[lenNew + lenPop];
            Array.Copy(Pop, genPop, lenPop);
            Routines.listPOPU = new ArrayList();
            for (int i = 0; i < Routines.population.Length; i++)
                Routines.listPOPU.Add(Routines.population[i].transYear);

            for (int i = 0; i < lenNew; i++)
            {
                int key = lastYear + i + 1;
                int idx = Routines.listPOPU.IndexOf(key);
                if (idx == -1)
                {
                    genPop[lenPop + i] = new Macro();
                    genPop[lenPop + i].transYear = key;
                    genPop[lenPop + i].population = Pop[lenPop - 1].population;

                    Routines.listPOPU.Add(key);
                }
            }
            return genPop;
        }

        public VehicleData[] CopyVehData(VehicleData[] orignal)
        {
            VehicleData[] objective = new VehicleData[orignal.Length];
            for (int i = 0; i < orignal.Length; i++)
            {
                objective[i] = new VehicleData();
                objective[i].group = orignal[i].group;
                objective[i].type = orignal[i].type;
                objective[i].fuel_type = orignal[i].fuel_type;
                objective[i].segment = orignal[i].segment;
                objective[i].OEM = orignal[i].OEM;
                objective[i].modelName = orignal[i].modelName;
                objective[i].modelYear = orignal[i].modelYear;
                objective[i].transYear = orignal[i].transYear;
                objective[i].oeGroupID = orignal[i].oeGroupID;
                objective[i].sgGroupID = orignal[i].sgGroupID;
                objective[i].fuGroupID = orignal[i].fuGroupID;
                objective[i].tGroupID = orignal[i].tGroupID;
                objective[i].lGroupID = orignal[i].lGroupID;
                objective[i].bGroupID = orignal[i].bGroupID;
                objective[i].price = orignal[i].price;
                objective[i].volume = orignal[i].volume;
                objective[i].share = orignal[i].share;
                objective[i].delta = orignal[i].delta;
                objective[i].profit = orignal[i].profit;
                objective[i].Elasticity = orignal[i].Elasticity;
                objective[i].DelsubP = orignal[i].DelsubP;
            }
            return objective;
        }


        public VehicleData[][] preOutData(VehicleData[][] optdata, Decision[] decision)
        {
            VehicleData[][] result = new VehicleData[optdata.Length][];
            for (int i = 0; i < optdata.Length; i++)
            {
                result[i] = new VehicleData[optdata[i].Length];
                //result[i] = new VehicleData[optdata[i].Length + 2 * Routines.numSelFordVeh];
                // result[i] = new VehicleData[optdata[i].Length + 2 * GetNumSelFordVehicle()];
                for (int j = 0; j < optdata[i].Length; j++)
                {
                    result[i][j] = new VehicleData();
                    result[i][j].OEM = optdata[i][j].OEM;
                    result[i][j].segment = optdata[i][j].segment;
                    result[i][j].modelName = optdata[i][j].modelName;
                    result[i][j].modelYear = optdata[i][j].modelYear;
                    result[i][j].transYear = optdata[i][j].transYear;
                    result[i][j].type = optdata[i][j].type;
                    result[i][j].oeGroupID = optdata[i][j].oeGroupID;
                    result[i][j].sgGroupID = optdata[i][j].sgGroupID;
                    result[i][j].fuGroupID = optdata[i][j].fuGroupID;
                    result[i][j].tGroupID = optdata[i][j].tGroupID;
                    result[i][j].lGroupID = optdata[i][j].lGroupID;
                    result[i][j].bGroupID = optdata[i][j].bGroupID;
                    result[i][j].price = optdata[i][j].price;
                    result[i][j].volume = optdata[i][j].volume;
                    result[i][j].share = optdata[i][j].share;
                    result[i][j].delta = optdata[i][j].delta;
                    result[i][j].profit = optdata[i][j].profit;

                    result[i][j].Elasticity = optdata[i][j].Elasticity;

                }

            }
            return result;
        }

        public VehicleData[][] preOutData(VehicleData[][] optdata)
        {
            VehicleData[][] result = new VehicleData[optdata.Length][];
            for (int i = 0; i < optdata.Length; i++)
            {
                result[i] = new VehicleData[optdata[i].Length];

                for (int j = 0; j < optdata[i].Length; j++)
                {
                    result[i][j] = new VehicleData();
                    result[i][j].group = optdata[i][j].group;
                    result[i][j].type = optdata[i][j].type;
                    result[i][j].fuel_type = optdata[i][j].fuel_type;
                    result[i][j].segment = optdata[i][j].segment;
                    result[i][j].OEM = optdata[i][j].OEM;
                    result[i][j].modelName = optdata[i][j].modelName;
                    result[i][j].modelYear = optdata[i][j].modelYear;
                    result[i][j].transYear = optdata[i][j].transYear;
                    result[i][j].type = optdata[i][j].type;
                    result[i][j].oeGroupID = optdata[i][j].oeGroupID;
                    result[i][j].sgGroupID = optdata[i][j].sgGroupID;
                    result[i][j].fuGroupID = optdata[i][j].fuGroupID;
                    result[i][j].tGroupID = optdata[i][j].tGroupID;
                    result[i][j].lGroupID = optdata[i][j].lGroupID;
                    result[i][j].bGroupID = optdata[i][j].bGroupID;
                    result[i][j].price = optdata[i][j].price * 10000;
                    result[i][j].volume = optdata[i][j].volume;
                    result[i][j].share = optdata[i][j].share;
                    result[i][j].delta = optdata[i][j].delta;
                    result[i][j].profit = optdata[i][j].profit;
                    result[i][j].Elasticity = optdata[i][j].Elasticity;
                    // MessageBox.Show("modelname=" + result[i][j].modelName + "   volume=" + optdata[i][j].volume.ToString());
                }

            }
            return result;
        }

        public string PrntDec(Decision[] DecisionTC)
        {
            int len = DecisionTC.Length;// Routines.numSeloptVeh;

            string strtmp = "";
            string stringText = "";
            string formatstd = "{0,12}";
            string formatint = "0,000";
            string formatexp = "0.00e0";

            stringText += "\r\n" + String.Format("{0,-30}", "VehicleName") + String.Format("{0,6}", "Year");
            stringText += String.Format(formatstd, "RetaiVolume");
            stringText += "\r\n";

            for (int id1 = 0; id1 < DecisionTC.Length; id1++)
            {
                stringText += String.Format("{0,-30}", DecisionTC[id1].modelName.ToUpper());
                stringText += String.Format("{0,6}", DecisionTC[id1].transYear);

                if (DecisionTC[id1].Volume > 1.0e6)
                    strtmp = DecisionTC[id1].Volume.ToString(formatexp);
                else
                    strtmp = DecisionTC[id1].Volume.ToString(formatint);


                stringText += String.Format(formatstd, strtmp) + "\r\n";


            }
            return stringText;
        }


        /*     public void ModifyDelta(VehicleData[] histData, BaseDelta[] baseDelta)
             {
                 int bsLength = 0;
                 for (int i = 0; i < histData.Length; i++)
                 {
                     if (histData[i].transYear == Routines.baseYear && histData[i].price > 0)
                         bsLength++;
                 }

                 VehicleData[] baseData = new VehicleData[bsLength];

                 int k = 0;
                 for (int i = 0; i < histData.Length; i++)
                 {
                     if (histData[i].transYear == Routines.baseYear  && histData[i].price > 0)
                     {

                         baseData[k] = new VehicleData();
                         baseData[k].group = histData[i].group;
                         baseData[k].type = histData[i].type;
                         baseData[k].fuel_type = histData[i].fuel_type;
                         baseData[k].segment = histData[i].segment;
                         baseData[k].OEM = histData[i].OEM;
                         baseData[k].modelName = histData[i].modelName;
                         baseData[k].modelYear = histData[i].modelYear;
                         baseData[k].transYear = histData[i].transYear;


                         baseData[k].bGroupID = histData[i].bGroupID;
                         baseData[k].lGroupID = histData[i].lGroupID;
                         baseData[k].tGroupID = histData[i].tGroupID;
                         baseData[k].fuGroupID = histData[i].fuGroupID;
                         baseData[k].sgGroupID = histData[i].sgGroupID;
                         baseData[k].oeGroupID = histData[i].oeGroupID;

                         baseData[k].price = histData[i].price;
                         baseData[k].volume = histData[i].volume;


                         k++;
                     }
                 }// generate base year data



                 double[] oeGshare = new double[Routines.oeGroup.Length];
                 double[] sgGshare = new double[Routines.sgGroup.Length];
                 double[] fuGshare = new double[Routines.fuGroup.Length];
                 double[] cGshare = new double[Routines.tGroup.Length];
                 double[] lGshare = new double[Routines.lGroup.Length];
                 double[] bGshare = new double[Routines.bGroup.Length];
                 double szero = 1;

                 for (int i = 0; i < baseData.Length; i++)
                 {

                     int oeId = baseData[i].oeGroupID;
                     int sgId = baseData[i].sgGroupID;
                     int fuId = baseData[i].fuGroupID;
                     int tId = baseData[i].tGroupID;
                     int lId = baseData[i].lGroupID;
                     int bId = baseData[i].bGroupID;
                     int Pkey = baseData[i].transYear;
                     int idp = Routines.listPOPU.IndexOf(Pkey);

                     baseData[i].share = baseData[i].volume / Routines.genPopu[idp].population;

                     oeGshare[oeId] += baseData[i].share;
                     sgGshare[sgId] += baseData[i].share;
                     fuGshare[fuId] += baseData[i].share;
                     cGshare[tId] += baseData[i].share;
                     lGshare[lId] += baseData[i].share;
                     bGshare[bId] += baseData[i].share;

                 }
                 for (int j = 0; j < bGshare.Length; j++)
                     szero -= bGshare[j];

                 for (int i = 0; i < baseData.Length; i++)
                 {

                     string key = baseData[i].OEM + baseData[i].modelName ;
                     int idx = Routines.listBDelta.IndexOf(key);
                     if (idx != -1)
                     {
                         int oeId = baseData[i].oeGroupID;
                         int sgId = baseData[i].sgGroupID;
                         int fuId = baseData[i].fuGroupID;
                         int tId = baseData[i].tGroupID;
                         int lId = baseData[i].lGroupID;
                         int bId = baseData[i].bGroupID;

                         double sj = baseData[i].share;
                         double price0 = baseData[i].price;
                         double lambda = Routines.oeGroup[oeId].lambda;
                         double tau = Routines.sgGroup[sgId].tau;
                         double theta = Routines.fuGroup[fuId].theta;
                         double phi = Routines.tGroup[tId].phi;
                         double rho = Routines.lGroup[lId].rho;
                         double sigma = Routines.bGroup[bId].sigma;
                         double alpha = Routines.sgGroup[sgId].alpha;

                         double Soe = oeGshare[oeId];
                         double Ssg = sgGshare[sgId];
                         double Sfu = fuGshare[fuId];
                         double Sc = cGshare[tId];
                         double Slu = lGshare[lId];
                         double Sb = bGshare[bId];





                         double tmp1 = sj / Soe;
                         double tmp2 = sj / Ssg;
                         double tmp3 = sj / Sfu;
                         double tmp4 = sj / Sc;
                         double tmp5 = sj / Slu;
                         double tmp6 = sj / Sb;
                         if (tmp1 == 1.0)  tmp1 = 0;
                         if (tmp2 == 1.0)  tmp2 = 0;
                         if (tmp3 == 1.0)  tmp3 = 0.0;
                         if (tmp4 == 1.0)  tmp4= 0.0;
                         double uplIncTime = 1.0;
                         double ddelta = sj / (Math.Exp(-(uplIncTime* price0*alpha-(tau - lambda) * Math.Log(1-tmp1)- (theta - tau) * Math.Log(1-tmp2)
                              - (phi - theta) * Math.Log(1-tmp3)- (rho - phi) * Math.Log(1-tmp4)
                          - (sigma - rho) * Math.Log(1-tmp5) - (1 - sigma) * Math.Log(1-tmp6) + Math.Log(1+sj/szero))/lambda) - 1);

                         double delta = -(alpha * price0 - lambda * Math.Log(sj + ddelta) - (tau - lambda) * Math.Log(Soe) - (theta - tau) * Math.Log(Ssg)
                                - (phi - theta) * Math.Log(Sfu) - (rho - phi) * Math.Log(Sc)
                                - (sigma - rho) * Math.Log(Slu) - (1.0 - sigma) * Math.Log(Sb) + Math.Log(szero)) ;

                         baseDelta[idx].ddelta = ddelta;
                         baseDelta[idx].delta0 = delta;

                         double price = -(delta - lambda * Math.Log(sj+ddelta) - (tau - lambda) * Math.Log(Soe) - (theta - tau) * Math.Log(Ssg)
                                       - (phi - theta) * Math.Log(Sfu) - (rho - phi) * Math.Log(Sc)
                                       - (sigma - rho) * Math.Log(Slu) - (1 - sigma) * Math.Log(Sb) + Math.Log(szero)) / alpha;
                         //calculate price and Ad elasticities on the base point
                         double price1 = -(delta - lambda * Math.Log(sj * (1.0 + 1.0e-6) + ddelta) - (tau - lambda) * Math.Log(Soe + sj * 1.0e-6) - (theta - tau) * Math.Log(Ssg + sj * 1.0e-6)
                                      - (phi - theta) * Math.Log(Sfu + sj * 1.0e-6) - (rho - phi) * Math.Log(Sc + sj * 1.0e-6)
                                      - (sigma - rho) * Math.Log(Slu + sj * 1.0e-6) - (1 - sigma) * Math.Log(Sb + sj * 1.0e-6) + Math.Log(szero - sj * 1.0e-6)) / alpha;

                        baseDelta[idx].PriceEla = (price1 / price - 1) / 1.0e-6;



                     }



                 }
             }
     */
        public void getDelta_new(VehicleData[] vData, int transyear, BaseDelta[] baseDelta)
        {
            int bsLength = 0;
            for (int i = 0; i < vData.Length; i++)
            {
                if (vData[i].price > 0)
                    bsLength++;
            }

            VehicleData[] baseData = new VehicleData[bsLength];

            int k = 0;
            for (int i = 0; i < vData.Length; i++)
            {
                if (vData[i].price > 0)
                {
                    baseData[k] = new VehicleData();
                    baseData[k].group = vData[i].group;
                    baseData[k].type = vData[i].type;
                    baseData[k].fuel_type = vData[i].fuel_type;
                    baseData[k].segment = vData[i].segment;
                    baseData[k].OEM = vData[i].OEM;
                    baseData[k].modelName = vData[i].modelName;
                    baseData[k].modelYear = vData[i].modelYear;
                    baseData[k].transYear = vData[i].transYear;

                    baseData[k].oeGroupID = vData[i].oeGroupID;
                    baseData[k].sgGroupID = vData[i].sgGroupID;
                    baseData[k].fuGroupID = vData[i].fuGroupID;
                    baseData[k].tGroupID = vData[i].tGroupID;
                    baseData[k].lGroupID = vData[i].lGroupID;
                    baseData[k].bGroupID = vData[i].bGroupID;

                    baseData[k].price = vData[i].price;
                    baseData[k].volume = vData[i].volume;

                    k++;
                }
            }

            double[] oeGshare = new double[Routines.oeGroup.Length];
            double[] sgGshare = new double[Routines.sgGroup.Length];
            double[] fuGshare = new double[Routines.fuGroup.Length];
            double[] cGshare = new double[Routines.tGroup.Length];
            double[] lGshare = new double[Routines.lGroup.Length];
            double[] bGshare = new double[Routines.bGroup.Length];
            double szero = 1;

            for (int i = 0; i < baseData.Length; i++)
            {
                int oeId = baseData[i].oeGroupID;
                int sgId = baseData[i].sgGroupID;
                int fuId = baseData[i].fuGroupID;
                int tId = baseData[i].tGroupID;
                int lId = baseData[i].lGroupID;
                int bId = baseData[i].bGroupID;
                int Pkey = baseData[i].transYear;
                int idp = Routines.listPOPU.IndexOf(Pkey);

                baseData[i].share = baseData[i].volume / Routines.genPopu[idp].population;

                oeGshare[oeId] += baseData[i].share;
                sgGshare[sgId] += baseData[i].share;
                fuGshare[fuId] += baseData[i].share;
                cGshare[tId] += baseData[i].share;
                lGshare[lId] += baseData[i].share;
                bGshare[bId] += baseData[i].share;

            }

            for (int j = 0; j < bGshare.Length; j++)
                szero -= bGshare[j];

            for (int i = 0; i < baseData.Length; i++)
            {


                string key = baseData[i].modelName;
                int idx = Routines.listBDelta.IndexOf(key);
                if (idx != -1 & baseData[i].OEM == Routines.optOEM)
                {

                    int oeId = baseData[i].oeGroupID;
                    int sgId = baseData[i].sgGroupID;
                    int fuId = baseData[i].fuGroupID;
                    int tId = baseData[i].tGroupID;
                    int lId = baseData[i].lGroupID;
                    int bId = baseData[i].bGroupID;

                    double sj = baseData[i].share;
                    double price0 = baseData[i].price;
                    double lambda = Routines.oeGroup[oeId].lambda;
                    double tau = Routines.sgGroup[sgId].tau;
                    double theta = Routines.fuGroup[fuId].theta;
                    double phi = Routines.tGroup[tId].phi;
                    double rho = Routines.lGroup[lId].rho;
                    double sigma = Routines.bGroup[bId].sigma;
                    double alpha = Routines.sgGroup[sgId].alpha * Routines.baseDelta[idx].adjFactor;

                    double Soe = oeGshare[oeId];
                    double Ssg = sgGshare[sgId];
                    double Sfu = fuGshare[fuId];
                    double Sc = cGshare[tId];
                    double Slu = lGshare[lId];
                    double Sb = bGshare[bId];





                    double tmp1 = sj / Soe;
                    double tmp2 = sj / Ssg;
                    double tmp3 = sj / Sfu;
                    double tmp4 = sj / Sc;
                    double tmp5 = sj / Slu;
                    double tmp6 = sj / Sb;
                    if (tmp1 == 1.0) tmp1 = 0;
                    if (tmp2 == 1.0) tmp2 = 0;
                    if (tmp3 == 1.0) tmp3 = 0.0;
                    if (tmp4 == 1.0) tmp4 = 0.0;
                    double uplIncTime = 1.0;
                    double ddelta = sj / (Math.Exp(-(uplIncTime * price0 * alpha - (tau - lambda) * Math.Log(1 - tmp1) - (theta - tau) * Math.Log(1 - tmp2)
                         - (phi - theta) * Math.Log(1 - tmp3) - (rho - phi) * Math.Log(1 - tmp4)
                     - (sigma - rho) * Math.Log(1 - tmp5) - (1 - sigma) * Math.Log(1 - tmp6) + Math.Log(1 + sj / szero)) / lambda) - 1);

                    double delta = -(alpha * price0 - lambda * Math.Log(sj + ddelta) - (tau - lambda) * Math.Log(Soe) - (theta - tau) * Math.Log(Ssg)
                           - (phi - theta) * Math.Log(Sfu) - (rho - phi) * Math.Log(Sc)
                           - (sigma - rho) * Math.Log(Slu) - (1.0 - sigma) * Math.Log(Sb) + Math.Log(szero));


                    double price = -(delta - lambda * Math.Log(sj + ddelta) - (tau - lambda) * Math.Log(Soe) - (theta - tau) * Math.Log(Ssg)
                                  - (phi - theta) * Math.Log(Sfu) - (rho - phi) * Math.Log(Sc)
                                  - (sigma - rho) * Math.Log(Slu) - (1 - sigma) * Math.Log(Sb) + Math.Log(szero)) / alpha;
                    //calculate price and Ad elasticities on the base point
                    double price1 = -(delta - lambda * Math.Log(sj * (1.0 + 1.0e-6) + ddelta) - (tau - lambda) * Math.Log(Soe + sj * 1.0e-6) - (theta - tau) * Math.Log(Ssg + sj * 1.0e-6)
                                 - (phi - theta) * Math.Log(Sfu + sj * 1.0e-6) - (rho - phi) * Math.Log(Sc + sj * 1.0e-6)
                                 - (sigma - rho) * Math.Log(Slu + sj * 1.0e-6) - (1 - sigma) * Math.Log(Sb + sj * 1.0e-6) + Math.Log(szero - sj * 1.0e-6)) / alpha;



                    baseDelta[idx].ddelta = ddelta;
                    baseDelta[idx].delta0 = delta;

                    baseDelta[idx].PriceEla = 1.0e-6 / (price1 / price - 1);
                    baseDelta[idx].modelYear = baseData[i].modelYear;
                }
            }
        }



        // calculate implied cost


        public void getAllDelta(VehicleData[][] scenario)
        {

            for (int t = 0; t < scenario.Length; t++)
            {

                double[] oeGshare = new double[Routines.oeGroup.Length];
                double[] sgGshare = new double[Routines.sgGroup.Length];
                double[] fuGshare = new double[Routines.fuGroup.Length];
                double[] cGshare = new double[Routines.tGroup.Length];
                double[] lGshare = new double[Routines.lGroup.Length];
                double[] bGshare = new double[Routines.bGroup.Length];
                double szero = 1.0;

                for (int i = 0; i < scenario[t].Length; i++)
                {

                    int oeId = scenario[t][i].oeGroupID;
                    int sgId = scenario[t][i].sgGroupID;
                    int fuId = scenario[t][i].fuGroupID;
                    int tId = scenario[t][i].tGroupID;
                    int lId = scenario[t][i].lGroupID;
                    int bId = scenario[t][i].bGroupID;
                    int Pkey = scenario[t][i].transYear;
                    int idp = Routines.listPOPU.IndexOf(Pkey);
                    scenario[t][i].share = scenario[t][i].volume / Routines.genPopu[idp].population;
                    oeGshare[oeId] += scenario[t][i].share;
                    sgGshare[sgId] += scenario[t][i].share;
                    fuGshare[fuId] += scenario[t][i].share;
                    cGshare[tId] += scenario[t][i].share;
                    lGshare[lId] += scenario[t][i].share;
                    bGshare[bId] += scenario[t][i].share;

                }
                for (int j = 0; j < bGshare.Length; j++)
                    szero -= bGshare[j];

                for (int i = 0; i < scenario[t].Length; i++)
                {

                    int indTY = scenario[t][i].transYear - Routines.baseYear - 1;
                    int oeId = scenario[t][i].oeGroupID;
                    int sgId = scenario[t][i].sgGroupID;
                    int fuId = scenario[t][i].fuGroupID;
                    int tId = scenario[t][i].tGroupID;
                    int lId = scenario[t][i].lGroupID;
                    int bId = scenario[t][i].bGroupID;

                    double lambda = Routines.oeGroup[oeId].lambda;
                    double tau = Routines.sgGroup[sgId].tau;
                    double theta = Routines.fuGroup[fuId].theta;
                    double phi = Routines.tGroup[tId].phi;
                    double rho = Routines.lGroup[lId].rho;
                    double sigma = Routines.bGroup[bId].sigma;
                    double alpha = Routines.sgGroup[sgId].alpha;
                    double sj = scenario[t][i].share;
                    double price0 = scenario[t][i].price;
                    double Soe = oeGshare[oeId];
                    double Ssg = sgGshare[sgId];
                    double Sfu = fuGshare[fuId];
                    double Sc = cGshare[tId];
                    double Slu = lGshare[lId];
                    double Sb = bGshare[bId];

                    double delta = -alpha * price0
                      + lambda * Math.Log(sj) + (tau - lambda) * Math.Log(Soe) + (theta - tau) * Math.Log(Ssg)
                      + (phi - theta) * Math.Log(Sfu) + (rho - phi) * Math.Log(Sc)
                      + (sigma - rho) * Math.Log(Slu) + (1.0 - sigma) * Math.Log(Sb) - Math.Log(szero);

                    scenario[t][i].DelsubP = delta;

                }
            }
        }


        public void Price2Share(VehicleData[][] scenario)
        {
            for (int i = 0; i < scenario.Length; i++)
            {

                double total = 1.0;
                double[] oeSum = new double[Routines.oeGroup.Length];
                int[] oeInsg = new int[Routines.oeGroup.Length];
                double[] sgSum = new double[Routines.sgGroup.Length];
                int[] sgInFu = new int[Routines.sgGroup.Length];
                double[] fuSum = new double[Routines.fuGroup.Length];
                int[] fuInType = new int[Routines.fuGroup.Length];
                double[] typeSum = new double[Routines.tGroup.Length];
                int[] cInlux = new int[Routines.tGroup.Length];
                double[] luSum = new double[Routines.lGroup.Length];
                int[] linB = new int[Routines.lGroup.Length];
                double[] bSum = new double[Routines.bGroup.Length];

                for (int j = 0; j < scenario[i].Length; j++)
                {

                    int oeId = scenario[i][j].oeGroupID;
                    oeInsg[oeId] = scenario[i][j].sgGroupID;
                    int sgId = scenario[i][j].sgGroupID;
                    sgInFu[sgId] = scenario[i][j].fuGroupID;
                    int fuId = scenario[i][j].fuGroupID;
                    fuInType[fuId] = scenario[i][j].tGroupID;
                    int cId = scenario[i][j].tGroupID;
                    cInlux[cId] = scenario[i][j].lGroupID;
                    int lId = scenario[i][j].lGroupID;
                    linB[lId] = scenario[i][j].bGroupID;

                    double alpham = Routines.sgGroup[sgId].alpha;
                    scenario[i][j].DelsubP += alpham * scenario[i][j].price;
                    oeSum[oeId] += Math.Exp(scenario[i][j].DelsubP / Routines.oeGroup[oeId].lambda);

                }
                for (int k = 0; k < Routines.oeGroup.Length; k++)
                {
                    int sgId = oeInsg[k];
                    sgSum[sgId] += Math.Pow(oeSum[k], Routines.oeGroup[k].lambda / Routines.sgGroup[sgId].tau);
                }

                for (int l = 0; l < Routines.sgGroup.Length; l++)
                {
                    int fuId = sgInFu[l];
                    fuSum[fuId] += Math.Pow(sgSum[l], Routines.sgGroup[l].tau / Routines.fuGroup[fuId].theta);
                }
                for (int m = 0; m < Routines.fuGroup.Length; m++)
                {
                    int cId = fuInType[m];
                    typeSum[cId] += Math.Pow(fuSum[m], Routines.fuGroup[m].theta / Routines.tGroup[cId].phi);
                }
                for (int n = 0; n < Routines.tGroup.Length; n++)
                {
                    int luId = cInlux[n];
                    luSum[luId] += Math.Pow(typeSum[n], Routines.tGroup[n].phi / Routines.lGroup[luId].rho);
                }

                for (int p = 0; p < Routines.lGroup.Length; p++)
                {
                    int bId = linB[p];
                    bSum[bId] += Math.Pow(luSum[p], Routines.lGroup[p].rho / Routines.bGroup[bId].sigma);
                }

                for (int q = 0; q < Routines.bGroup.Length; q++)
                    total += Math.Pow(bSum[q], Routines.bGroup[q].sigma);

                for (int j = 0; j < scenario[i].Length; j++)
                {

                    int oeId = scenario[i][j].oeGroupID;
                    int sgId = scenario[i][j].sgGroupID;
                    int fuId = scenario[i][j].fuGroupID;
                    int tId = scenario[i][j].tGroupID;
                    int lId = scenario[i][j].lGroupID;
                    int bId = scenario[i][j].bGroupID;
                    double lambda = Routines.oeGroup[oeId].lambda;
                    double alpha = Routines.sgGroup[sgId].alpha;
                    double tau = Routines.sgGroup[sgId].tau;
                    double theta = Routines.fuGroup[fuId].theta;
                    double phi = Routines.tGroup[tId].phi;
                    double rho = Routines.lGroup[lId].rho;
                    double sigma = Routines.bGroup[bId].sigma;
                    double share = Math.Exp(scenario[i][j].DelsubP / lambda) * Math.Pow(oeSum[oeId], lambda / tau - 1.0) * Math.Pow(sgSum[sgId], tau / theta - 1.0)
                                        * Math.Pow(fuSum[fuId], theta / phi - 1.0) * Math.Pow(typeSum[tId], phi / rho - 1.0) * Math.Pow(luSum[lId], rho / sigma - 1.0) * Math.Pow(bSum[bId], sigma - 1.0) / total;
                    int Pkey = scenario[i][j].transYear;
                    int idp = Routines.listPOPU.IndexOf(Pkey);
                    scenario[i][j].share = share;
                    scenario[i][j].volume = share * Routines.genPopu[idp].population;



                }

            }
        }
        public void Sales2Price(VehicleData[] basedata)
        {


            int Pkey = Routines.baseYear;
            int idp = Routines.listPOPU.IndexOf(Pkey);
            double[] oeGshare = new double[Routines.oeGroup.Length];
            double[] sgGshare = new double[Routines.sgGroup.Length];
            double[] fuGshare = new double[Routines.fuGroup.Length];
            double[] cGshare = new double[Routines.tGroup.Length];
            double[] lGshare = new double[Routines.lGroup.Length];
            double[] bGshare = new double[Routines.bGroup.Length];
            double szero = 1.0;
            for (int i = 0; i < basedata.Length; i++)
            {

                int oeId = basedata[i].oeGroupID;
                int sgId = basedata[i].sgGroupID;
                int fuId = basedata[i].fuGroupID;
                int tId = basedata[i].tGroupID;
                int lId = basedata[i].lGroupID;
                int bId = basedata[i].bGroupID;

                double volumej = basedata[i].volume;
                basedata[i].share = volumej / Routines.genPopu[idp].population;
                oeGshare[oeId] += basedata[i].share;
                sgGshare[sgId] += basedata[i].share;
                fuGshare[fuId] += basedata[i].share;
                cGshare[tId] += basedata[i].share;
                lGshare[lId] += basedata[i].share;
                bGshare[bId] += basedata[i].share;

            }
            for (int j = 0; j < bGshare.Length; j++)
                szero -= bGshare[j];


            for (int i = 0; i < basedata.Length; i++)
            {



                int oeId = basedata[i].oeGroupID;
                int sgId = basedata[i].sgGroupID;
                int fuId = basedata[i].fuGroupID;
                int tId = basedata[i].tGroupID;
                int lId = basedata[i].lGroupID;
                int bId = basedata[i].bGroupID;
                double deltaj = basedata[i].DelsubP;
                double lambda = Routines.oeGroup[oeId].lambda;
                double tau = Routines.sgGroup[sgId].tau;
                double theta = Routines.fuGroup[fuId].theta;
                double phi = Routines.tGroup[tId].phi;
                double rho = Routines.lGroup[lId].rho;
                double sigma = Routines.bGroup[bId].sigma;
                double alpha = Routines.sgGroup[sgId].alpha;
                double sj = basedata[i].share;

                double Soe = oeGshare[oeId];
                double Ssg = sgGshare[sgId];
                double Sfu = fuGshare[fuId];
                double Sc = cGshare[tId];
                double Slu = lGshare[lId];
                double Sb = bGshare[bId];
                double price = -(deltaj - lambda * Math.Log(sj) - (tau - lambda) * Math.Log(Soe) - (theta - tau) * Math.Log(Ssg)
                            - (phi - theta) * Math.Log(Sfu) - (rho - phi) * Math.Log(Sc)
                            - (sigma - rho) * Math.Log(Slu) - (1 - sigma) * Math.Log(Sb) + Math.Log(szero)) / alpha;

                double price1 = -(deltaj - lambda * Math.Log(sj * (1.0 + 1.0e-6)) - (tau - lambda) * Math.Log(Soe + sj * 1.0e-6) - (theta - tau) * Math.Log(Ssg + sj * 1.0e-6)
                                - (phi - theta) * Math.Log(Sfu + sj * 1.0e-6) - (rho - phi) * Math.Log(Sc + sj * 1.0e-6)
                                - (sigma - rho) * Math.Log(Slu + sj * 1.0e-6) - (1 - sigma) * Math.Log(Sb + sj * 1.0e-6) + Math.Log(szero - sj * 1.0e-6)) / alpha;


                basedata[i].price = price;
                basedata[i].Elasticity = 1.0e-6 / (price1 / price - 1);

            }
        }
        ///////////////////////////////////////////////////////////////////////////
    }
}
