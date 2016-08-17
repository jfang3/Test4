using System;
using System.Collections.Generic;
using System.Text;

namespace WindowsFormsApplication2
{
    //Primary Key = OEM + modelName + transactionYear
    public class FordConstraint
    {
        public string OEM;
        public string modelName;
        public int transYear;
        public double productionMin;
        public double productionMax;
        public int prodgroup_id;       //add on 7/18/2011
        //public double lease2VMax;				//lease 2 volume max
        //public double lease3VMax;				//lease 3 volume max
        //public double lease4VMax;				//lease 4 volume max
        //public double lease5VMax;				//lease 5 volume max
        public double rentalVMin;				//rental volume min
        public double rentalVMax;				//rental volume max
        public double retailVMax;

        public double rentalPrice;
        public double remarketCostRental;
        //public double remarketCostLease;

        public double rentalElast;
        public double rentalVol0;
        public double fleetElast;
        public double fleetVol0;


        public double variableCost;
        public double DMretail;         // dealer margin
        //public double DMlease2;			
        //public double DMlease3;
        //public double DMlease4;
        //public double DMlease5;
        //public double DMofflease2;
        //public double DMofflease3;
        //public double DMofflease4;
        //public double DMofflease5;
        public double DMoffrental;

        public double fleetPrice;
        public double fleetVMax;
        public double fleetVMin;

        public double rentalriskPrice;
        public double rentalriskVMax;
        public double rentalriskVMin;
        public double rentalriskElast;
        public double rentalriskVol0;

        public double varCostFleet;
        public double varCostRental;
        public double varCostRentRisk;

        public double gasMPG;
        public double fuelTarget;
        public string vehType;

        public double gasMPG2WD;
        public double fuelTarget2WD;
        public double volPercent2wd;

        public FordConstraint()
        {
        }
    }
}
