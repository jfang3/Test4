using System;
using System.Collections.Generic;
using System.Text;

namespace WindowsFormsApplication2
{
    public class BaseDelta
    {
        public string OEM;
        public string segment;
        public string modelName;
 
        public int modelYear;
        public int oeGroupID;
        public int sgGroupID;
        public int fuGroupID;
        public int cGroupID;
        public int lGroupID;
        public int bGroupID;
        public double delta0;      
        public double ddelta;
        public double styleAgeDep;
        public double majImpact;
        public double majStd;
        public double PriceEla;
        public double impliedCost;
        public double adjFactor;
        

        public BaseDelta()
        {
        }
    }
}
