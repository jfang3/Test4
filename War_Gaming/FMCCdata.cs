using System;
using System.Collections.Generic;
using System.Text;

namespace WindowsFormsApplication2
{
    //Primary Key = OEM + modelName
    public class FMCCdata
    {
        public string OEM;
        public string modelName;

        public double interceptC;
        public double alpha1;
        public double alpha2;

        public double interceptL;
        public double alpha3;

        public double interceptL2;
        public double alphaL2;
        public double interceptL3;
        public double alphaL3;
        public double interceptL4;
        public double alphaL4;
        public double interceptL5;
        public double alphaL5;

        public double discountRate;
        public double interestRateDiff;

        public FMCCdata()
        {
        }
    }
}
