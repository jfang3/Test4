using System;
using System.Collections.Generic;
using System.Text;

namespace WindowsFormsApplication2
{
    /// annual population
    public class Macro
    {
        public int transYear;
        public double population;
        public double income;
        public double gasPrice;
        public double CPI;

        public double carCafeSTD;
        public double truckCafeSTD;

        public double TotAdBudget;
        public double DealAdBudget;

        //5 Flexible production constraints
        public double cap_pg1;
        public double cap_pg2;
        public double cap_pg3;
        public double cap_pg4;
        public double cap_pg5;
    }
}