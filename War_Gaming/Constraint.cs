using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApplication2
{
    //Primary Key = OEM + modelName + transactionYear
    public class Constraint
    {
        public string OEM;
        public string ModelName;
        public int ModelYear;
        //public int transYear;
        public double productionMin;
        public double productionMax;
        public double variableCost;
        public double footprint;
        public double mpg;
        public double target;
        public string category;
        public string engType;
        public double VMT;
      
        public Constraint()
        {
        }
    }
}
