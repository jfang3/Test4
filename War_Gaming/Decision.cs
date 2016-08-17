using System;
using System.Collections.Generic;
using System.Text;

namespace WindowsFormsApplication2
{
    public class Decision
    {
        //Primary Key = OEM + modelName + transactionYear
        public string OEM;
        public string modelName;
        public int transYear;
        public double Volume;
             

        public Decision()
        {
        }
    }
}
