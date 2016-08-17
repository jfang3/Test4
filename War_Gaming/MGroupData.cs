using System;

namespace WindowsFormsApplication2
{
	public class MGroupData
	{
        public int mgroupID;
		public double alpha;
		public double rho;
        //public double kOfIncome;
        //public double kOfGasprice;
        public double coeOfIncome;
        public double coeOfGasprice;

        //added on 4/3/2008 for calculating main part of coeff_of_gasPrice impact
        public double coeOfGasprice1;

        public double coeOfGaspriceHEV;     //added on 9/19/2007 for HEV
        // public double[] gasImpact;
       // public double[] incomeImpact;
        public double ar;
        public double arHEV;                //added on 9/21/2007 for HEV
        public double initialGas;
        public double initialGasHEV;        //added on 9/21/2007 for HEV
        public double initialIncome;
        public int sgroupID;

        public double coeNLgas;             //coefficient of nonlinear part of gas impact.  added on 6/4/2008
        public double coeNLgas_HEV;         //coefficient of nonlinear part of gas impact for HEV.  added on 6/4/2008
        public double threshold_NLgasPrice;             //gas price threshold for nonlinear part of gas impact.  added on 6/4/2008
        public double threshold_NLgasPrice_HEV;         //gas price threshold for nonlinear part of gas impact for HEV.  added on 6/4/2008
        public double coeOfGasprice_new;             //new coefficient (gama) of gas impact.  added on 6/4/2008
        public double coeOfGaspriceHEV_new;             //new coefficient (gama) of gas impact.  added on 6/4/2008


		public MGroupData()
		{
           
		}
	}
}
