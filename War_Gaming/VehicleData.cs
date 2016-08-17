using System;
using System.Collections.Generic;
using System.Text;

namespace WindowsFormsApplication2
{
    public class VehicleData
    {
        public string group;
        public string type;
        public string fuel_type;
        public string segment;
        public string OEM;
        public string modelName;
        public int modelYear;
        public int transYear;

        public int oeGroupID;//oem ,lambda
        public int sgGroupID;//segment,tau
        public int fuGroupID;//fuel, theta
        public int tGroupID;//Car,Truck SUV, phi
        public int lGroupID;//Luxury-NonLuxury rho
        public int bGroupID;// Buy-NonBuy sigma
        public double price;
        public double volume;
        public double share;
        public double delta;
        public double profit;
        public double Elasticity;
        public double DelsubP;

        public VehicleData()
        {
        }
    }
}
