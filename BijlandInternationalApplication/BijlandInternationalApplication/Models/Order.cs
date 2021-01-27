using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace BijlandInternationalApplication.Models
{
    [Serializable]
    public class Order
    {
        #region Properties
        private int _id { get; set; }
        public DateTime date { get; set; }
        public string region { get; set; }
        public string rep { get; set; }
        public string item { get; set; }
        public int units { get; set; }
        public float unitCost { get; set; }
        #endregion

        #region Constructor
        public Order(int id, DateTime date, string region, string rep, string item, int units, float unitCost)
        {
            _id = id;
            this.date = date;
            this.region = region;
            this.rep = rep;
            this.item = item;
            this.units = units;
            this.unitCost = unitCost;
        }

        public int GetId()
        {
            return _id;
        }

        public double GetTotalPrice()
        {
            return Math.Round(Convert.ToDouble(unitCost * units), 2);
        }
        #endregion
    }
}