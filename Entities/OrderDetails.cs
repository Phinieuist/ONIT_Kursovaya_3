using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ONIT_Kurs_3.Entities
{
    public class OrderDetails
    {
        public int OrderID { get; set; }
        public int ProductID { get; set; }
        public decimal UnitPrice { get; set; }
        //public Int16 Quantity { get; set; }
        //public SqlSingle Discount { get; set; }
        //public string UnitPrice { get; set; }
        public string Quantity { get; set; }
        public string Discount { get; set; }

        public OrderDetails() { }

        //public OrderDetails(int orderID, int productID, decimal unitPrice, Int16 quantity, SqlSingle discount)
        //{
        //    OrderID = orderID;
        //    ProductID = productID;
        //    UnitPrice = unitPrice;
        //    Quantity = quantity;
        //    Discount = discount;
        //}
    }
}
