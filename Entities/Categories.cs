using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ONIT_Kurs_3.Entities
{
    public class Categories
    {
        public int CategoryID { get; set; }
        public string CategoryName { get; set; }
        public string Description { get; set; }
        public byte[] Picture { get; set; }

        public Categories() { }

        public Categories(int categoryID, string categoryName, string description, byte[] picture)
        {
            CategoryID = categoryID;
            CategoryName = categoryName;
            Description = description;
            Picture = picture;
        }
    }
}
