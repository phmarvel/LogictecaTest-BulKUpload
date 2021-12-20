using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace LogictecaTest.Models
{
    [Index(nameof(Part_SKU),IsUnique =true)]
    public class Item
    {
        public int Id { get; set; }
        public string Band { get; set; }
        public string Category_Code { get; set; }
        public string Manufacturer { get; set; }
        public string Part_SKU { get; set; }
        public string Item_Description { get; set; }
        public string List_Price { get; set; }
        public string Minimum_Discount { get; set; }
        public string Discounted_Price { get; set; }
    }
}
