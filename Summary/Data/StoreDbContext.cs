using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Summary.Data
{
    class StoreDbContext
    {
        public List<Store> TodaysStores { get; set; }

        public StoreDbContext()
        {
            TodaysStores = new List<Store>();
        }
    }
}
