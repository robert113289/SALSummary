using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Summary
{
    class Store
    {
        public int StoreNumber { get; set; }
        public int NumberOfRegisters { get; set; }
        public int NumberOfUpgradedRegisters { get; set; }
        public List<POS> Registers { get; set; }
        

        public Store() { }


        public Store(string storeNumber, string numberOfRegisters)
        {
            this.StoreNumber = int.Parse(storeNumber);
            this.NumberOfRegisters = int.Parse(numberOfRegisters);
            this.NumberOfUpgradedRegisters = 0;
            this.Registers = new List<POS>();

        }

        public bool IsUpgraded()
        {
            foreach (POS register in Registers)
            {
                int numberOfUpgradedRegisters = 0;
                if (register.IsUpgraded())
                {
                    numberOfUpgradedRegisters++;
                }
                
            }
            this.NumberOfUpgradedRegisters = NumberOfUpgradedRegisters;

            return NumberOfUpgradedRegisters == NumberOfUpgradedRegisters;
            
  
        }
        public override string ToString()
        {
            
        }
    }
}
