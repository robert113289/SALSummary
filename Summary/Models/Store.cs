using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Summary
{
    class Store
    {
        public int XcellRowID { get; set; }
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
            int numberOfUpgradedRegisters = 0;
            foreach (POS register in Registers)
            {
                
                if (register.IsUpgraded())
                {
                    numberOfUpgradedRegisters++;
                }
                
            }
            this.NumberOfUpgradedRegisters = numberOfUpgradedRegisters;

            return numberOfUpgradedRegisters == this.NumberOfRegisters;
        }


        public string UpgradeStatus()
        {
            string response = "";
            bool upgradeStatus = this.IsUpgraded();
            if (upgradeStatus == true)
            {
                response = "Success";
            }
            if (upgradeStatus == false)
            {
                int numberOfFailedPosUprades = NumberOfRegisters - NumberOfUpgradedRegisters;
                response = "FAILURE: " + numberOfFailedPosUprades.ToString()+ "/" + NumberOfRegisters + " pos failed the upgrade.";
            }
            return response;
        }


        public override string ToString()
        {
            string upgradeStatus = this.UpgradeStatus();
            string storeNumber = StoreNumber.ToString();
            string numberOfRegisters = NumberOfRegisters.ToString();
            
            string numberOfUpgradedRegisters = NumberOfUpgradedRegisters.ToString();

            string response = string.Format("Store Number: {0} \n Number of registers: {1} \n Number of Upgraded registers {2} \n Upgrade Status= {3}", storeNumber, numberOfRegisters, numberOfUpgradedRegisters, upgradeStatus);

            return response;
        }
    }
}
