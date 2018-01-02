using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Summary
{
    class POS
    {
        public string Name { get; set; }
        public int Number { get; set; }
        public string OS { get; set; }
        public string XPI { get; set; }
        public string Contactless { get; set; }

        public POS(string os, string xpi, string contactless)
        {
            this.OS = os;
            this.XPI = xpi;
            this.Contactless = contactless;
        }

        public bool IsUpgraded()
        {
            bool answer = new bool();
            if (this.OS == "RFS30251000" && this.XPI == "5200u15" && this.Contactless == "4-1.16.05A4")
            {
                answer = true;
            }
            else
            {
                answer = false;
            }

            return answer;
        }


    }
}
