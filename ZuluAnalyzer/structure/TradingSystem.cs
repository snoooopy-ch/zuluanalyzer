using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ZuluAnalyzer
{
    class TradingSystem
    {
        public string trader { get; set; }
        public string currency { get; set; }
        public string type { get; set; }
        public float high { get; set; }
        public float worst { get; set; }
        public int count { get; set; }
        public float profit { get; set; }
        public float spec { get; set; }
        public float lot { get; set; }
        public float maxLot { get; set; }
        public float maxWorst { get; set; }
    }
}
