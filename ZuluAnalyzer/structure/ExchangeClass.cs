using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ZuluAnalyzer
{
    class ExchangeClass
    {

        public class ExchangeTable
        {
            [JsonProperty("ticker")]
            public string ticker { get; set; }

            [JsonProperty("open")]
            public float open { get; set; }
            [JsonProperty("date")]
            public string date { get; set; }
        }

        public class ExchangeRate
        {
            [JsonProperty("forexList")]
            public List<ExchangeTable> forexList { get; set; }
        }
    }
}
